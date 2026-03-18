import { useEffect, useMemo, useRef, useState } from "react";
import type {
    AIDetectConfig,
    AIDetectRunKey,
    AIDetectStageKey,
    FileViewState,
    NamedAIDetectConfig,
    ParsedColumn,
    ParsedRow,
} from "../../types";
import type { MainSection, SettingsSection } from "../types";
import {
    AI_RUN_ALL_KEY,
    AI_STAGE_LABELS,
    AI_STAGE_ORDER,
    DEFAULT_AI_BATCH_CONCURRENCY,
    DEFAULT_AI_CONFIG_NAME,
    INITIAL_AI_BATCH_TASK,
} from "../constants";
import {
    buildFinalVerdictExtraFields,
    buildAIDetectFieldsForRow,
    cloneAIDetectConfig,
    composeAISaveText,
    createDefaultAIDetectConfig,
    formatDuration,
    normalizeAIBatchConcurrency,
    normalizeAIDetectConfigForColumns,
    normalizeLoadedAIDetectConfig,
    normalizeLoadedNamedAIDetectConfigs,
    normalizeNamedAIDetectConfigsForColumns,
    requestAIDetectResult,
} from "../ai-helpers";

type NavigateToSection = (
    section: MainSection,
    settingsSection?: SettingsSection,
    rowId?: string | null,
    options?: { replace?: boolean },
) => void;

type RunAllStageState = {
    startedAt: number | null;
    running: boolean;
};

const createRunAllStageState = (): Record<
    AIDetectStageKey,
    RunAllStageState
> => ({
    precheck: { startedAt: null, running: false },
    context_audit: { startedAt: null, running: false },
    independent_solving: { startedAt: null, running: false },
    final_verdict: { startedAt: null, running: false },
});

export const useAIManager = ({
    activeFile,
    activeFileId,
    selectedRow,
    selectedRowId,
    setErrorMessage,
    navigateToSection,
    updateRowAIResult,
    persistFileState,
    flushPendingAIResults,
    latestFileStateRef,
}: {
    activeFile: FileViewState | null;
    activeFileId: string | null;
    selectedRow: ParsedRow | null;
    selectedRowId: string | null;
    setErrorMessage: (value: string) => void;
    navigateToSection: NavigateToSection;
    updateRowAIResult: (
        fileId: string,
        rowId: string,
        stageKey: AIDetectStageKey,
        resultText: string,
    ) => void;
    persistFileState: (file: FileViewState) => Promise<void>;
    flushPendingAIResults: (fileId: string) => Promise<void>;
    latestFileStateRef: React.MutableRefObject<Record<string, FileViewState>>;
}) => {
    const [isAIStageConfigModalOpen, setIsAIStageConfigModalOpen] =
        useState(false);
    const [isAIProfileModalOpen, setIsAIProfileModalOpen] = useState(false);
    const [aiConfigLoading, setAIConfigLoading] = useState(false);
    const [aiConfigSaving, setAIConfigSaving] = useState(false);
    const [aiConfigList, setAIConfigList] = useState<NamedAIDetectConfig[]>(
        () => {
            const config = createDefaultAIDetectConfig();
            return [
                {
                    name: DEFAULT_AI_CONFIG_NAME,
                    config,
                },
            ];
        },
    );
    const [aiConfig, setAIConfig] = useState<AIDetectConfig>(() =>
        createDefaultAIDetectConfig(),
    );
    const [draftAIConfig, setDraftAIConfig] = useState<AIDetectConfig>(() =>
        createDefaultAIDetectConfig(),
    );
    const [aiConfigFormMessage, setAIConfigFormMessage] = useState("");
    const [isAIDetecting, setIsAIDetecting] = useState(false);
    const [aiThinkingText, setAIThinkingText] = useState("");
    const [aiResultText, setAIResultText] = useState("");
    const [aiResultMessage, setAIResultMessage] = useState("");
    const [activeAIRunKey, setActiveAIRunKey] =
        useState<AIDetectRunKey>("precheck");
    const [aiModalStageKey, setAiModalStageKey] =
        useState<AIDetectStageKey>("precheck");
    const [aiDetectElapsedMs, setAIDetectElapsedMs] = useState(0);
    const [isAIRunModalOpen, setIsAIRunModalOpen] = useState(false);
    const [runAllStageState, setRunAllStageState] = useState<
        Record<AIDetectStageKey, RunAllStageState>
    >(createRunAllStageState);
    const [aiBatchTask, setAIBatchTask] = useState(INITIAL_AI_BATCH_TASK);
    const [aiBatchConcurrency, setAIBatchConcurrency] = useState<number>(
        DEFAULT_AI_BATCH_CONCURRENCY,
    );
    const [rowStreamProgress, setRowStreamProgress] = useState<
        Record<string, Partial<Record<AIDetectStageKey, number>>>
    >({});
    const [rowBatchStatuses, setRowBatchStatuses] = useState<
        Record<string, "success" | "failed">
    >({});

    const aiStreamAbortRef = useRef<AbortController | null>(null);
    const aiBatchAbortRef = useRef<AbortController | null>(null);
    const aiDetectStartedAtRef = useRef<number | null>(null);
    const aiBatchPersistTimerRef = useRef<number | null>(null);
    const rowStreamCharsRef = useRef<
        Record<string, Partial<Record<AIDetectStageKey, number>>>
    >({});
    const rowStreamProgressRef = useRef<
        Record<string, Partial<Record<AIDetectStageKey, number>>>
    >({});
    const rowStreamFlushTimerRef = useRef<number | null>(null);

    useEffect(() => {
        return () => {
            aiStreamAbortRef.current?.abort();
            aiStreamAbortRef.current = null;
            aiBatchAbortRef.current?.abort();
            aiBatchAbortRef.current = null;
            if (rowStreamFlushTimerRef.current !== null) {
                window.clearTimeout(rowStreamFlushTimerRef.current);
                rowStreamFlushTimerRef.current = null;
            }
            if (aiBatchPersistTimerRef.current !== null) {
                window.clearInterval(aiBatchPersistTimerRef.current);
                aiBatchPersistTimerRef.current = null;
            }
            rowStreamCharsRef.current = {};
            rowStreamProgressRef.current = {};
        };
    }, []);

    useEffect(() => {
        if (!isAIDetecting) {
            return;
        }
        const startedAt = aiDetectStartedAtRef.current ?? Date.now();
        aiDetectStartedAtRef.current = startedAt;
        const timerId = window.setInterval(() => {
            setAIDetectElapsedMs(Date.now() - startedAt);
        }, 250);
        return () => {
            window.clearInterval(timerId);
        };
    }, [isAIDetecting]);

    useEffect(() => {
        setAIThinkingText("");
        setAIResultText("");
        setAIResultMessage("");
        aiStreamAbortRef.current?.abort();
        aiStreamAbortRef.current = null;
        aiDetectStartedAtRef.current = null;
        setAIDetectElapsedMs(0);
        setIsAIDetecting(false);
    }, [activeFileId, selectedRowId]);

    useEffect(() => {
        resetRowStreamProgress();
        resetRowBatchStatuses();
    }, [activeFileId]);

    useEffect(() => {
        if (!activeFile) {
            const nextConfig = createDefaultAIDetectConfig();
            setAIConfigList([
                {
                    name: DEFAULT_AI_CONFIG_NAME,
                    config: nextConfig,
                },
            ]);
            setAIConfig(nextConfig);
            setDraftAIConfig(cloneAIDetectConfig(nextConfig));
            setAIConfigFormMessage("");
            setAIThinkingText("");
            setAIResultText("");
            setAIResultMessage("");
            aiDetectStartedAtRef.current = null;
            setAIDetectElapsedMs(0);
            setAIConfigLoading(false);
            setIsAIStageConfigModalOpen(false);
            setIsAIProfileModalOpen(false);
            return;
        }

        let disposed = false;
        const controller = new AbortController();
        setAIConfigLoading(true);

        const loadAIDetectConfig = async () => {
            try {
                const response = await fetch(
                    `/api/ai-config/${encodeURIComponent(activeFile.fileName)}`,
                    { signal: controller.signal },
                );
                if (!response.ok) {
                    throw new Error("加载 AI 配置失败");
                }

                const payload = (await response.json()) as {
                    configs?: unknown;
                    activeConfigName?: unknown;
                    config?: unknown;
                };
                let loadedConfigs = normalizeLoadedNamedAIDetectConfigs(
                    payload.configs,
                );
                if (loadedConfigs.length === 0) {
                    loadedConfigs = [
                        {
                            name: DEFAULT_AI_CONFIG_NAME,
                            config: normalizeLoadedAIDetectConfig(
                                payload.config,
                            ),
                        },
                    ];
                }

                const normalizedConfigs =
                    normalizeNamedAIDetectConfigsForColumns(
                        loadedConfigs,
                        activeFile.columns,
                    );
                const activeConfig = normalizedConfigs[0].config;

                if (disposed) {
                    return;
                }
                setAIConfigList([
                    {
                        name: DEFAULT_AI_CONFIG_NAME,
                        config: activeConfig,
                    },
                ]);
                setAIConfig(activeConfig);
                setDraftAIConfig(cloneAIDetectConfig(activeConfig));
                setAIConfigFormMessage("");
            } catch {
                if (disposed) {
                    return;
                }
                const fallbackConfig = normalizeAIDetectConfigForColumns(
                    createDefaultAIDetectConfig(),
                    activeFile.columns,
                );
                setAIConfigList([
                    {
                        name: DEFAULT_AI_CONFIG_NAME,
                        config: fallbackConfig,
                    },
                ]);
                setAIConfig(fallbackConfig);
                setDraftAIConfig(cloneAIDetectConfig(fallbackConfig));
                setAIConfigFormMessage("");
            } finally {
                if (!disposed) {
                    setAIConfigLoading(false);
                }
            }
        };

        void loadAIDetectConfig();

        return () => {
            disposed = true;
            controller.abort();
        };
    }, [activeFile?.fileId, activeFile?.fileName]);

    useEffect(() => {
        if (!activeFile) {
            return;
        }
        const normalizedConfigs = normalizeNamedAIDetectConfigsForColumns(
            aiConfigList,
            activeFile.columns,
        );
        const nextSelectedConfig =
            normalizedConfigs[0]?.config ??
            normalizeAIDetectConfigForColumns(
                createDefaultAIDetectConfig(),
                activeFile.columns,
            );

        setAIConfigList([
            {
                name: DEFAULT_AI_CONFIG_NAME,
                config: nextSelectedConfig,
            },
        ]);
        setAIConfig(nextSelectedConfig);
        setDraftAIConfig((previous) =>
            normalizeAIDetectConfigForColumns(previous, activeFile.columns),
        );
    }, [activeFile?.fileId, activeFile?.columns]);

    const isAIBatchRunning = aiBatchTask.status === "running";
    const aiBatchProgressPercent =
        aiBatchTask.total > 0
            ? Math.round((aiBatchTask.completed / aiBatchTask.total) * 100)
            : 0;
    const aiDetectElapsedText = formatDuration(aiDetectElapsedMs);
    const aiMergedStreamText = useMemo(
        () => composeAISaveText(aiResultText, aiThinkingText),
        [aiResultText, aiThinkingText],
    );

    const canRunAIDetect =
        Boolean(selectedRow) &&
        !isAIDetecting &&
        !aiConfigLoading &&
        !isAIBatchRunning;
    const runAllTimerText =
        isAIDetecting && activeAIRunKey === AI_RUN_ALL_KEY
            ? aiDetectElapsedText
            : "";

    const startRunAllStage = (stageKey: AIDetectStageKey) => {
        setRunAllStageState((previous) => ({
            ...previous,
            [stageKey]: { startedAt: Date.now(), running: true },
        }));
    };

    const stopRunAllStage = (stageKey: AIDetectStageKey) => {
        setRunAllStageState((previous) => {
            const current = previous[stageKey];
            if (!current?.running && current?.startedAt === null) {
                return previous;
            }
            return {
                ...previous,
                [stageKey]: { startedAt: null, running: false },
            };
        });
    };

    const startRunAllStages = () => {
        const now = Date.now();
        setRunAllStageState({
            precheck: { startedAt: now, running: true },
            context_audit: { startedAt: now, running: true },
            independent_solving: { startedAt: now, running: true },
            final_verdict: { startedAt: null, running: false },
        });
    };

    const resetRunAllStageState = () => {
        setRunAllStageState(createRunAllStageState());
    };

    const runAllStageTimers = useMemo(() => {
        if (!(isAIDetecting && activeAIRunKey === AI_RUN_ALL_KEY)) {
            return {};
        }
        const now = Date.now();
        const timers: Partial<Record<AIDetectStageKey, string>> = {};
        AI_STAGE_ORDER.forEach((stageKey) => {
            const state = runAllStageState[stageKey];
            if (state?.running && state.startedAt) {
                timers[stageKey] = formatDuration(now - state.startedAt);
            }
        });
        return timers;
    }, [isAIDetecting, activeAIRunKey, runAllStageState, aiDetectElapsedMs]);

    const modalStageKey = aiModalStageKey;
    const modalStageConfig = activeFile ? aiConfig.stages[modalStageKey] : null;
    const modalFieldLabels = useMemo(() => {
        if (!activeFile || !selectedRow || !modalStageConfig) {
            return [];
        }
        const fields = buildAIDetectFieldsForRow(
            activeFile.columns,
            selectedRow,
            modalStageConfig.submitFieldKeys,
        );
        const labels = fields.map((field) => field.title);
        if (modalStageKey === "final_verdict") {
            labels.push(
                ...buildFinalVerdictExtraFields(
                    selectedRow.aiResults?.independent_solving ?? "",
                ).map((field) => field.title),
            );
        }
        return Array.from(new Set(labels));
    }, [activeFile, selectedRow, modalStageConfig, modalStageKey]);

    const openAIRunModalForStage = (stageKey: AIDetectStageKey) => {
        if (!selectedRow) {
            return;
        }
        setAiModalStageKey(stageKey);
        if (!(isAIDetecting && activeAIRunKey === AI_RUN_ALL_KEY)) {
            setActiveAIRunKey(stageKey);
        }
        const existingResult = selectedRow.aiResults?.[stageKey] ?? "";
        setAIThinkingText("");
        setAIResultText(existingResult);
        setAIResultMessage("");
        setIsAIRunModalOpen(true);
    };

    const scheduleRowStreamFlush = () => {
        if (rowStreamFlushTimerRef.current !== null) {
            return;
        }
        rowStreamFlushTimerRef.current = window.setTimeout(() => {
            rowStreamFlushTimerRef.current = null;
            setRowStreamProgress({ ...rowStreamProgressRef.current });
        }, 120);
    };

    const stopBatchPersistLoop = () => {
        if (aiBatchPersistTimerRef.current !== null) {
            window.clearInterval(aiBatchPersistTimerRef.current);
            aiBatchPersistTimerRef.current = null;
        }
    };

    const startBatchPersistLoop = (fileId: string) => {
        stopBatchPersistLoop();
        aiBatchPersistTimerRef.current = window.setInterval(() => {
            const latest = latestFileStateRef.current[fileId];
            if (latest) {
                void persistFileState(latest);
            }
        }, 1500);
    };

    const resetRowStreamProgress = () => {
        if (rowStreamFlushTimerRef.current !== null) {
            window.clearTimeout(rowStreamFlushTimerRef.current);
            rowStreamFlushTimerRef.current = null;
        }
        rowStreamCharsRef.current = {};
        rowStreamProgressRef.current = {};
        setRowStreamProgress({});
    };

    const resetRowBatchStatuses = () => {
        setRowBatchStatuses({});
    };

    const markRowBatchStatus = (
        rowId: string,
        status: "success" | "failed",
    ) => {
        setRowBatchStatuses((previous) => {
            if (previous[rowId] === status) {
                return previous;
            }
            return { ...previous, [rowId]: status };
        });
    };

    const setRowStageProgress = (
        rowId: string,
        stageKey: AIDetectStageKey,
        nextProgress: number,
    ) => {
        const rowProgress = rowStreamProgressRef.current[rowId] ?? {};
        const current = rowProgress[stageKey] ?? 0;
        if (nextProgress <= current) {
            return;
        }
        rowProgress[stageKey] = nextProgress;
        rowStreamProgressRef.current[rowId] = rowProgress;
        scheduleRowStreamFlush();
    };

    const primeRowStageProgress = (
        rows: ParsedRow[],
        stageKeys: readonly AIDetectStageKey[],
        value: number = 2,
    ) => {
        if (rows.length === 0) {
            return;
        }
        rows.forEach((row) => {
            stageKeys.forEach((stageKey) => {
                setRowStageProgress(row.rowId, stageKey, value);
            });
        });
    };

    const markRowStageRunning = (rowId: string, stageKey: AIDetectStageKey) => {
        setRowStageProgress(rowId, stageKey, 2);
    };

    const updateRowStageProgress = (
        rowId: string,
        stageKey: AIDetectStageKey,
        chunkLength: number,
    ) => {
        if (chunkLength <= 0) {
            return;
        }
        const rowChars = rowStreamCharsRef.current[rowId] ?? {};
        const nextChars = (rowChars[stageKey] ?? 0) + chunkLength;
        rowChars[stageKey] = nextChars;
        rowStreamCharsRef.current[rowId] = rowChars;
        const computed = Math.min(
            95,
            Math.round(10 + Math.log1p(nextChars) * 12),
        );
        setRowStageProgress(rowId, stageKey, computed);
    };

    const finalizeRowStageProgress = (
        rowId: string,
        stageKey: AIDetectStageKey,
    ) => {
        setRowStageProgress(rowId, stageKey, 100);
    };

    const getStageLabel = (stageKey: AIDetectStageKey) =>
        AI_STAGE_LABELS[stageKey]?.shortTitle ?? stageKey;

    const resolveProfileForStage = (
        config: AIDetectConfig,
        stageConfig: AIDetectConfig["stages"][AIDetectStageKey],
    ) => {
        const profileItem =
            config.profiles.find(
                (item) => item.name === stageConfig.profileName,
            ) ?? config.profiles[0];
        return profileItem?.profile ?? null;
    };

    const validateStageSetup = (
        stageKey: AIDetectStageKey,
        stageConfig: AIDetectConfig["stages"][AIDetectStageKey],
        profile: AIDetectConfig["profiles"][number]["profile"] | null,
        includeStagePrefix: boolean,
    ): string | null => {
        const prefix = includeStagePrefix
            ? `【${getStageLabel(stageKey)}】`
            : "";
        if (!profile) {
            return `${prefix}请先配置接口`;
        }
        if (profile.model.trim().length === 0) {
            return `${prefix}请先配置模型`;
        }
        if (profile.url.trim().length === 0) {
            return `${prefix}请先配置 Base URL`;
        }
        if (profile.apiKey.trim().length === 0) {
            return `${prefix}请先配置 API Key`;
        }
        if (stageConfig.submitFieldKeys.length === 0) {
            return `${prefix}请先在 AI 配置中选择提交回答字段`;
        }
        if (stageConfig.prompt.trim().length === 0) {
            return `${prefix}请先配置 Prompt`;
        }
        return null;
    };

    const buildStageFieldsForRow = (
        columns: ParsedColumn[],
        row: ParsedRow,
        stageKey: AIDetectStageKey,
        stageConfig: AIDetectConfig["stages"][AIDetectStageKey],
        independentSolvingResult?: string,
    ) => {
        const fields = buildAIDetectFieldsForRow(
            columns,
            row,
            stageConfig.submitFieldKeys,
        );
        if (stageKey === "final_verdict") {
            fields.push(
                ...buildFinalVerdictExtraFields(independentSolvingResult ?? ""),
            );
        }
        return fields;
    };

    const runStageRequest = async ({
        stageConfig,
        profile,
        fields,
        signal,
        rowId,
        stageKey,
    }: {
        stageConfig: AIDetectConfig["stages"][AIDetectStageKey];
        profile: AIDetectConfig["profiles"][number]["profile"];
        fields: ReturnType<typeof buildAIDetectFieldsForRow>;
        signal: AbortSignal;
        rowId?: string;
        stageKey?: AIDetectStageKey;
    }): Promise<string> => {
        if (fields.length === 0) {
            throw new Error("没有可提交的回答字段");
        }
        if (rowId && stageKey) {
            markRowStageRunning(rowId, stageKey);
        }
        const streamResult = await requestAIDetectResult(
            {
                provider: profile.provider,
                url: profile.url,
                model: profile.model,
                apiKey: profile.apiKey,
                prompt: stageConfig.prompt,
                fields,
                reasoningEffort: profile.reasoningEffort,
                retryCount: profile.retryCount,
            },
            {
                signal,
                onAnswerChunk: (chunk) => {
                    if (rowId && stageKey) {
                        updateRowStageProgress(rowId, stageKey, chunk.length);
                    }
                },
                onThinkingChunk: (chunk) => {
                    if (rowId && stageKey) {
                        updateRowStageProgress(rowId, stageKey, chunk.length);
                    }
                },
            },
        );
        const text = streamResult.answerText.trim();
        if (text.trim().length === 0) {
            throw new Error("AI 返回为空");
        }
        if (rowId && stageKey && !signal.aborted) {
            finalizeRowStageProgress(rowId, stageKey);
        }
        return text;
    };

    const runBatchAllStages = async ({
        normalizedConfig,
        targetRows,
        targetColumns,
        targetFileId,
        targetFileName,
        normalizedTargetRowIds,
    }: {
        normalizedConfig: AIDetectConfig;
        targetRows: ParsedRow[];
        targetColumns: ParsedColumn[];
        targetFileId: string;
        targetFileName: string;
        normalizedTargetRowIds: string[] | null;
    }) => {
        const stageRunMap = new Map<
            AIDetectStageKey,
            {
                stageConfig: AIDetectConfig["stages"][AIDetectStageKey];
                profile: AIDetectConfig["profiles"][number]["profile"];
            }
        >();
        for (const stageKey of AI_STAGE_ORDER) {
            const stageConfig = normalizedConfig.stages[stageKey];
            const profile = resolveProfileForStage(
                normalizedConfig,
                stageConfig,
            );
            const error = validateStageSetup(
                stageKey,
                stageConfig,
                profile,
                true,
            );
            if (error || !profile) {
                setErrorMessage(error ?? "请先配置接口");
                return;
            }
            stageRunMap.set(stageKey, { stageConfig, profile });
        }

        const precheckConfig = stageRunMap.get("precheck");
        const contextConfig = stageRunMap.get("context_audit");
        const independentConfig = stageRunMap.get("independent_solving");
        const finalConfig = stageRunMap.get("final_verdict");
        if (
            !precheckConfig ||
            !contextConfig ||
            !independentConfig ||
            !finalConfig
        ) {
            setErrorMessage("阶段配置缺失");
            return;
        }

        let nextCursor = 0;
        const requestedConcurrency =
            normalizeAIBatchConcurrency(aiBatchConcurrency);
        const workerCount = Math.min(requestedConcurrency, targetRows.length);

        aiBatchAbortRef.current?.abort();
        const controller = new AbortController();
        aiBatchAbortRef.current = controller;
        startBatchPersistLoop(targetFileId);
        resetRowStreamProgress();
        resetRowBatchStatuses();
        setAIBatchTask({
            status: "running",
            fileId: targetFileId,
            fileName: targetFileName,
            total: targetRows.length,
            completed: 0,
            success: 0,
            failed: 0,
            message:
                normalizedTargetRowIds && normalizedTargetRowIds.length > 0
                    ? `已选择 ${targetRows.length} 条，执行全部，并发 ${workerCount} 线程`
                    : `执行全部，并发 ${workerCount} 线程`,
        });
        setErrorMessage("");
        setAIResultMessage("");
        primeRowStageProgress(targetRows, AI_STAGE_ORDER);

        const runWorker = async () => {
            while (!controller.signal.aborted) {
                const currentIndex = nextCursor;
                nextCursor += 1;
                if (currentIndex >= targetRows.length) {
                    return;
                }

                const row = targetRows[currentIndex];
                let rowFailed = false;
                try {
                    const precheckPromise = runStageRequest({
                        stageConfig: precheckConfig.stageConfig,
                        profile: precheckConfig.profile,
                        fields: buildStageFieldsForRow(
                            targetColumns,
                            row,
                            "precheck",
                            precheckConfig.stageConfig,
                        ),
                        signal: controller.signal,
                        rowId: row.rowId,
                        stageKey: "precheck",
                    });
                    const contextPromise = runStageRequest({
                        stageConfig: contextConfig.stageConfig,
                        profile: contextConfig.profile,
                        fields: buildStageFieldsForRow(
                            targetColumns,
                            row,
                            "context_audit",
                            contextConfig.stageConfig,
                        ),
                        signal: controller.signal,
                        rowId: row.rowId,
                        stageKey: "context_audit",
                    });
                    const independentPromise = runStageRequest({
                        stageConfig: independentConfig.stageConfig,
                        profile: independentConfig.profile,
                        fields: buildStageFieldsForRow(
                            targetColumns,
                            row,
                            "independent_solving",
                            independentConfig.stageConfig,
                        ),
                        signal: controller.signal,
                        rowId: row.rowId,
                        stageKey: "independent_solving",
                    });

                    let independentResult: string | null = null;
                    try {
                        independentResult = await independentPromise;
                        updateRowAIResult(
                            targetFileId,
                            row.rowId,
                            "independent_solving",
                            independentResult,
                        );
                    } catch {
                        rowFailed = true;
                    }

                    if (independentResult && !controller.signal.aborted) {
                        try {
                            const finalText = await runStageRequest({
                                stageConfig: finalConfig.stageConfig,
                                profile: finalConfig.profile,
                                fields: buildStageFieldsForRow(
                                    targetColumns,
                                    row,
                                    "final_verdict",
                                    finalConfig.stageConfig,
                                    independentResult,
                                ),
                                signal: controller.signal,
                                rowId: row.rowId,
                                stageKey: "final_verdict",
                            });
                            updateRowAIResult(
                                targetFileId,
                                row.rowId,
                                "final_verdict",
                                finalText,
                            );
                        } catch {
                            rowFailed = true;
                        }
                    } else {
                        rowFailed = true;
                    }

                    const [precheckSettled, contextSettled] =
                        await Promise.allSettled([
                            precheckPromise,
                            contextPromise,
                        ]);
                    if (precheckSettled.status === "fulfilled") {
                        updateRowAIResult(
                            targetFileId,
                            row.rowId,
                            "precheck",
                            precheckSettled.value,
                        );
                    } else {
                        rowFailed = true;
                    }
                    if (contextSettled.status === "fulfilled") {
                        updateRowAIResult(
                            targetFileId,
                            row.rowId,
                            "context_audit",
                            contextSettled.value,
                        );
                    } else {
                        rowFailed = true;
                    }
                } catch {
                    rowFailed = true;
                }

                if (controller.signal.aborted) {
                    return;
                }
                markRowBatchStatus(row.rowId, rowFailed ? "failed" : "success");
                setAIBatchTask((previous) => ({
                    ...previous,
                    completed: previous.completed + 1,
                    success: previous.success + (rowFailed ? 0 : 1),
                    failed: previous.failed + (rowFailed ? 1 : 0),
                }));
            }
        };

        try {
            await Promise.all(
                Array.from({ length: workerCount }, () => runWorker()),
            );

            if (controller.signal.aborted) {
                return;
            }

            setAIBatchTask((previous) => ({
                ...previous,
                message: "执行全部完成，正在写入 AI 检测结果",
            }));
            await flushPendingAIResults(targetFileId);
            const latestFile = latestFileStateRef.current[targetFileId];
            if (latestFile) {
                await persistFileState(latestFile);
            }
            setAIBatchTask((previous) => ({
                ...previous,
                status: "completed",
                message: "执行全部完成，已写入 AI 检测结果",
            }));
            setErrorMessage("");
        } catch (error) {
            if (controller.signal.aborted) {
                return;
            }
            const message =
                error instanceof Error
                    ? error.message
                    : "批量 AI 回答任务执行失败";
            setAIBatchTask((previous) => ({
                ...previous,
                status: "completed",
                message,
            }));
            setErrorMessage(message);
        } finally {
            if (aiBatchAbortRef.current === controller) {
                aiBatchAbortRef.current = null;
            }
            stopBatchPersistLoop();
            resetRowStreamProgress();
        }
    };

    const syncActiveAIConfigState = (nextConfig: AIDetectConfig) => {
        setAIConfig(nextConfig);
        setAIConfigList([
            {
                name: DEFAULT_AI_CONFIG_NAME,
                config: nextConfig,
            },
        ]);
    };

    const prepareDraftAIConfig = () => {
        if (!activeFile) {
            return false;
        }
        setDraftAIConfig(
            cloneAIDetectConfig(
                normalizeAIDetectConfigForColumns(aiConfig, activeFile.columns),
            ),
        );
        setAIConfigFormMessage("");
        return true;
    };

    const onOpenAIStageConfigModal = () => {
        if (!prepareDraftAIConfig()) {
            return;
        }
        setIsAIStageConfigModalOpen(true);
        setIsAIProfileModalOpen(false);
        navigateToSection("settings", "ai");
    };

    const onOpenAIProfileModal = () => {
        if (!prepareDraftAIConfig()) {
            return;
        }
        setIsAIProfileModalOpen(true);
        setIsAIStageConfigModalOpen(false);
        navigateToSection("settings", "ai");
    };

    const onCancelAIStageConfigModal = () => {
        setDraftAIConfig(cloneAIDetectConfig(aiConfig));
        setAIConfigFormMessage("");
        setIsAIStageConfigModalOpen(false);
    };

    const onCancelAIProfileModal = () => {
        setDraftAIConfig(cloneAIDetectConfig(aiConfig));
        setAIConfigFormMessage("");
        setIsAIProfileModalOpen(false);
    };

    const updateDraftStageConfig = (
        stageKey: AIDetectStageKey,
        updater: (
            stage: AIDetectConfig["stages"][AIDetectStageKey],
        ) => AIDetectConfig["stages"][AIDetectStageKey],
    ) => {
        setDraftAIConfig((previous) => ({
            ...previous,
            stages: {
                ...previous.stages,
                [stageKey]: updater(previous.stages[stageKey]),
            },
        }));
    };

    const onToggleDraftAISubmitField = (
        stageKey: AIDetectStageKey,
        columnKey: string,
    ) => {
        updateDraftStageConfig(stageKey, (stage) => {
            const exists = stage.submitFieldKeys.includes(columnKey);
            const submitFieldKeys = exists
                ? stage.submitFieldKeys.filter((key) => key !== columnKey)
                : [...stage.submitFieldKeys, columnKey];
            return {
                ...stage,
                submitFieldKeys,
            };
        });
    };

    const onSaveAIConfig = async (skipStageValidation = false) => {
        if (!activeFile) {
            return;
        }

        const nextConfigName = DEFAULT_AI_CONFIG_NAME;

        const nextConfig = normalizeAIDetectConfigForColumns(
            draftAIConfig,
            activeFile.columns,
        );

        if (!nextConfig.profiles || nextConfig.profiles.length === 0) {
            setAIConfigFormMessage("请至少配置一个接口");
            return;
        }

        const profileNameSet = new Set<string>();
        for (const profileItem of nextConfig.profiles) {
            const profileName = profileItem.name.trim();
            if (profileName.length === 0) {
                setAIConfigFormMessage("接口配置名称不能为空");
                return;
            }
            if (profileNameSet.has(profileName)) {
                setAIConfigFormMessage(`接口配置名称重复：${profileName}`);
                return;
            }
            profileNameSet.add(profileName);
            const profile = profileItem.profile;
            if (profile.model.trim().length === 0) {
                setAIConfigFormMessage(`【${profileName}】模型不能为空`);
                return;
            }
            if (profile.url.trim().length === 0) {
                setAIConfigFormMessage(`【${profileName}】Base URL 不能为空`);
                return;
            }
            if (profile.apiKey.trim().length === 0) {
                setAIConfigFormMessage(`【${profileName}】API Key 不能为空`);
                return;
            }
        }

        if (!skipStageValidation) {
            for (const stageKey of AI_STAGE_ORDER) {
                const stageConfig = nextConfig.stages[stageKey];
                const stageLabel =
                    AI_STAGE_LABELS[stageKey]?.shortTitle ?? stageKey;

                if (!profileNameSet.has(stageConfig.profileName)) {
                    setAIConfigFormMessage(
                        `【${stageLabel}】请选择有效的接口配置`,
                    );
                    return;
                }
                if (stageConfig.submitFieldKeys.length === 0) {
                    setAIConfigFormMessage(
                        `【${stageLabel}】请至少选择一个提交回答字段`,
                    );
                    return;
                }
                if (stageConfig.prompt.trim().length === 0) {
                    setAIConfigFormMessage(`【${stageLabel}】Prompt 不能为空`);
                    return;
                }
            }
        }

        setAIConfigSaving(true);
        setAIConfigFormMessage("");
        setErrorMessage("");

        try {
            const response = await fetch(
                `/api/ai-config/${encodeURIComponent(activeFile.fileName)}`,
                {
                    method: "PUT",
                    headers: {
                        "Content-Type": "application/json",
                    },
                    body: JSON.stringify({
                        name: nextConfigName,
                        profiles: nextConfig.profiles,
                        stages: nextConfig.stages,
                        setActive: true,
                    }),
                },
            );

            if (!response.ok) {
                const payload = (await response.json().catch(() => ({}))) as {
                    message?: string;
                };
                throw new Error(payload.message ?? "保存 AI 配置失败");
            }

            setAIConfigList((previous) => [
                {
                    name: nextConfigName,
                    config: nextConfig,
                },
                ...previous.filter((item) => item.name !== nextConfigName),
            ]);
            setAIConfig(nextConfig);
            setDraftAIConfig(cloneAIDetectConfig(nextConfig));
            setAIConfigFormMessage("");
            setIsAIStageConfigModalOpen(false);
            setIsAIProfileModalOpen(false);
        } catch (error) {
            const message =
                error instanceof Error ? error.message : "保存 AI 配置失败";
            setAIConfigFormMessage(message);
        } finally {
            setAIConfigSaving(false);
        }
    };

    const onSaveAIProfileConfig = () => onSaveAIConfig(true);
    const onSaveAIStageConfig = () => onSaveAIConfig(false);

    const onRunAIDetect = async (runKeyOverride?: AIDetectRunKey) => {
        if (!activeFile || !selectedRow) {
            return;
        }
        if (isAIBatchRunning) {
            setAIResultMessage("批量 AI 任务运行中，暂不可发起单条回答");
            return;
        }

        const runKey = runKeyOverride ?? activeAIRunKey;
        const normalizedConfig = normalizeAIDetectConfigForColumns(
            aiConfig,
            activeFile.columns,
        );
        syncActiveAIConfigState(normalizedConfig);
        if (runKey === AI_RUN_ALL_KEY) {
            const stageRunMap = new Map<
                AIDetectStageKey,
                {
                    stageConfig: AIDetectConfig["stages"][AIDetectStageKey];
                    profile: AIDetectConfig["profiles"][number]["profile"];
                }
            >();
            for (const stageKey of AI_STAGE_ORDER) {
                const stageConfig = normalizedConfig.stages[stageKey];
                const profile = resolveProfileForStage(
                    normalizedConfig,
                    stageConfig,
                );
                const error = validateStageSetup(
                    stageKey,
                    stageConfig,
                    profile,
                    true,
                );
                if (error || !profile) {
                    setAIResultMessage(error ?? "请先配置接口");
                    return;
                }
                stageRunMap.set(stageKey, { stageConfig, profile });
            }

            aiStreamAbortRef.current?.abort();
            const controller = new AbortController();
            aiStreamAbortRef.current = controller;
            aiDetectStartedAtRef.current = Date.now();
            setAIDetectElapsedMs(0);
            setIsAIDetecting(true);
            setAIThinkingText("");
            setAIResultText("");
            setAIResultMessage("");
            setErrorMessage("");
            startRunAllStages();

            const stageResults: Partial<Record<AIDetectStageKey, string>> = {};
            const stageErrors: string[] = [];
            const formatStageError = (
                stageKey: AIDetectStageKey,
                error: unknown,
            ) => {
                const message =
                    error instanceof Error ? error.message : "AI 回答失败";
                stageErrors.push(`【${getStageLabel(stageKey)}】${message}`);
            };

            try {
                const precheckConfig = stageRunMap.get("precheck");
                const contextConfig = stageRunMap.get("context_audit");
                const independentConfig = stageRunMap.get(
                    "independent_solving",
                );
                const finalConfig = stageRunMap.get("final_verdict");
                if (
                    !precheckConfig ||
                    !contextConfig ||
                    !independentConfig ||
                    !finalConfig
                ) {
                    setAIResultMessage("阶段配置缺失");
                    return;
                }

                const precheckPromise = runStageRequest({
                    stageConfig: precheckConfig.stageConfig,
                    profile: precheckConfig.profile,
                    fields: buildStageFieldsForRow(
                        activeFile.columns,
                        selectedRow,
                        "precheck",
                        precheckConfig.stageConfig,
                    ),
                    signal: controller.signal,
                })
                    .then((text) => {
                        stageResults.precheck = text;
                        updateRowAIResult(
                            activeFile.fileId,
                            selectedRow.rowId,
                            "precheck",
                            text,
                        );
                        return text;
                    })
                    .finally(() => {
                        stopRunAllStage("precheck");
                    });
                const contextPromise = runStageRequest({
                    stageConfig: contextConfig.stageConfig,
                    profile: contextConfig.profile,
                    fields: buildStageFieldsForRow(
                        activeFile.columns,
                        selectedRow,
                        "context_audit",
                        contextConfig.stageConfig,
                    ),
                    signal: controller.signal,
                })
                    .then((text) => {
                        stageResults.context_audit = text;
                        updateRowAIResult(
                            activeFile.fileId,
                            selectedRow.rowId,
                            "context_audit",
                            text,
                        );
                        return text;
                    })
                    .finally(() => {
                        stopRunAllStage("context_audit");
                    });
                const independentPromise = runStageRequest({
                    stageConfig: independentConfig.stageConfig,
                    profile: independentConfig.profile,
                    fields: buildStageFieldsForRow(
                        activeFile.columns,
                        selectedRow,
                        "independent_solving",
                        independentConfig.stageConfig,
                    ),
                    signal: controller.signal,
                })
                    .then((text) => {
                        stageResults.independent_solving = text;
                        updateRowAIResult(
                            activeFile.fileId,
                            selectedRow.rowId,
                            "independent_solving",
                            text,
                        );
                        return text;
                    })
                    .finally(() => {
                        stopRunAllStage("independent_solving");
                    });

                let independentResult: string | null = null;
                try {
                    independentResult = await independentPromise;
                } catch (error) {
                    formatStageError("independent_solving", error);
                }

                if (independentResult && !controller.signal.aborted) {
                    try {
                        startRunAllStage("final_verdict");
                        const finalText = await runStageRequest({
                            stageConfig: finalConfig.stageConfig,
                            profile: finalConfig.profile,
                            fields: buildStageFieldsForRow(
                                activeFile.columns,
                                selectedRow,
                                "final_verdict",
                                finalConfig.stageConfig,
                                independentResult,
                            ),
                            signal: controller.signal,
                        });
                        stageResults.final_verdict = finalText;
                        updateRowAIResult(
                            activeFile.fileId,
                            selectedRow.rowId,
                            "final_verdict",
                            finalText,
                        );
                    } catch (error) {
                        formatStageError("final_verdict", error);
                    } finally {
                        stopRunAllStage("final_verdict");
                    }
                } else {
                    stageErrors.push(
                        `【${getStageLabel("final_verdict")}】缺少第三阶段结果，无法执行`,
                    );
                }

                const [precheckSettled, contextSettled] =
                    await Promise.allSettled([precheckPromise, contextPromise]);
                if (precheckSettled.status !== "fulfilled") {
                    formatStageError("precheck", precheckSettled.reason);
                }
                if (contextSettled.status !== "fulfilled") {
                    formatStageError("context_audit", contextSettled.reason);
                }

                if (controller.signal.aborted) {
                    setAIResultMessage("AI 回答已取消");
                    return;
                }

                const combinedText = AI_STAGE_ORDER.filter(
                    (stageKey) => stageResults[stageKey],
                )
                    .map(
                        (stageKey) =>
                            `【${getStageLabel(stageKey)}】\n${stageResults[stageKey]}`,
                    )
                    .join("\n\n");
                if (combinedText.trim().length > 0) {
                    setAIResultText(combinedText);
                }
                setAIResultMessage(
                    stageErrors.length > 0
                        ? `部分阶段失败：${stageErrors.join("；")}`
                        : "AI 回答完成，已写入 AI 检测结果",
                );
            } catch (error) {
                if (controller.signal.aborted) {
                    setAIResultMessage("AI 回答已取消");
                } else {
                    const message =
                        error instanceof Error ? error.message : "AI 回答失败";
                    setAIResultMessage(message);
                }
            } finally {
                if (aiStreamAbortRef.current === controller) {
                    aiStreamAbortRef.current = null;
                }
                if (aiDetectStartedAtRef.current) {
                    setAIDetectElapsedMs(
                        Date.now() - aiDetectStartedAtRef.current,
                    );
                    aiDetectStartedAtRef.current = null;
                }
                setIsAIDetecting(false);
                resetRunAllStageState();
            }
            return;
        }

        const stageKey = runKey;
        const stageConfig = normalizedConfig.stages[stageKey];
        const stageLabel = AI_STAGE_LABELS[stageKey]?.shortTitle ?? "";
        const profileItem =
            normalizedConfig.profiles.find(
                (item) => item.name === stageConfig.profileName,
            ) ?? normalizedConfig.profiles[0];
        const profile = profileItem?.profile ?? null;

        if (!profile) {
            setAIResultMessage("请先配置接口");
            return;
        }
        if (profile.model.trim().length === 0) {
            setAIResultMessage("请先配置模型");
            return;
        }
        if (profile.url.trim().length === 0) {
            setAIResultMessage("请先配置 Base URL");
            return;
        }
        if (profile.apiKey.trim().length === 0) {
            setAIResultMessage("请先配置 API Key");
            return;
        }
        if (stageConfig.submitFieldKeys.length === 0) {
            setAIResultMessage("请先在 AI 配置中选择提交回答字段");
            return;
        }
        if (stageConfig.prompt.trim().length === 0) {
            setAIResultMessage("请先配置 Prompt");
            return;
        }

        if (stageKey === "final_verdict") {
            const independentSolvingResult =
                selectedRow.aiResults?.independent_solving?.trim() ?? "";
            if (independentSolvingResult.length === 0) {
                setAIResultMessage(
                    "请先执行第三阶段（Independent Solving），Final Verdict 需要依赖其结果",
                );
                return;
            }
        }

        const fields = buildStageFieldsForRow(
            activeFile.columns,
            selectedRow,
            stageKey,
            stageConfig,
            selectedRow.aiResults?.independent_solving ?? "",
        );

        if (fields.length === 0) {
            setAIResultMessage("当前记录没有可提交的回答字段");
            return;
        }

        aiStreamAbortRef.current?.abort();
        const controller = new AbortController();
        aiStreamAbortRef.current = controller;
        aiDetectStartedAtRef.current = Date.now();
        setAIDetectElapsedMs(0);
        setIsAIDetecting(true);
        setAIThinkingText("");
        setAIResultText("");
        setAIResultMessage("");
        setErrorMessage("");

        try {
            const streamResult = await requestAIDetectResult(
                {
                    provider: profile.provider,
                    url: profile.url,
                    model: profile.model,
                    apiKey: profile.apiKey,
                    prompt: stageConfig.prompt,
                    fields,
                    reasoningEffort: profile.reasoningEffort,
                    retryCount: profile.retryCount,
                },
                {
                    signal: controller.signal,
                    onAnswerChunk: (chunk) => {
                        setAIResultText((previous) => previous + chunk);
                    },
                    onThinkingChunk: (chunk) => {
                        setAIThinkingText((previous) => previous + chunk);
                    },
                },
            );
            setAIResultText(streamResult.answerText);
            setAIThinkingText("");
            const answerText = streamResult.answerText.trim();
            if (answerText.length === 0) {
                setAIResultMessage("AI 返回为空");
            } else {
                updateRowAIResult(
                    activeFile.fileId,
                    selectedRow.rowId,
                    stageKey,
                    answerText,
                );
                setAIResultMessage(
                    `AI 回答完成${stageLabel ? `（${stageLabel}）` : ""}，已写入 AI 检测结果`,
                );
            }
        } catch (error) {
            if (controller.signal.aborted) {
                setAIResultMessage("AI 回答已取消");
            } else {
                const message =
                    error instanceof Error ? error.message : "AI 回答失败";
                setAIResultMessage(message);
            }
        } finally {
            if (aiStreamAbortRef.current === controller) {
                aiStreamAbortRef.current = null;
            }
            if (aiDetectStartedAtRef.current) {
                setAIDetectElapsedMs(Date.now() - aiDetectStartedAtRef.current);
                aiDetectStartedAtRef.current = null;
            }
            setIsAIDetecting(false);
        }
    };

    const onRunAllAIDetect = () => {
        if (!canRunAIDetect) {
            return;
        }
        setActiveAIRunKey(AI_RUN_ALL_KEY);
        setAIThinkingText("");
        setAIResultText("");
        setAIResultMessage("");
        void onRunAIDetect(AI_RUN_ALL_KEY);
    };

    const onRunBatchAIAnswer = async (rowIds?: string[]) => {
        if (!activeFile) {
            return;
        }
        if (isAIDetecting || isAIBatchRunning) {
            return;
        }

        const normalizedConfig = normalizeAIDetectConfigForColumns(
            aiConfig,
            activeFile.columns,
        );
        syncActiveAIConfigState(normalizedConfig);
        const rowIdSet = new Set(activeFile.rows.map((row) => row.rowId));
        const normalizedTargetRowIds = rowIds
            ? Array.from(new Set(rowIds.filter((rowId) => rowIdSet.has(rowId))))
            : null;
        const selectedRowIdSet = normalizedTargetRowIds
            ? new Set(normalizedTargetRowIds)
            : null;
        const targetRows =
            normalizedTargetRowIds && normalizedTargetRowIds.length > 0
                ? activeFile.rows.filter(
                      (row) => selectedRowIdSet?.has(row.rowId) === true,
                  )
                : normalizedTargetRowIds
                  ? []
                  : activeFile.rows;
        if (targetRows.length === 0) {
            setErrorMessage(
                normalizedTargetRowIds
                    ? "请先至少选择一条数据再执行批量回答"
                    : "当前文件没有可执行的行数据",
            );
            return;
        }

        const targetFileId = activeFile.fileId;
        const targetFileName = activeFile.fileName;
        const targetColumns = activeFile.columns;

        if (activeAIRunKey === AI_RUN_ALL_KEY) {
            await runBatchAllStages({
                normalizedConfig,
                targetRows,
                targetColumns,
                targetFileId,
                targetFileName,
                normalizedTargetRowIds,
            });
            return;
        }

        const stageKey = activeAIRunKey;
        const stageConfig = normalizedConfig.stages[stageKey];
        const stageLabel = AI_STAGE_LABELS[stageKey]?.shortTitle ?? "";
        const profileItem =
            normalizedConfig.profiles.find(
                (item) => item.name === stageConfig.profileName,
            ) ?? normalizedConfig.profiles[0];
        const profile = profileItem?.profile ?? null;

        if (!profile) {
            setErrorMessage("请先配置接口");
            return;
        }
        if (profile.model.trim().length === 0) {
            setErrorMessage("请先配置模型");
            return;
        }
        if (profile.url.trim().length === 0) {
            setErrorMessage("请先配置 Base URL");
            return;
        }
        if (profile.apiKey.trim().length === 0) {
            setErrorMessage("请先配置 API Key");
            return;
        }
        if (stageConfig.submitFieldKeys.length === 0) {
            setErrorMessage("请先在 AI 配置中选择提交回答字段");
            return;
        }
        if (stageConfig.prompt.trim().length === 0) {
            setErrorMessage("请先配置 Prompt");
            return;
        }

        if (stageKey === "final_verdict") {
            const rowsWithoutIndependentSolving = activeFile.rows.filter(
                (row) =>
                    !row.aiResults?.independent_solving ||
                    row.aiResults.independent_solving.trim().length === 0,
            );
            if (rowsWithoutIndependentSolving.length > 0) {
                setErrorMessage(
                    `Final Verdict 批量回答需要依赖第三阶段（Independent Solving）结果，当前有 ${rowsWithoutIndependentSolving.length} 条缺失`,
                );
                return;
            }
        }

        let nextCursor = 0;
        const requestedConcurrency =
            normalizeAIBatchConcurrency(aiBatchConcurrency);
        const workerCount = Math.min(requestedConcurrency, targetRows.length);

        aiBatchAbortRef.current?.abort();
        const controller = new AbortController();
        aiBatchAbortRef.current = controller;
        startBatchPersistLoop(targetFileId);
        resetRowStreamProgress();
        resetRowBatchStatuses();
        setAIBatchTask({
            status: "running",
            fileId: targetFileId,
            fileName: targetFileName,
            total: targetRows.length,
            completed: 0,
            success: 0,
            failed: 0,
            message:
                normalizedTargetRowIds && normalizedTargetRowIds.length > 0
                    ? `已选择 ${targetRows.length} 条，批量执行 ${stageLabel}，并发 ${workerCount} 线程`
                    : `批量执行 ${stageLabel}，并发 ${workerCount} 线程`,
        });
        setErrorMessage("");
        setAIResultMessage("");
        primeRowStageProgress(targetRows, [stageKey]);

        const runWorker = async () => {
            while (!controller.signal.aborted) {
                const currentIndex = nextCursor;
                nextCursor += 1;
                if (currentIndex >= targetRows.length) {
                    return;
                }

                const row = targetRows[currentIndex];
                let rowFailed = false;
                try {
                    const fields = buildStageFieldsForRow(
                        targetColumns,
                        row,
                        stageKey,
                        stageConfig,
                        stageKey === "final_verdict"
                            ? (row.aiResults?.independent_solving ?? "")
                            : undefined,
                    );
                    const resultText = await runStageRequest({
                        stageConfig,
                        profile,
                        fields,
                        signal: controller.signal,
                        rowId: row.rowId,
                        stageKey,
                    });
                    updateRowAIResult(
                        targetFileId,
                        row.rowId,
                        stageKey,
                        resultText,
                    );
                } catch {
                    rowFailed = true;
                }

                if (controller.signal.aborted) {
                    return;
                }
                markRowBatchStatus(row.rowId, rowFailed ? "failed" : "success");
                setAIBatchTask((previous) => ({
                    ...previous,
                    completed: previous.completed + 1,
                    success: previous.success + (rowFailed ? 0 : 1),
                    failed: previous.failed + (rowFailed ? 1 : 0),
                }));
            }
        };

        try {
            await Promise.all(
                Array.from({ length: workerCount }, () => runWorker()),
            );

            if (controller.signal.aborted) {
                return;
            }

            setAIBatchTask((previous) => ({
                ...previous,
                message: "批量完成，正在写入 AI 检测结果",
            }));
            await flushPendingAIResults(targetFileId);
            const latestFile = latestFileStateRef.current[targetFileId];
            if (latestFile) {
                await persistFileState(latestFile);
            }
            setAIBatchTask((previous) => ({
                ...previous,
                status: "completed",
                message: `结果已写入 AI 检测结果${stageLabel ? `（${stageLabel}）` : ""}`,
            }));
            setErrorMessage("");
        } catch (error) {
            if (controller.signal.aborted) {
                return;
            }

            const message =
                error instanceof Error
                    ? error.message
                    : "批量 AI 回答任务执行失败";
            setAIBatchTask((previous) => ({
                ...previous,
                status: "completed",
                message,
            }));
            setErrorMessage(message);
        } finally {
            if (aiBatchAbortRef.current === controller) {
                aiBatchAbortRef.current = null;
            }
            stopBatchPersistLoop();
            resetRowStreamProgress();
        }
    };

    const onAIResultTextChange = (value: string) => {
        setAIThinkingText("");
        setAIResultText(value);
    };

    return {
        aiConfigList,
        aiConfig,
        draftAIConfig,
        setDraftAIConfig,
        aiConfigFormMessage,
        aiConfigLoading,
        aiConfigSaving,
        isAIStageConfigModalOpen,
        isAIProfileModalOpen,
        isAIRunModalOpen,
        setIsAIRunModalOpen,
        aiBatchTask,
        aiBatchProgressPercent,
        isAIBatchRunning,
        aiBatchConcurrency,
        setAIBatchConcurrency,
        activeAIRunKey,
        setActiveAIRunKey,
        aiModalStageKey,
        modalStageKey,
        modalFieldLabels,
        aiThinkingText,
        aiResultText,
        aiResultMessage,
        aiDetectElapsedText,
        aiMergedStreamText,
        canRunAIDetect,
        runAllTimerText,
        runAllStageTimers,
        rowStreamProgress,
        rowBatchStatuses,
        isAIDetecting,
        onOpenAIStageConfigModal,
        onOpenAIProfileModal,
        onCancelAIStageConfigModal,
        onCancelAIProfileModal,
        onToggleDraftAISubmitField,
        onSaveAIStageConfig,
        onSaveAIProfileConfig,
        onRunAIDetect,
        onRunAllAIDetect,
        onRunBatchAIAnswer,
        openAIRunModalForStage,
        onAIResultTextChange,
    };
};
