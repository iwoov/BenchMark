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
    normalizeLoadedNamedAIDetectConfigs,
    normalizeNamedAIDetectConfigsForColumns,
    requestAIChatResult,
    requestAIDetectResult,
} from "../ai-helpers";
import type { AIChatMessage, AIChatMessagePayload } from "../types";

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
    const [isAIRouteModalOpen, setIsAIRouteModalOpen] = useState(false);
    const [isAIChatConfigModalOpen, setIsAIChatConfigModalOpen] =
        useState(false);
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
    const [chatMessages, setChatMessages] = useState<AIChatMessage[]>([]);
    const [chatInput, setChatInput] = useState("");
    const [chatStatusMessage, setChatStatusMessage] = useState("");
    const [activeChatRouteName, setActiveChatRouteName] = useState("");
    const [isAIChatting, setIsAIChatting] = useState(false);
    const [aiChatElapsedMs, setAIChatElapsedMs] = useState(0);

    const aiStreamAbortRef = useRef<AbortController | null>(null);
    const aiChatAbortRef = useRef<AbortController | null>(null);
    const aiBatchAbortRef = useRef<AbortController | null>(null);
    const aiDetectStartedAtRef = useRef<number | null>(null);
    const aiChatStartedAtRef = useRef<number | null>(null);
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
            aiChatAbortRef.current?.abort();
            aiChatAbortRef.current = null;
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
        if (!isAIChatting) {
            return;
        }
        const startedAt = aiChatStartedAtRef.current ?? Date.now();
        aiChatStartedAtRef.current = startedAt;
        const timerId = window.setInterval(() => {
            setAIChatElapsedMs(Date.now() - startedAt);
        }, 250);
        return () => {
            window.clearInterval(timerId);
        };
    }, [isAIChatting]);

    useEffect(() => {
        setAIThinkingText("");
        setAIResultText("");
        setAIResultMessage("");
        aiStreamAbortRef.current?.abort();
        aiStreamAbortRef.current = null;
        aiDetectStartedAtRef.current = null;
        setAIDetectElapsedMs(0);
        setIsAIDetecting(false);
        aiChatAbortRef.current?.abort();
        aiChatAbortRef.current = null;
        aiChatStartedAtRef.current = null;
        setAIChatElapsedMs(0);
        setIsAIChatting(false);
        setChatMessages([]);
        setChatInput("");
        setChatStatusMessage("");
        setActiveChatRouteName(aiConfig.chat.routeName);
    }, [activeFileId, selectedRowId, aiConfig.chat.routeName]);

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
            setIsAIRouteModalOpen(false);
            setIsAIChatConfigModalOpen(false);
            setActiveChatRouteName(nextConfig.chat.routeName);
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
                    providers?: unknown;
                    routes?: unknown;
                    stages?: unknown;
                    chat?: unknown;
                };
                let loadedConfigs = normalizeLoadedNamedAIDetectConfigs([
                    {
                        name: DEFAULT_AI_CONFIG_NAME,
                        config: {
                            providers: payload.providers,
                            routes: payload.routes,
                            stages: payload.stages,
                            chat: payload.chat,
                        },
                    },
                ]);
                if (loadedConfigs.length === 0) {
                    loadedConfigs = [
                        {
                            name: DEFAULT_AI_CONFIG_NAME,
                            config: createDefaultAIDetectConfig(),
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
                setActiveChatRouteName(activeConfig.chat.routeName);
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
                setActiveChatRouteName(fallbackConfig.chat.routeName);
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
        setActiveChatRouteName((previous) => {
            if (
                previous &&
                nextSelectedConfig.routes.some((item) => item.name === previous)
            ) {
                return previous;
            }
            return nextSelectedConfig.chat.routeName;
        });
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
    const aiChatElapsedText = formatDuration(aiChatElapsedMs);
    const aiMergedStreamText = useMemo(
        () => composeAISaveText(aiResultText, aiThinkingText),
        [aiResultText, aiThinkingText],
    );

    const canRunAIDetect =
        Boolean(selectedRow) &&
        !isAIDetecting &&
        !isAIChatting &&
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

    const resolveRouteForStage = (
        config: AIDetectConfig,
        stageConfig: AIDetectConfig["stages"][AIDetectStageKey],
    ) => {
        const routeItem =
            config.routes.find((item) => item.name === stageConfig.routeName) ??
            config.routes[0];
        return routeItem ?? null;
    };

    const validateStageSetup = (
        stageKey: AIDetectStageKey,
        stageConfig: AIDetectConfig["stages"][AIDetectStageKey],
        route: AIDetectConfig["routes"][number] | null,
        includeStagePrefix: boolean,
    ): string | null => {
        const prefix = includeStagePrefix
            ? `【${getStageLabel(stageKey)}】`
            : "";
        if (!route) {
            return `${prefix}请先配置模型路由`;
        }
        if (route.model.trim().length === 0) {
            return `${prefix}请先配置模型`;
        }
        if (route.steps.length === 0) {
            return `${prefix}请先至少配置一个提供商回退步骤`;
        }
        if (stageConfig.submitFieldKeys.length === 0) {
            return `${prefix}请先在 AI 配置中选择提交回答字段`;
        }
        if (stageConfig.prompt.trim().length === 0) {
            return `${prefix}请先配置 Prompt`;
        }
        return null;
    };

    const buildChatFieldsForRow = (
        columns: ParsedColumn[],
        row: ParsedRow,
        config: AIDetectConfig,
    ) =>
        buildAIDetectFieldsForRow(
            columns,
            row,
            config.chat.defaultSubmitFieldKeys,
        );

    const validateChatSetup = (
        config: AIDetectConfig,
        routeName: string,
    ): {
        route: AIDetectConfig["routes"][number] | null;
        error: string | null;
    } => {
        const route =
            config.routes.find((item) => item.name === routeName) ??
            config.routes.find((item) => item.name === config.chat.routeName) ??
            config.routes[0] ??
            null;
        if (!route) {
            return { route: null, error: "请先配置模型路由" };
        }
        if (route.model.trim().length === 0) {
            return { route, error: "请先配置模型" };
        }
        if (route.steps.length === 0) {
            return { route, error: "请先至少配置一个提供商回退步骤" };
        }
        if (config.chat.prompt.trim().length === 0) {
            return { route, error: "请先配置聊天 Prompt" };
        }
        return { route, error: null };
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
        route,
        fields,
        signal,
        rowId,
        stageKey,
    }: {
        stageConfig: AIDetectConfig["stages"][AIDetectStageKey];
        route: AIDetectConfig["routes"][number];
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
                routeName: route.name,
                prompt: stageConfig.prompt,
                fields,
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
                route: AIDetectConfig["routes"][number];
            }
        >();
        for (const stageKey of AI_STAGE_ORDER) {
            const stageConfig = normalizedConfig.stages[stageKey];
            const route = resolveRouteForStage(
                normalizedConfig,
                stageConfig,
            );
            const error = validateStageSetup(
                stageKey,
                stageConfig,
                route,
                true,
            );
            if (error || !route) {
                setErrorMessage(error ?? "请先配置模型路由");
                return;
            }
            stageRunMap.set(stageKey, { stageConfig, route });
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
                        route: precheckConfig.route,
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
                        route: contextConfig.route,
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
                        route: independentConfig.route,
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
                                route: finalConfig.route,
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
        setActiveChatRouteName((previous) => {
            if (nextConfig.routes.some((item) => item.name === previous)) {
                return previous;
            }
            return nextConfig.chat.routeName;
        });
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
        setIsAIRouteModalOpen(false);
        setIsAIChatConfigModalOpen(false);
        navigateToSection("settings", "ai");
    };

    const onOpenAIProfileModal = () => {
        if (!prepareDraftAIConfig()) {
            return;
        }
        setIsAIProfileModalOpen(true);
        setIsAIStageConfigModalOpen(false);
        setIsAIRouteModalOpen(false);
        setIsAIChatConfigModalOpen(false);
        navigateToSection("settings", "ai");
    };

    const onOpenAIRouteModal = () => {
        if (!prepareDraftAIConfig()) {
            return;
        }
        setIsAIRouteModalOpen(true);
        setIsAIProfileModalOpen(false);
        setIsAIStageConfigModalOpen(false);
        setIsAIChatConfigModalOpen(false);
        navigateToSection("settings", "ai");
    };

    const onOpenAIChatConfigModal = () => {
        if (!prepareDraftAIConfig()) {
            return;
        }
        setIsAIChatConfigModalOpen(true);
        setIsAIProfileModalOpen(false);
        setIsAIStageConfigModalOpen(false);
        setIsAIRouteModalOpen(false);
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

    const onCancelAIRouteModal = () => {
        setDraftAIConfig(cloneAIDetectConfig(aiConfig));
        setAIConfigFormMessage("");
        setIsAIRouteModalOpen(false);
    };

    const onCancelAIChatConfigModal = () => {
        setDraftAIConfig(cloneAIDetectConfig(aiConfig));
        setAIConfigFormMessage("");
        setIsAIChatConfigModalOpen(false);
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

    const onToggleDraftAIChatSubmitField = (columnKey: string) => {
        setDraftAIConfig((previous) => {
            const exists =
                previous.chat.defaultSubmitFieldKeys.includes(columnKey);
            const defaultSubmitFieldKeys = exists
                ? previous.chat.defaultSubmitFieldKeys.filter(
                      (key) => key !== columnKey,
                  )
                : [...previous.chat.defaultSubmitFieldKeys, columnKey];
            return {
                ...previous,
                chat: {
                    ...previous.chat,
                    defaultSubmitFieldKeys,
                },
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

        if (!nextConfig.providers || nextConfig.providers.length === 0) {
            setAIConfigFormMessage("请至少配置一个模型提供商");
            return;
        }

        const providerNameSet = new Set<string>();
        for (const provider of nextConfig.providers) {
            const providerName = provider.name.trim();
            if (providerName.length === 0) {
                setAIConfigFormMessage("模型提供商名称不能为空");
                return;
            }
            if (providerNameSet.has(providerName)) {
                setAIConfigFormMessage(`模型提供商名称重复：${providerName}`);
                return;
            }
            providerNameSet.add(providerName);
            if (provider.apiUrl.trim().length === 0) {
                setAIConfigFormMessage(`【${providerName}】API URL 不能为空`);
                return;
            }
            if (provider.apiKey.trim().length === 0) {
                setAIConfigFormMessage(`【${providerName}】API Key 不能为空`);
                return;
            }
        }

        if (!nextConfig.routes || nextConfig.routes.length === 0) {
            setAIConfigFormMessage("请至少配置一个模型路由");
            return;
        }

        const routeNameSet = new Set<string>();
        for (const route of nextConfig.routes) {
            const routeName = route.name.trim();
            if (routeName.length === 0) {
                setAIConfigFormMessage("模型路由名称不能为空");
                return;
            }
            if (routeNameSet.has(routeName)) {
                setAIConfigFormMessage(`模型路由名称重复：${routeName}`);
                return;
            }
            routeNameSet.add(routeName);
            if (route.model.trim().length === 0) {
                setAIConfigFormMessage(`【${routeName}】模型不能为空`);
                return;
            }
            if (route.steps.length === 0) {
                setAIConfigFormMessage(`【${routeName}】请至少配置一个回退步骤`);
                return;
            }
            for (const step of route.steps) {
                if (!providerNameSet.has(step.providerName)) {
                    setAIConfigFormMessage(
                        `【${routeName}】引用了不存在的模型提供商：${step.providerName}`,
                    );
                    return;
                }
            }
        }

        if (!skipStageValidation) {
            for (const stageKey of AI_STAGE_ORDER) {
                const stageConfig = nextConfig.stages[stageKey];
                const stageLabel =
                    AI_STAGE_LABELS[stageKey]?.shortTitle ?? stageKey;

                if (!routeNameSet.has(stageConfig.routeName)) {
                    setAIConfigFormMessage(
                        `【${stageLabel}】请选择有效的模型路由`,
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
            const saveProvidersResponse = await fetch("/api/ai-config/providers", {
                method: "PUT",
                headers: {
                    "Content-Type": "application/json",
                },
                body: JSON.stringify({
                    providers: nextConfig.providers,
                }),
            });
            if (!saveProvidersResponse.ok) {
                const payload = (await saveProvidersResponse
                    .json()
                    .catch(() => ({}))) as { message?: string };
                throw new Error(payload.message ?? "保存模型提供商失败");
            }

            const saveRoutesResponse = await fetch("/api/ai-config/routes", {
                method: "PUT",
                headers: {
                    "Content-Type": "application/json",
                },
                body: JSON.stringify({
                    routes: nextConfig.routes,
                }),
            });
            if (!saveRoutesResponse.ok) {
                const payload = (await saveRoutesResponse
                    .json()
                    .catch(() => ({}))) as { message?: string };
                throw new Error(payload.message ?? "保存模型路由失败");
            }

            if (!skipStageValidation) {
                const saveStagesResponse = await fetch(
                    `/api/ai-config/${encodeURIComponent(activeFile.fileName)}/stages`,
                    {
                        method: "PUT",
                        headers: {
                            "Content-Type": "application/json",
                        },
                        body: JSON.stringify({
                            stages: nextConfig.stages,
                        }),
                    },
                );
                if (!saveStagesResponse.ok) {
                    const payload = (await saveStagesResponse
                        .json()
                        .catch(() => ({}))) as { message?: string };
                    throw new Error(payload.message ?? "保存阶段任务失败");
                }
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
            setIsAIRouteModalOpen(false);
            setIsAIChatConfigModalOpen(false);
        } catch (error) {
            const message =
                error instanceof Error ? error.message : "保存 AI 配置失败";
            setAIConfigFormMessage(message);
        } finally {
            setAIConfigSaving(false);
        }
    };

    const onSaveAIProfileConfig = () => onSaveAIConfig(true);
    const onSaveAIRouteConfig = () => onSaveAIConfig(true);
    const onSaveAIStageConfig = () => onSaveAIConfig(false);

    const onSaveAIChatConfig = async () => {
        if (!activeFile) {
            return;
        }

        const nextConfig = normalizeAIDetectConfigForColumns(
            draftAIConfig,
            activeFile.columns,
        );
        const routeNameSet = new Set(nextConfig.routes.map((item) => item.name));
        if (!routeNameSet.has(nextConfig.chat.routeName)) {
            setAIConfigFormMessage("聊天模型路由无效");
            return;
        }
        if (nextConfig.chat.prompt.trim().length === 0) {
            setAIConfigFormMessage("聊天 Prompt 不能为空");
            return;
        }

        setAIConfigSaving(true);
        setAIConfigFormMessage("");
        setErrorMessage("");

        try {
            const response = await fetch(
                `/api/ai-config/${encodeURIComponent(activeFile.fileName)}/chat`,
                {
                    method: "PUT",
                    headers: {
                        "Content-Type": "application/json",
                    },
                    body: JSON.stringify({
                        chat: nextConfig.chat,
                    }),
                },
            );
            if (!response.ok) {
                const payload = (await response
                    .json()
                    .catch(() => ({}))) as { message?: string };
                throw new Error(payload.message ?? "保存聊天配置失败");
            }

            syncActiveAIConfigState(nextConfig);
            setDraftAIConfig(cloneAIDetectConfig(nextConfig));
            setActiveChatRouteName(nextConfig.chat.routeName);
            setAIConfigFormMessage("");
            setIsAIChatConfigModalOpen(false);
        } catch (error) {
            const message =
                error instanceof Error ? error.message : "保存聊天配置失败";
            setAIConfigFormMessage(message);
        } finally {
            setAIConfigSaving(false);
        }
    };

    const onClearAIChatSession = () => {
        if (isAIChatting) {
            return;
        }
        setChatMessages([]);
        setChatInput("");
        setChatStatusMessage("");
    };

    const onSendAIChatMessage = async () => {
        if (!activeFile || !selectedRow) {
            return;
        }
        if (isAIDetecting || isAIBatchRunning || isAIChatting) {
            return;
        }

        const content = chatInput.trim();
        if (content.length === 0) {
            setChatStatusMessage("请输入聊天内容");
            return;
        }

        const normalizedConfig = normalizeAIDetectConfigForColumns(
            aiConfig,
            activeFile.columns,
        );
        syncActiveAIConfigState(normalizedConfig);
        const { route, error } = validateChatSetup(
            normalizedConfig,
            activeChatRouteName,
        );
        if (error || !route) {
            setChatStatusMessage(error ?? "请先配置聊天模型路由");
            return;
        }
        if (route.name !== activeChatRouteName) {
            setActiveChatRouteName(route.name);
        }

        const userMessage: AIChatMessage = {
            id: `user-${Date.now()}`,
            role: "user",
            content,
            createdAt: Date.now(),
            status: "done",
        };
        const assistantMessageId = `assistant-${Date.now()}`;
        const nextMessages = [...chatMessages, userMessage];
        const requestMessages: AIChatMessagePayload[] = nextMessages.map(
            (message) => ({
                role: message.role,
                content: message.content,
            }),
        );
        const chatFields = buildChatFieldsForRow(
            activeFile.columns,
            selectedRow,
            normalizedConfig,
        );

        setChatMessages([
            ...nextMessages,
            {
                id: assistantMessageId,
                role: "assistant",
                content: "",
                createdAt: Date.now(),
                status: "streaming",
            },
        ]);
        setChatInput("");
        setChatStatusMessage("");
        setIsAIChatting(true);
        setAIChatElapsedMs(0);
        aiChatStartedAtRef.current = Date.now();
        aiChatAbortRef.current?.abort();
        const controller = new AbortController();
        aiChatAbortRef.current = controller;

        try {
            const streamResult = await requestAIChatResult(
                {
                    routeName: route.name,
                    prompt: normalizedConfig.chat.prompt,
                    messages: requestMessages,
                    fields: chatFields,
                },
                {
                    signal: controller.signal,
                    onAnswerChunk: (chunk) => {
                        setChatMessages((previous) =>
                            previous.map((message) =>
                                message.id === assistantMessageId
                                    ? {
                                          ...message,
                                          content: message.content + chunk,
                                      }
                                    : message,
                            ),
                        );
                    },
                },
            );
            const answerText = streamResult.answerText.trim();
            if (answerText.length === 0) {
                setChatMessages((previous) =>
                    previous.map((message) =>
                        message.id === assistantMessageId
                            ? {
                                  ...message,
                                  content: "AI 返回为空",
                                  status: "error",
                              }
                            : message,
                    ),
                );
                setChatStatusMessage("AI 返回为空");
            } else {
                setChatMessages((previous) =>
                    previous.map((message) =>
                        message.id === assistantMessageId
                            ? {
                                  ...message,
                                  content: answerText,
                                  status: "done",
                              }
                            : message,
                    ),
                );
                setChatStatusMessage("");
            }
        } catch (error) {
            if (controller.signal.aborted) {
                setChatStatusMessage("聊天已取消");
            } else {
                const message =
                    error instanceof Error ? error.message : "AI 聊天失败";
                setChatMessages((previous) =>
                    previous.map((item) =>
                        item.id === assistantMessageId
                            ? {
                                  ...item,
                                  content: item.content || message,
                                  status: "error",
                              }
                            : item,
                    ),
                );
                setChatStatusMessage(message);
            }
        } finally {
            if (aiChatAbortRef.current === controller) {
                aiChatAbortRef.current = null;
            }
            if (aiChatStartedAtRef.current) {
                setAIChatElapsedMs(Date.now() - aiChatStartedAtRef.current);
                aiChatStartedAtRef.current = null;
            }
            setIsAIChatting(false);
        }
    };

    const onRunAIDetect = async (runKeyOverride?: AIDetectRunKey) => {
        if (!activeFile || !selectedRow) {
            return;
        }
        if (isAIChatting) {
            setAIResultMessage("AI 聊天进行中，暂不可发起检测");
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
                    route: AIDetectConfig["routes"][number];
                }
            >();
            for (const stageKey of AI_STAGE_ORDER) {
                const stageConfig = normalizedConfig.stages[stageKey];
                const route = resolveRouteForStage(
                    normalizedConfig,
                    stageConfig,
                );
                const error = validateStageSetup(
                    stageKey,
                    stageConfig,
                    route,
                    true,
                );
                if (error || !route) {
                    setAIResultMessage(error ?? "请先配置模型路由");
                    return;
                }
                stageRunMap.set(stageKey, { stageConfig, route });
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
                    route: precheckConfig.route,
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
                    route: contextConfig.route,
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
                    route: independentConfig.route,
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
                            route: finalConfig.route,
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
        const route =
            normalizedConfig.routes.find(
                (item) => item.name === stageConfig.routeName,
            ) ?? normalizedConfig.routes[0];

        if (!route) {
            setAIResultMessage("请先配置模型路由");
            return;
        }
        if (route.model.trim().length === 0) {
            setAIResultMessage("请先配置模型");
            return;
        }
        if (route.steps.length === 0) {
            setAIResultMessage("请先配置模型提供商回退步骤");
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
                    routeName: route.name,
                    prompt: stageConfig.prompt,
                    fields,
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
        if (isAIDetecting || isAIBatchRunning || isAIChatting) {
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
        const route =
            normalizedConfig.routes.find(
                (item) => item.name === stageConfig.routeName,
            ) ?? normalizedConfig.routes[0];

        if (!route) {
            setErrorMessage("请先配置模型路由");
            return;
        }
        if (route.model.trim().length === 0) {
            setErrorMessage("请先配置模型");
            return;
        }
        if (route.steps.length === 0) {
            setErrorMessage("请先配置模型提供商回退步骤");
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
                        route,
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
        isAIRouteModalOpen,
        isAIChatConfigModalOpen,
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
        chatMessages,
        chatInput,
        setChatInput,
        chatStatusMessage,
        activeChatRouteName,
        setActiveChatRouteName,
        isAIChatting,
        aiChatElapsedText,
        onOpenAIStageConfigModal,
        onOpenAIProfileModal,
        onOpenAIRouteModal,
        onOpenAIChatConfigModal,
        onCancelAIStageConfigModal,
        onCancelAIProfileModal,
        onCancelAIRouteModal,
        onCancelAIChatConfigModal,
        onToggleDraftAISubmitField,
        onToggleDraftAIChatSubmitField,
        onSaveAIStageConfig,
        onSaveAIProfileConfig,
        onSaveAIRouteConfig,
        onSaveAIChatConfig,
        onRunAIDetect,
        onRunAllAIDetect,
        onRunBatchAIAnswer,
        onSendAIChatMessage,
        onClearAIChatSession,
        openAIRunModalForStage,
        onAIResultTextChange,
    };
};
