import { useEffect, useMemo, useRef, useState } from "react";
import type {
    AICleaningToolKey,
    AICleaningToolResult,
    AIDetectStageKey,
    AIEvaluationAttemptResult,
    FileViewState,
    FilterCondition,
    ParsedFile,
} from "../../types";
import type { ColumnPrefsConfig, MainSection, SettingsSection } from "../types";
import { AI_STAGE_ORDER } from "../constants";
import {
    applyColumnConfigToFile,
    downloadBlob,
    getAllColumnKeys,
    getFieldSignature,
    getFileNameFromDisposition,
    normalizeColumnSelection,
    normalizeLoadedFileState,
    normalizeStatisticsConfig,
    toViewState,
} from "../file-helpers";
import type { StatisticsChartType } from "../../types";

type NavigateToSection = (
    section: MainSection,
    settingsSection?: SettingsSection,
    rowId?: string | null,
    options?: { replace?: boolean },
) => void;

type PendingConfigMode = "import" | "edit";
type UploadMode = "create" | "merge" | "add-datasource";
type ProjectNameDialogMode = "create" | "rename" | "add-datasource";
const LAST_ACTIVE_FILE_ID_STORAGE_KEY = "benchmark:last-active-file-id";
const FILTER_CONDITIONS_STORAGE_PREFIX = "benchmark:filters:";

function saveFilterConditionsToLocal(
    fileId: string,
    conditions: FilterCondition[],
): void {
    try {
        const key = `${FILTER_CONDITIONS_STORAGE_PREFIX}${fileId}`;
        if (conditions.length === 0) {
            window.localStorage.removeItem(key);
            return;
        }
        window.localStorage.setItem(key, JSON.stringify(conditions));
    } catch {
        // Ignore storage errors.
    }
}

function loadFilterConditionsFromLocal(fileId: string): FilterCondition[] {
    try {
        const key = `${FILTER_CONDITIONS_STORAGE_PREFIX}${fileId}`;
        const raw = window.localStorage.getItem(key);
        if (!raw) {
            return [];
        }
        const parsed = JSON.parse(raw) as unknown;
        if (!Array.isArray(parsed)) {
            return [];
        }
        return parsed.filter(
            (item): item is FilterCondition =>
                !!item &&
                typeof item === "object" &&
                typeof (item as FilterCondition).id === "string" &&
                typeof (item as FilterCondition).columnKey === "string" &&
                typeof (item as FilterCondition).value === "string",
        );
    } catch {
        return [];
    }
}

export const useFileStore = ({
    navigateToSection,
    setErrorMessage,
}: {
    navigateToSection: NavigateToSection;
    setErrorMessage: (value: string) => void;
}) => {
    const [files, setFiles] = useState<FileViewState[]>([]);
    const [activeFileId, setActiveFileId] = useState<string | null>(() => {
        if (typeof window === "undefined") {
            return null;
        }
        const saved = window.localStorage.getItem(
            LAST_ACTIVE_FILE_ID_STORAGE_KEY,
        );
        return saved && saved.trim().length > 0 ? saved : null;
    });
    const [isUploading, setIsUploading] = useState(false);
    const [isExporting, setIsExporting] = useState(false);
    const [isExportingEvaluation, setIsExportingEvaluation] = useState(false);
    const [pendingFile, setPendingFile] = useState<ParsedFile | null>(null);
    const [pendingSelectedDisplayKeys, setPendingSelectedDisplayKeys] =
        useState<string[]>([]);
    const [pendingEditableColumnKeys, setPendingEditableColumnKeys] = useState<
        string[]
    >([]);
    const [pendingConfigNotice, setPendingConfigNotice] = useState<string>("");
    const [pendingConfigMode, setPendingConfigMode] =
        useState<PendingConfigMode>("import");
    const [initialLoadComplete, setInitialLoadComplete] = useState(false);
    const [projectNameDialogMode, setProjectNameDialogMode] =
        useState<ProjectNameDialogMode | null>(null);
    const [projectNameDraft, setProjectNameDraft] = useState("");
    const [projectNameDialogError, setProjectNameDialogError] = useState("");
    const [projectNameTargetFileId, setProjectNameTargetFileId] = useState<
        string | null
    >(null);
    const [pendingDataSourceNameDraft, setPendingDataSourceNameDraft] =
        useState("");
    const [removingFileId, setRemovingFileId] = useState<string | null>(null);

    const uploadInputRef = useRef<HTMLInputElement>(null);
    const persistTimersRef = useRef<Record<string, number>>({});
    const pendingPersistRef = useRef<Record<string, FileViewState>>({});
    const persistAIResultsQueueRef = useRef<Record<string, Promise<void>>>({});
    const latestFileStateRef = useRef<Record<string, FileViewState>>({});
    const stateVersionRef = useRef<Record<string, number>>({});
    const pendingUploadModeRef = useRef<UploadMode>("create");
    const pendingCreateProjectNameRef = useRef("");
    const pendingDataSourceNameRef = useRef("");
    const pendingDataSourceGroupIdRef = useRef<string | null>(null);
    const pendingMergeTargetFileIdRef = useRef<string | null>(null);

    const activeFile = useMemo(
        () =>
            files.find((item) => item.fileId === activeFileId) ??
            files[0] ??
            null,
        [files, activeFileId],
    );

    useEffect(() => {
        if (typeof window === "undefined") {
            return;
        }
        if (activeFileId && activeFileId.trim().length > 0) {
            window.localStorage.setItem(
                LAST_ACTIVE_FILE_ID_STORAGE_KEY,
                activeFileId,
            );
            return;
        }
        window.localStorage.removeItem(LAST_ACTIVE_FILE_ID_STORAGE_KEY);
    }, [activeFileId]);

    useEffect(() => {
        files.forEach((file) => {
            latestFileStateRef.current[file.fileId] = file;
        });
    }, [files]);

    useEffect(() => {
        let disposed = false;
        const controller = new AbortController();
        const shouldLogAIResults =
            typeof window !== "undefined" &&
            window.localStorage.getItem("debug_ai_results") === "1";

        const logAIResultsSummary = (loadedFiles: FileViewState[]) => {
            loadedFiles.forEach((file) => {
                const stageCounts: Record<AIDetectStageKey, number> = {
                    precheck: 0,
                    context_audit: 0,
                    independent_solving: 0,
                    final_verdict: 0,
                };
                let rowsWithAI = 0;
                file.rows.forEach((row) => {
                    const aiResults = row.aiResults ?? {};
                    let hasAny = false;
                    AI_STAGE_ORDER.forEach((stageKey) => {
                        const value = aiResults[stageKey];
                        if (
                            typeof value === "string" &&
                            value.trim().length > 0
                        ) {
                            stageCounts[stageKey] += 1;
                            hasAny = true;
                        }
                    });
                    if (hasAny) {
                        rowsWithAI += 1;
                    }
                });
                // eslint-disable-next-line no-console
                console.log(
                    `[AIResultsLoaded] fileId=${file.fileId} fileName=${file.fileName} rows=${file.rows.length} rowsWithAI=${rowsWithAI} stageCounts=${JSON.stringify(
                        stageCounts,
                    )}`,
                );
            });
        };

        const loadPersistedFiles = async () => {
            try {
                const response = await fetch("/api/files", {
                    signal: controller.signal,
                });
                if (!response.ok) {
                    setInitialLoadComplete(true);
                    return;
                }

                const payload = (await response.json()) as { files?: unknown };
                const rawFiles = Array.isArray(payload.files)
                    ? payload.files
                    : [];
                const restoredFiles = rawFiles
                    .map((item) => normalizeLoadedFileState(item))
                    .filter((item): item is FileViewState => item !== null)
                    .map((file) => {
                        const savedConditions = loadFilterConditionsFromLocal(
                            file.fileId,
                        );
                        if (savedConditions.length === 0) {
                            return file;
                        }
                        const columnKeys = new Set(
                            file.columns.map((column) => column.key),
                        );
                        const validConditions = savedConditions.filter(
                            (condition) =>
                                columnKeys.has(condition.columnKey) &&
                                condition.value.trim().length > 0,
                        );
                        return validConditions.length > 0
                            ? { ...file, filterConditions: validConditions }
                            : file;
                    });

                if (shouldLogAIResults && restoredFiles.length > 0) {
                    logAIResultsSummary(restoredFiles);
                }

                if (disposed) {
                    return;
                }

                if (restoredFiles.length === 0) {
                    setActiveFileId(null);
                    setInitialLoadComplete(true);
                    return;
                }

                setFiles((previous) => {
                    if (previous.length === 0) {
                        return restoredFiles;
                    }

                    const merged = new Map<string, FileViewState>();
                    previous.forEach((file) => merged.set(file.fileId, file));
                    restoredFiles.forEach((file) =>
                        merged.set(file.fileId, file),
                    );
                    return Array.from(merged.values());
                });
                setActiveFileId((previous) =>
                    previous &&
                    restoredFiles.some((file) => file.fileId === previous)
                        ? previous
                        : restoredFiles[0].fileId,
                );
                setInitialLoadComplete(true);
            } catch {
                if (!disposed) {
                    setInitialLoadComplete(true);
                }
            }
        };

        void loadPersistedFiles();

        return () => {
            disposed = true;
            controller.abort();
        };
    }, []);

    useEffect(() => {
        return () => {
            Object.values(persistTimersRef.current).forEach((timerId) => {
                window.clearTimeout(timerId);
            });
            persistTimersRef.current = {};
            pendingPersistRef.current = {};
        };
    }, []);

    const createVersionedStatePayload = (file: FileViewState) => {
        const currentVersion = stateVersionRef.current[file.fileId] ?? 0;
        const nextVersion = Math.max(Date.now(), currentVersion + 1);
        stateVersionRef.current[file.fileId] = nextVersion;
        const shouldPreserveRows =
            file.detailLoaded !== true &&
            (file.rowCount ?? 0) > 0 &&
            file.rows.length === 0;
        const sanitizedState: FileViewState = {
            ...file,
            filterConditions: [],
            rows: shouldPreserveRows
                ? []
                : file.rows.map(
                      ({ cleaningResults, evaluationResults, ...row }) => row,
                  ),
        };
        return {
            state: {
                ...sanitizedState,
                ...(shouldPreserveRows ? { rows: undefined } : {}),
                clientStateVersion: nextVersion,
            },
            preserveRows: shouldPreserveRows,
            clientStateVersion: nextVersion,
        };
    };

    const persistFileState = async (file: FileViewState) => {
        try {
            const statePayload = createVersionedStatePayload(file);
            const response = await fetch(
                `/api/files/${encodeURIComponent(file.fileId)}/state`,
                {
                    method: "PUT",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({
                        state: statePayload.state,
                        preserveRows: statePayload.preserveRows,
                    }),
                },
            );
            if (response.ok || response.status === 409) {
                return;
            }
            throw new Error(`save failed: ${response.status}`);
        } catch (error) {
            // eslint-disable-next-line no-console
            console.error(
                "[PersistFileState] Failed to save file state:",
                error,
            );
        }
    };

    const persistAIResults = (
        fileId: string,
        stageKey: AIDetectStageKey,
        results: Record<string, string>,
        fallbackState?: FileViewState,
    ) => {
        if (Object.keys(results).length === 0) {
            return;
        }
        const enqueuePersist = (job: () => Promise<void>) => {
            const previous =
                persistAIResultsQueueRef.current[fileId] ?? Promise.resolve();
            const next = previous.then(job).catch((error) => {
                // eslint-disable-next-line no-console
                console.error(
                    "[PersistAIResults] Failed to save AI results:",
                    error,
                );
                if (fallbackState) {
                    void persistFileState(fallbackState);
                }
            });
            persistAIResultsQueueRef.current[fileId] = next;
        };

        enqueuePersist(async () => {
            const endpoint = `/api/files/${encodeURIComponent(fileId)}/ai-results`;
            const payload = JSON.stringify({ stageKey, results });
            const response = await fetch(endpoint, {
                method: "PUT",
                headers: { "Content-Type": "application/json" },
                body: payload,
            });
            if (response.ok) {
                return;
            }
            if (response.status === 404 && fallbackState) {
                await persistFileState(fallbackState);
                const retryResponse = await fetch(endpoint, {
                    method: "PUT",
                    headers: { "Content-Type": "application/json" },
                    body: payload,
                });
                if (retryResponse.ok) {
                    return;
                }
            }
            throw new Error("Failed to save AI results");
        });
    };

    const flushPendingAIResults = async (fileId: string) => {
        const pending = persistAIResultsQueueRef.current[fileId];
        if (!pending) {
            return;
        }
        try {
            await pending;
        } catch (error) {
            // eslint-disable-next-line no-console
            console.error("[PersistAIResults] Pending flush failed:", error);
        }
    };

    const cancelScheduledPersist = (fileId: string) => {
        const timerId = persistTimersRef.current[fileId];
        if (timerId !== undefined) {
            window.clearTimeout(timerId);
            delete persistTimersRef.current[fileId];
        }
        delete pendingPersistRef.current[fileId];
    };

    const schedulePersistFileState = (
        file: FileViewState,
        delayMs: number = 400,
    ) => {
        cancelScheduledPersist(file.fileId);
        pendingPersistRef.current[file.fileId] = file;
        const timerId = window.setTimeout(() => {
            const latest = pendingPersistRef.current[file.fileId];
            if (latest) {
                persistFileState(latest);
            }
            delete pendingPersistRef.current[file.fileId];
            delete persistTimersRef.current[file.fileId];
        }, delayMs);
        persistTimersRef.current[file.fileId] = timerId;
    };

    const upsertFileState = (nextFile: FileViewState) => {
        latestFileStateRef.current[nextFile.fileId] = nextFile;
        setFiles((previous) => {
            const exists = previous.some(
                (file) => file.fileId === nextFile.fileId,
            );
            if (!exists) {
                return [...previous, nextFile];
            }
            return previous.map((file) =>
                file.fileId === nextFile.fileId ? nextFile : file,
            );
        });
    };

    const resetProjectNameDialog = () => {
        setProjectNameDialogMode(null);
        setProjectNameDraft("");
        setProjectNameDialogError("");
        setProjectNameTargetFileId(null);
        setPendingDataSourceNameDraft("");
    };

    const patchActiveFile = (
        updater: (file: FileViewState) => FileViewState,
    ) => {
        if (!activeFile) {
            return;
        }

        const nextFile = updater(activeFile);
        if (nextFile === activeFile) {
            return;
        }

        setFiles((previous) =>
            previous.map((file) =>
                file.fileId === nextFile.fileId ? nextFile : file,
            ),
        );
        latestFileStateRef.current[nextFile.fileId] = nextFile;
        schedulePersistFileState(nextFile);
    };

    const onEditCell = (rowId: string, columnKey: string, value: string) => {
        patchActiveFile((file) => ({
            ...file,
            rows: file.rows.map((row) => {
                if (row.rowId !== rowId) {
                    return row;
                }

                const currentCell = row.values[columnKey];

                return {
                    ...row,
                    values: {
                        ...row.values,
                        [columnKey]:
                            currentCell?.type === "image" && currentCell.src
                                ? {
                                      type: "image",
                                      src: currentCell.src,
                                      value,
                                  }
                                : {
                                      type: "text",
                                      value,
                                  },
                    },
                };
            }),
        }));
    };

    const onToggleRowEnabled = (rowId: string, enabled: boolean) => {
        patchActiveFile((file) => ({
            ...file,
            rows: file.rows.map((row) =>
                row.rowId === rowId ? { ...row, enabled } : row,
            ),
        }));
    };

    const updateRowAIResult = (
        fileId: string,
        rowId: string,
        stageKey: AIDetectStageKey,
        resultText: string,
    ) => {
        setFiles((previous) =>
            previous.map((file) => {
                if (file.fileId !== fileId) {
                    return file;
                }
                const nextRows = file.rows.map((row) => {
                    if (row.rowId !== rowId) {
                        return row;
                    }
                    return {
                        ...row,
                        aiResults: {
                            ...(row.aiResults ?? {}),
                            [stageKey]: resultText,
                        },
                    };
                });
                const nextFile: FileViewState = {
                    ...file,
                    rows: nextRows,
                };
                return nextFile;
            }),
        );

        cancelScheduledPersist(fileId);
        persistAIResults(
            fileId,
            stageKey,
            { [rowId]: resultText },
            latestFileStateRef.current[fileId],
        );
    };

    const updateRowCleaningResult = (
        fileId: string,
        rowId: string,
        toolKey: AICleaningToolKey,
        result: AICleaningToolResult,
        mappedFieldValues?: Record<string, string>,
    ) => {
        let nextFileForPersist: FileViewState | null = null;
        const normalizedMappedFieldValues = mappedFieldValues ?? {};

        setFiles((previous) =>
            previous.map((file) => {
                if (file.fileId !== fileId) {
                    return file;
                }
                const nextRows = file.rows.map((row) => {
                    if (row.rowId !== rowId) {
                        return row;
                    }
                    const nextValues = { ...row.values };
                    Object.entries(normalizedMappedFieldValues).forEach(
                        ([columnKey, value]) => {
                            const currentCell = row.values[columnKey];
                            nextValues[columnKey] =
                                currentCell?.type === "image" && currentCell.src
                                    ? {
                                          type: "image",
                                          src: currentCell.src,
                                          srcList: currentCell.srcList,
                                          value,
                                      }
                                    : {
                                          type: "text",
                                          value,
                                      };
                        },
                    );
                    return {
                        ...row,
                        values: nextValues,
                        cleaningResults: {
                            ...(row.cleaningResults ?? {}),
                            [toolKey]: result,
                        },
                    };
                });
                const nextFile: FileViewState = {
                    ...file,
                    rows: nextRows,
                };
                nextFileForPersist = nextFile;
                latestFileStateRef.current[fileId] = nextFile;
                return nextFile;
            }),
        );

        if (
            nextFileForPersist &&
            Object.keys(normalizedMappedFieldValues).length > 0
        ) {
            schedulePersistFileState(nextFileForPersist);
        }
    };

    const updateRowEvaluationResults = (
        fileId: string,
        rowId: string,
        taskId: string,
        results: AIEvaluationAttemptResult[],
    ) => {
        setFiles((previous) =>
            previous.map((file) => {
                if (file.fileId !== fileId) {
                    return file;
                }
                const nextRows = file.rows.map((row) =>
                    row.rowId === rowId
                        ? {
                              ...row,
                              evaluationResults: {
                                  ...(row.evaluationResults ?? {}),
                                  [taskId]: results,
                              },
                          }
                        : row,
                );
                const nextFile: FileViewState = {
                    ...file,
                    rows: nextRows,
                };
                latestFileStateRef.current[fileId] = nextFile;
                return nextFile;
            }),
        );
    };

    const persistColumnPrefs = (file: FileViewState) => {
        const prefsFileName = file.sourceFileName ?? file.fileName;
        fetch(`/api/column-prefs/${encodeURIComponent(prefsFileName)}`, {
            method: "PUT",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                fieldSignature: getFieldSignature(file.columns),
                displayKeys: file.selectedDisplayColumnKeys,
                editableKeys: file.selectedEditableColumnKeys,
            }),
        }).catch(() => {});
    };

    const resetPendingConfigState = () => {
        setPendingFile(null);
        setPendingSelectedDisplayKeys([]);
        setPendingEditableColumnKeys([]);
        setPendingConfigNotice("");
        setPendingConfigMode("import");
    };

    const onOpenActiveFileConfig = () => {
        if (!activeFile) {
            return;
        }
        setPendingFile(activeFile);
        setPendingSelectedDisplayKeys(activeFile.selectedDisplayColumnKeys);
        setPendingEditableColumnKeys(activeFile.selectedEditableColumnKeys);
        setPendingConfigNotice("");
        setPendingConfigMode("edit");
        navigateToSection("settings", "fields");
    };

    const onTogglePendingDisplayColumn = (columnKey: string) => {
        if (!pendingFile) {
            return;
        }
        setPendingSelectedDisplayKeys((previous) => {
            const shouldHide = previous.includes(columnKey);
            const next = shouldHide
                ? previous.filter((key) => key !== columnKey)
                : [...previous, columnKey];

            if (shouldHide) {
                setPendingEditableColumnKeys((editableKeys) =>
                    editableKeys.filter((key) => key !== columnKey),
                );
            }
            return next;
        });
    };

    const onTogglePendingEditableColumn = (columnKey: string) => {
        if (!pendingFile) {
            return;
        }
        setPendingEditableColumnKeys((previous) => {
            const exists = previous.includes(columnKey);
            const next = exists
                ? previous.filter((key) => key !== columnKey)
                : [...previous, columnKey];

            if (!exists) {
                setPendingSelectedDisplayKeys((displayKeys) =>
                    displayKeys.includes(columnKey)
                        ? displayKeys
                        : [...displayKeys, columnKey],
                );
            }

            return next;
        });
    };

    const onPendingSelectAllDisplayColumns = () => {
        if (!pendingFile) {
            return;
        }
        setPendingSelectedDisplayKeys(getAllColumnKeys(pendingFile.columns));
    };

    const onPendingClearDisplayColumns = () => {
        setPendingSelectedDisplayKeys([]);
        setPendingEditableColumnKeys([]);
    };

    const onPendingClearEditableColumns = () => {
        setPendingEditableColumnKeys([]);
    };

    const onCancelPendingFile = () => {
        resetPendingConfigState();
    };

    const onConfirmPendingFile = () => {
        if (!pendingFile) {
            return;
        }

        if (pendingConfigMode === "edit") {
            patchActiveFile((file) => {
                const nextFile = applyColumnConfigToFile(
                    file,
                    pendingSelectedDisplayKeys,
                    pendingEditableColumnKeys,
                );
                persistColumnPrefs(nextFile);
                return nextFile;
            });
            resetPendingConfigState();
            return;
        }

        const nextFile = toViewState(
            pendingFile,
            pendingSelectedDisplayKeys,
            pendingEditableColumnKeys,
        );
        upsertFileState(nextFile);
        setActiveFileId(nextFile.fileId);
        persistColumnPrefs(nextFile);
        persistFileState(nextFile);
        resetPendingConfigState();
    };

    const onToggleDisplayColumn = (columnKey: string) => {
        patchActiveFile((file) => {
            if (file.selectedEditableColumnKeys.includes(columnKey)) {
                return file;
            }

            const exists = file.selectedDisplayColumnKeys.includes(columnKey);
            const selectedDisplayColumnKeys = exists
                ? file.selectedDisplayColumnKeys.filter(
                      (key) => key !== columnKey,
                  )
                : [...file.selectedDisplayColumnKeys, columnKey];

            const normalized = normalizeColumnSelection(
                file.columns,
                selectedDisplayColumnKeys,
                file.selectedEditableColumnKeys,
            );
            const nextFile: FileViewState = {
                ...file,
                selectedDisplayColumnKeys: normalized.displayKeys,
                selectedEditableColumnKeys: normalized.editableKeys,
                filterConditions: file.filterConditions,
            };
            persistColumnPrefs(nextFile);
            return nextFile;
        });
    };

    const onUpdateFilterConditions = (filterConditions: FilterCondition[]) => {
        patchActiveFile((file) => {
            saveFilterConditionsToLocal(file.fileId, filterConditions);
            return {
                ...file,
                filterConditions,
            };
        });
    };

    const onClearFilterConditions = () => {
        patchActiveFile((file) => {
            if (file.filterConditions.length === 0) {
                return file;
            }
            saveFilterConditionsToLocal(file.fileId, []);
            return {
                ...file,
                filterConditions: [],
            };
        });
    };

    const onToggleStatisticsField = (fieldKey: string) => {
        patchActiveFile((file) => {
            const exists =
                file.statisticsConfig.selectedFieldKeys.includes(fieldKey);
            const selectedFieldKeys = exists
                ? file.statisticsConfig.selectedFieldKeys.filter(
                      (key) => key !== fieldKey,
                  )
                : [...file.statisticsConfig.selectedFieldKeys, fieldKey];
            return {
                ...file,
                statisticsConfig: normalizeStatisticsConfig(file.columns, {
                    ...file.statisticsConfig,
                    selectedFieldKeys,
                }),
            };
        });
    };

    const onSetStatisticsChartType = (
        fieldKey: string,
        chartType: StatisticsChartType,
    ) => {
        patchActiveFile((file) => ({
            ...file,
            statisticsConfig: normalizeStatisticsConfig(file.columns, {
                ...file.statisticsConfig,
                chartTypeByField: {
                    ...file.statisticsConfig.chartTypeByField,
                    [fieldKey]: chartType,
                },
            }),
        }));
    };

    const onExportFile = async () => {
        if (!activeFile) {
            return;
        }

        setIsExporting(true);
        setErrorMessage("");

        try {
            const responseState = await fetch(
                `/api/files/${encodeURIComponent(activeFile.fileId)}`,
            );
            if (!responseState.ok) {
                throw new Error("加载导出数据失败");
            }
            const statePayload = (await responseState.json()) as {
                file?: unknown;
            };
            const fullFile = normalizeLoadedFileState(statePayload.file);
            if (!fullFile) {
                throw new Error("导出数据无效");
            }

            const exportColumns = fullFile.columns;
            const headers = exportColumns.map((column) => column.title);
            const rows = fullFile.rows.map((row) =>
                exportColumns.map(
                    (column) => row.values[column.key]?.value ?? "",
                ),
            );

            const response = await fetch("/api/files/export", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    fileName: fullFile.fileName,
                    headers,
                    rows,
                }),
            });

            if (!response.ok) {
                const payload = (await response.json().catch(() => ({}))) as {
                    message?: string;
                };
                throw new Error(payload.message ?? "导出失败");
            }

            const blob = await response.blob();
            const headerFileName = getFileNameFromDisposition(
                response.headers.get("Content-Disposition"),
            );
            const fallbackFileName = `${activeFile.fileName.replace(/\.[^.]+$/, "")}-导出.xlsx`;
            downloadBlob(blob, headerFileName ?? fallbackFileName);
        } catch (error) {
            const message = error instanceof Error ? error.message : "导出失败";
            setErrorMessage(message);
        } finally {
            setIsExporting(false);
        }
    };

    const onExportEvaluation = async (rowIds?: string[]) => {
        if (!activeFile) {
            return;
        }

        setIsExportingEvaluation(true);
        setErrorMessage("");

        try {
            const response = await fetch(
                `/api/files/${encodeURIComponent(activeFile.fileId)}/export-evaluation`,
                {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ rowIds: rowIds ?? [] }),
                },
            );

            if (!response.ok) {
                const payload = (await response.json().catch(() => ({}))) as {
                    message?: string;
                };
                throw new Error(payload.message ?? "导出评测结果失败");
            }

            const blob = await response.blob();
            const headerFileName = getFileNameFromDisposition(
                response.headers.get("Content-Disposition"),
            );
            const fallbackFileName = `${activeFile.fileName.replace(/\.[^.]+$/, "")}-评测结果.json`;
            downloadBlob(blob, headerFileName ?? fallbackFileName);
        } catch (error) {
            const message =
                error instanceof Error ? error.message : "导出评测结果失败";
            setErrorMessage(message);
        } finally {
            setIsExportingEvaluation(false);
        }
    };

    const onOpenCreateProjectDialog = () => {
        setErrorMessage("");
        setProjectNameDialogError("");
        setProjectNameDraft("");
        setPendingDataSourceNameDraft("");
        setProjectNameTargetFileId(null);
        setProjectNameDialogMode("create");
    };

    const onOpenRenameProjectDialog = (fileId: string) => {
        const targetFile =
            files.find((file) => file.fileId === fileId) ?? activeFile;
        if (!targetFile) {
            return;
        }
        setErrorMessage("");
        setProjectNameDialogError("");
        setProjectNameDraft(targetFile.fileName);
        setPendingDataSourceNameDraft("");
        setProjectNameTargetFileId(targetFile.fileId);
        setProjectNameDialogMode("rename");
    };

    const onOpenAddDatasourceDialog = (fileId: string) => {
        const targetFile =
            files.find((file) => file.fileId === fileId) ?? activeFile;
        if (!targetFile) {
            return;
        }
        setErrorMessage("");
        setProjectNameDialogError("");
        setProjectNameDraft(targetFile.fileName);
        setPendingDataSourceNameDraft("");
        setProjectNameTargetFileId(fileId);
        setProjectNameDialogMode("add-datasource");
    };

    const onCancelProjectNameDialog = () => {
        resetProjectNameDialog();
    };

    const onStartMergeUpload = (fileId: string) => {
        pendingUploadModeRef.current = "merge";
        pendingMergeTargetFileIdRef.current = fileId;
        uploadInputRef.current?.click();
    };

    const onConfirmProjectNameDialog = async () => {
        const nextProjectName = projectNameDraft.trim();
        if (!nextProjectName) {
            const message = "项目名称不能为空";
            setProjectNameDialogError(message);
            setErrorMessage(message);
            return;
        }

        setProjectNameDialogError("");
        setErrorMessage("");

        if (projectNameDialogMode === "create") {
            pendingCreateProjectNameRef.current = nextProjectName;
            pendingDataSourceNameRef.current =
                pendingDataSourceNameDraft.trim();
            pendingDataSourceGroupIdRef.current = null;
            pendingUploadModeRef.current = "create";
            pendingMergeTargetFileIdRef.current = null;
            resetProjectNameDialog();
            uploadInputRef.current?.click();
            return;
        }

        if (projectNameDialogMode === "add-datasource") {
            const targetFileId = projectNameTargetFileId;
            const targetFile = targetFileId
                ? (files.find((f) => f.fileId === targetFileId) ?? activeFile)
                : activeFile;
            if (!targetFile) {
                return;
            }
            const projectGroupId = targetFile.projectId ?? targetFile.fileId;
            pendingCreateProjectNameRef.current = targetFile.fileName;
            pendingDataSourceNameRef.current =
                pendingDataSourceNameDraft.trim();
            pendingDataSourceGroupIdRef.current = projectGroupId;
            pendingUploadModeRef.current = "add-datasource";
            pendingMergeTargetFileIdRef.current = null;
            resetProjectNameDialog();
            uploadInputRef.current?.click();
            return;
        }

        if (projectNameDialogMode !== "rename" || !projectNameTargetFileId) {
            return;
        }

        try {
            const response = await fetch(
                `/api/files/${encodeURIComponent(projectNameTargetFileId)}/name`,
                {
                    method: "PUT",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ fileName: nextProjectName }),
                },
            );
            if (!response.ok) {
                const payload = (await response.json().catch(() => ({}))) as {
                    message?: string;
                };
                throw new Error(payload.message ?? "项目重命名失败");
            }

            const payload = (await response.json()) as {
                file?: unknown;
                files?: unknown[];
            };
            const allRaw =
                Array.isArray(payload.files) && payload.files.length > 0
                    ? payload.files
                    : [payload.file];
            const renamedFiles = allRaw
                .map((item) => normalizeLoadedFileState(item))
                .filter((f): f is FileViewState => f !== null);
            if (renamedFiles.length === 0) {
                throw new Error("项目重命名结果无效");
            }

            renamedFiles.forEach(upsertFileState);
            const primaryFile =
                renamedFiles.find(
                    (f) => f.fileId === projectNameTargetFileId,
                ) ?? renamedFiles[0];
            if (pendingFile?.fileId === primaryFile.fileId) {
                setPendingFile(primaryFile);
            }
            setActiveFileId(primaryFile.fileId);
            resetProjectNameDialog();
        } catch (error) {
            const message =
                error instanceof Error ? error.message : "项目重命名失败";
            setProjectNameDialogError(message);
            setErrorMessage(message);
        }
    };

    const onUploadFile = async (event: React.ChangeEvent<HTMLInputElement>) => {
        const selected = event.target.files?.[0];
        if (!selected) {
            return;
        }

        setIsUploading(true);
        setErrorMessage("");

        try {
            const uploadMode = pendingUploadModeRef.current;
            const mergeTargetFileId =
                pendingMergeTargetFileIdRef.current ??
                activeFile?.fileId ??
                null;
            const mergeTargetFile =
                files.find((file) => file.fileId === mergeTargetFileId) ??
                activeFile;

            if (uploadMode === "merge" && !mergeTargetFile) {
                throw new Error("请先选择要更新的项目");
            }
            if (
                uploadMode === "create" &&
                pendingCreateProjectNameRef.current.trim().length === 0
            ) {
                throw new Error("缺少项目名称");
            }
            if (
                uploadMode === "add-datasource" &&
                !pendingDataSourceGroupIdRef.current
            ) {
                throw new Error("缺少项目标识");
            }

            if (uploadMode === "merge" && mergeTargetFile) {
                cancelScheduledPersist(mergeTargetFile.fileId);
                await persistFileState(mergeTargetFile);
                await flushPendingAIResults(mergeTargetFile.fileId);
            }

            const formData = new FormData();
            formData.append("file", selected);
            formData.append("mode", uploadMode);
            if (uploadMode === "create") {
                formData.append(
                    "projectName",
                    pendingCreateProjectNameRef.current,
                );
                if (pendingDataSourceNameRef.current) {
                    formData.append(
                        "dataSourceName",
                        pendingDataSourceNameRef.current,
                    );
                }
            }
            if (uploadMode === "merge" && mergeTargetFile) {
                formData.append("targetFileId", mergeTargetFile.fileId);
            }
            if (uploadMode === "add-datasource") {
                formData.append(
                    "projectName",
                    pendingCreateProjectNameRef.current,
                );
                formData.append(
                    "dataSourceGroupId",
                    pendingDataSourceGroupIdRef.current ?? "",
                );
                if (pendingDataSourceNameRef.current) {
                    formData.append(
                        "dataSourceName",
                        pendingDataSourceNameRef.current,
                    );
                }
            }
            const response = await fetch("/api/files/upload", {
                method: "POST",
                body: formData,
            });

            if (!response.ok) {
                const payload = (await response.json().catch(() => ({}))) as {
                    message?: string;
                };
                throw new Error(payload.message ?? "导入文件失败");
            }

            const payload = (await response.json()) as {
                file?: ParsedFile;
                summary?: {
                    insertedCount?: number;
                    updatedCount?: number;
                    totalRows?: number;
                };
            };
            const parsed = payload.file ?? (payload as ParsedFile);
            if (uploadMode === "merge" || uploadMode === "add-datasource") {
                const mergedFile = normalizeLoadedFileState(parsed);
                if (!mergedFile) {
                    throw new Error("项目导入结果无效");
                }
                upsertFileState(mergedFile);
                setActiveFileId(mergedFile.fileId);
                resetPendingConfigState();
                return;
            }
            let parsedImageCellCount = 0;
            let parsedTextLikeImageCellCount = 0;
            const textLikeSamples: string[] = [];
            parsed.rows.forEach((row) => {
                parsed.columns.forEach((column) => {
                    const cell = row.values[column.key];
                    if (!cell) {
                        return;
                    }
                    if (
                        cell.type === "image" &&
                        typeof cell.src === "string" &&
                        cell.src
                    ) {
                        parsedImageCellCount += 1;
                        return;
                    }
                    if (
                        cell.type === "text" &&
                        typeof cell.value === "string" &&
                        /\.(png|jpe?g|webp|gif|bmp|tiff?)([?#].*)?$/i.test(
                            cell.value.trim(),
                        )
                    ) {
                        parsedTextLikeImageCellCount += 1;
                        if (textLikeSamples.length < 8) {
                            textLikeSamples.push(
                                `row=${row.rowId} column=${column.title} value=${cell.value.trim()}`,
                            );
                        }
                    }
                });
            });
            // eslint-disable-next-line no-console
            console.log(
                `[UIParsedImage] file=${parsed.fileName} imageCells=${parsedImageCellCount} textLikeImageCells=${parsedTextLikeImageCellCount}`,
            );
            if (textLikeSamples.length > 0) {
                // eslint-disable-next-line no-console
                console.log(
                    `[UIParsedImageTextLike] ${JSON.stringify(textLikeSamples)}`,
                );
            }
            const defaultDisplayKeys = getAllColumnKeys(parsed.columns);
            let initialDisplayKeys = defaultDisplayKeys;
            let initialEditableKeys: string[] = [];
            let shouldShowColumnModal = true;
            let nextPendingNotice = "";
            const prefsFileName = parsed.sourceFileName ?? parsed.fileName;

            try {
                const prefsRes = await fetch(
                    `/api/column-prefs/${encodeURIComponent(prefsFileName)}`,
                );
                if (prefsRes.ok) {
                    const prefsData = (await prefsRes.json()) as {
                        config: ColumnPrefsConfig | null;
                    };
                    if (prefsData.config) {
                        const normalizedSaved = normalizeColumnSelection(
                            parsed.columns,
                            prefsData.config.displayKeys,
                            prefsData.config.editableKeys,
                        );
                        const currentSignature = getFieldSignature(
                            parsed.columns,
                        );
                        if (
                            prefsData.config.fieldSignature === currentSignature
                        ) {
                            const nextFile = toViewState(
                                parsed,
                                normalizedSaved.displayKeys,
                                normalizedSaved.editableKeys,
                            );
                            upsertFileState(nextFile);
                            setActiveFileId(nextFile.fileId);
                            persistFileState(nextFile);
                            shouldShowColumnModal = false;
                        } else {
                            nextPendingNotice =
                                "检测到该 Excel 字段与已保存配置不一致，请重新选择并保存新配置。";
                            initialDisplayKeys = normalizedSaved.displayKeys;
                            initialEditableKeys = normalizedSaved.editableKeys;
                        }
                    }
                }
            } catch {
                // Ignore and fall back to default selection
            }

            if (shouldShowColumnModal) {
                setPendingFile(parsed);
                setPendingSelectedDisplayKeys(initialDisplayKeys);
                setPendingEditableColumnKeys(initialEditableKeys);
                setPendingConfigNotice(nextPendingNotice);
                setPendingConfigMode("import");
            }
        } catch (error) {
            const message = error instanceof Error ? error.message : "上传失败";
            setErrorMessage(message);
        } finally {
            pendingCreateProjectNameRef.current = "";
            pendingDataSourceNameRef.current = "";
            pendingDataSourceGroupIdRef.current = null;
            pendingMergeTargetFileIdRef.current = null;
            setIsUploading(false);
            event.target.value = "";
        }
    };

    const onRenameDataSource = async (
        fileId: string,
        newName: string,
    ): Promise<void> => {
        try {
            const response = await fetch(
                `/api/files/${encodeURIComponent(fileId)}/datasource-name`,
                {
                    method: "PUT",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ dataSourceName: newName }),
                },
            );
            if (!response.ok) {
                const payload = (await response.json().catch(() => ({}))) as {
                    message?: string;
                };
                throw new Error(payload.message ?? "保存数据源名称失败");
            }
            const payload = (await response.json()) as { file?: unknown };
            const updated = normalizeLoadedFileState(payload.file);
            if (updated) {
                upsertFileState(updated);
            }
        } catch (error) {
            const message =
                error instanceof Error ? error.message : "保存数据源名称失败";
            setErrorMessage(message);
        }
    };

    const onRemoveFile = async (fileId: string) => {
        const targetFile = files.find((file) => file.fileId === fileId);
        if (!targetFile) {
            return;
        }
        const confirmed =
            typeof window === "undefined"
                ? true
                : window.confirm(
                      `确定删除项目“${targetFile.fileName}”吗？此操作会同时删除该项目的本地状态和 AI 清洗结果，且不可撤销。`,
                  );
        if (!confirmed) {
            return;
        }

        setErrorMessage("");
        setRemovingFileId(fileId);

        try {
            const response = await fetch(
                `/api/files/${encodeURIComponent(fileId)}`,
                {
                    method: "DELETE",
                },
            );
            if (!response.ok) {
                const payload = (await response.json().catch(() => ({}))) as {
                    message?: string;
                };
                throw new Error(payload.message ?? "删除项目失败");
            }

            cancelScheduledPersist(fileId);
            delete pendingPersistRef.current[fileId];
            delete persistTimersRef.current[fileId];
            delete latestFileStateRef.current[fileId];
            delete persistAIResultsQueueRef.current[fileId];
            delete stateVersionRef.current[fileId];

            setFiles((previous) => {
                const next = previous.filter((file) => file.fileId !== fileId);
                if (activeFileId === fileId) {
                    setActiveFileId(next[0]?.fileId ?? null);
                }
                return next;
            });
        } catch (error) {
            const message =
                error instanceof Error ? error.message : "删除项目失败";
            setErrorMessage(message);
        } finally {
            setRemovingFileId((previous) =>
                previous === fileId ? null : previous,
            );
        }
    };

    return {
        files,
        setFiles,
        activeFileId,
        setActiveFileId,
        activeFile,
        isUploading,
        isExporting,
        isExportingEvaluation,
        uploadInputRef,
        pendingFile,
        pendingSelectedDisplayKeys,
        pendingEditableColumnKeys,
        pendingConfigNotice,
        pendingConfigMode,
        initialLoadComplete,
        projectNameDialogMode,
        projectNameDraft,
        setProjectNameDraft,
        projectNameDialogError,
        projectNameTargetFileId,
        pendingDataSourceNameDraft,
        setPendingDataSourceNameDraft,
        removingFileId,
        persistFileState,
        schedulePersistFileState,
        cancelScheduledPersist,
        persistAIResults,
        flushPendingAIResults,
        latestFileStateRef,
        updateRowAIResult,
        updateRowCleaningResult,
        updateRowEvaluationResults,
        onEditCell,
        onToggleRowEnabled,
        onToggleDisplayColumn,
        onUpdateFilterConditions,
        onClearFilterConditions,
        onToggleStatisticsField,
        onSetStatisticsChartType,
        onOpenActiveFileConfig,
        onTogglePendingDisplayColumn,
        onTogglePendingEditableColumn,
        onPendingSelectAllDisplayColumns,
        onPendingClearDisplayColumns,
        onPendingClearEditableColumns,
        onCancelPendingFile,
        onConfirmPendingFile,
        onOpenCreateProjectDialog,
        onOpenRenameProjectDialog,
        onOpenAddDatasourceDialog,
        onCancelProjectNameDialog,
        onConfirmProjectNameDialog,
        onStartMergeUpload,
        onUploadFile,
        onExportFile,
        onExportEvaluation,
        onRenameDataSource,
        onRemoveFile,
    };
};
