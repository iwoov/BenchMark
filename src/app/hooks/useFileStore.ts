import { useEffect, useMemo, useRef, useState } from "react";
import type {
    AICleaningToolKey,
    AICleaningToolResult,
    AIDetectStageKey,
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
type UploadMode = "create" | "merge";

export const useFileStore = ({
    navigateToSection,
    setErrorMessage,
}: {
    navigateToSection: NavigateToSection;
    setErrorMessage: (value: string) => void;
}) => {
    const [files, setFiles] = useState<FileViewState[]>([]);
    const [activeFileId, setActiveFileId] = useState<string | null>(null);
    const [isUploading, setIsUploading] = useState(false);
    const [isExporting, setIsExporting] = useState(false);
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

    const uploadInputRef = useRef<HTMLInputElement>(null);
    const persistTimersRef = useRef<Record<string, number>>({});
    const pendingPersistRef = useRef<Record<string, FileViewState>>({});
    const persistAIResultsQueueRef = useRef<Record<string, Promise<void>>>({});
    const latestFileStateRef = useRef<Record<string, FileViewState>>({});
    const stateVersionRef = useRef<Record<string, number>>({});
    const pendingUploadModeRef = useRef<UploadMode>("create");

    const activeFile = useMemo(
        () => files.find((item) => item.fileId === activeFileId) ?? null,
        [files, activeFileId],
    );

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
                    .filter((item): item is FileViewState => item !== null);

                if (shouldLogAIResults && restoredFiles.length > 0) {
                    logAIResultsSummary(restoredFiles);
                }

                if (disposed) {
                    return;
                }

                if (restoredFiles.length === 0) {
                    setInitialLoadComplete(true);
                    return;
                }

                setFiles((previous) => {
                    if (previous.length === 0) {
                        return restoredFiles;
                    }

                    const merged = new Map<string, FileViewState>();
                    restoredFiles.forEach((file) =>
                        merged.set(file.fileId, file),
                    );
                    previous.forEach((file) => merged.set(file.fileId, file));
                    return Array.from(merged.values());
                });
                setActiveFileId(
                    (previous) => previous ?? restoredFiles[0].fileId,
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
        const sanitizedState: FileViewState = {
            ...file,
            rows: file.rows.map(({ cleaningResults, ...row }) => row),
        };
        return {
            ...sanitizedState,
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
                    body: JSON.stringify({ state: statePayload }),
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
        setFiles((previous) => [...previous, nextFile]);
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
        patchActiveFile((file) => ({
            ...file,
            filterConditions,
        }));
    };

    const onClearFilterConditions = () => {
        patchActiveFile((file) =>
            file.filterConditions.length === 0
                ? file
                : {
                      ...file,
                      filterConditions: [],
                  },
        );
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
            const exportColumns = activeFile.columns;
            const headers = exportColumns.map((column) => column.title);
            const rows = activeFile.rows.map((row) =>
                exportColumns.map(
                    (column) => row.values[column.key]?.value ?? "",
                ),
            );

            const response = await fetch("/api/files/export", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    fileName: activeFile.fileName,
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

    const onUploadClick = (mode: UploadMode = "create") => {
        pendingUploadModeRef.current = mode;
        uploadInputRef.current?.click();
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
            if (uploadMode === "merge" && !activeFile) {
                throw new Error("请先选择要更新的项目");
            }

            if (uploadMode === "merge" && activeFile) {
                cancelScheduledPersist(activeFile.fileId);
                await persistFileState(activeFile);
                await flushPendingAIResults(activeFile.fileId);
            }

            const formData = new FormData();
            formData.append("file", selected);
            formData.append("mode", uploadMode);
            if (uploadMode === "merge" && activeFile) {
                formData.append("targetFileId", activeFile.fileId);
            }
            const response = await fetch("/api/files/upload", {
                method: "POST",
                body: formData,
            });

            if (!response.ok) {
                const payload = (await response.json().catch(() => ({}))) as {
                    message?: string;
                };
                throw new Error(payload.message ?? "文件解析失败");
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
            if (uploadMode === "merge") {
                const mergedFile = normalizeLoadedFileState(parsed);
                if (!mergedFile) {
                    throw new Error("项目导入结果无效");
                }
                setFiles((previous) => {
                    const exists = previous.some(
                        (file) => file.fileId === mergedFile.fileId,
                    );
                    if (!exists) {
                        return [...previous, mergedFile];
                    }
                    return previous.map((file) =>
                        file.fileId === mergedFile.fileId ? mergedFile : file,
                    );
                });
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
                            setFiles((previous) => [...previous, nextFile]);
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
            setIsUploading(false);
            event.target.value = "";
        }
    };

    const onRemoveFile = (fileId: string) => {
        cancelScheduledPersist(fileId);
        fetch(`/api/files/${encodeURIComponent(fileId)}`, {
            method: "DELETE",
        }).catch(() => {});
        setFiles((previous) => {
            const next = previous.filter((file) => file.fileId !== fileId);
            if (activeFileId === fileId) {
                setActiveFileId(next[0]?.fileId ?? null);
            }
            return next;
        });
    };

    return {
        files,
        setFiles,
        activeFileId,
        setActiveFileId,
        activeFile,
        isUploading,
        isExporting,
        uploadInputRef,
        pendingFile,
        pendingSelectedDisplayKeys,
        pendingEditableColumnKeys,
        pendingConfigNotice,
        pendingConfigMode,
        initialLoadComplete,
        persistFileState,
        schedulePersistFileState,
        cancelScheduledPersist,
        persistAIResults,
        flushPendingAIResults,
        latestFileStateRef,
        updateRowAIResult,
        updateRowCleaningResult,
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
        onUploadClick,
        onUploadFile,
        onExportFile,
        onRemoveFile,
    };
};
