import { useEffect, useMemo, useRef, useState } from "react";
import type {
  AIDetectConfig,
  AIDetectStageKey,
  FileViewState,
  NamedAIDetectConfig,
  ParsedCell,
  ParsedColumn,
  ParsedFile,
  ParsedRow,
} from "./types";
import {
  ALL_FILTER_VALUE,
  AI_STAGE_LABELS,
  AI_STAGE_ORDER,
  DEFAULT_AI_BATCH_CONCURRENCY,
  DEFAULT_AI_CONFIG_NAME,
  INITIAL_AI_BATCH_TASK,
  LIST_PAGE_SIZE_OPTIONS,
  MAX_AI_BATCH_CONCURRENCY,
  MIN_AI_BATCH_CONCURRENCY,
} from "./app/constants";
import {
  applyColumnConfigToFile,
  downloadBlob,
  getAllColumnKeys,
  getCellText,
  getDistinctOptions,
  getFieldSignature,
  getFileNameFromDisposition,
  isFeedbackColumnTitle,
  isInspectorColumnTitle,
  isOpensourceColumnTitle,
  isQualifiedColumnTitle,
  logUIImageRenderError,
  normalizeFilterSelection,
  normalizeColumnSelection,
  normalizeLoadedFileState,
  toViewState,
} from "./app/file-helpers";
import {
  buildAIDetectFieldsForRow,
  cloneAIDetectConfig,
  composeAISaveText,
  createDefaultAIDetectConfig,
  formatDuration,
  normalizeAIBatchConcurrency,
  normalizeAIDetectConfigForColumns,
  normalizeLoadedAIDetectConfig,
  normalizeAIConfigName,
  normalizeLoadedNamedAIDetectConfigs,
  normalizeNamedAIDetectConfigsForColumns,
  pickAIConfigName,
  requestAIDetectResult,
} from "./app/ai-helpers";
import { IconFile } from "./app/icons";
import {
  LatexRenderer,
  hasLatexSyntax,
  shouldAutoDisplayLatex,
} from "./app/latex";
import { buildHashRoute, parseHashRoute } from "./app/routes";
import type {
  AIBatchTaskState,
  ColumnPrefsConfig,
  MainSection,
  RouteState,
  SettingsSection,
} from "./app/types";
import { HeaderBar } from "./app/components/HeaderBar";
import { WorkspaceSidebar } from "./app/components/WorkspaceSidebar";
import { ListPage } from "./app/components/ListPage";
import { DetailPage } from "./app/components/DetailPage";
import { SettingsPage } from "./app/components/SettingsPage";
import { ColumnConfigModal } from "./app/components/ColumnConfigModal";
import { AIProfileModal } from "./app/components/AIProfileModal";
import { AIStageConfigModal } from "./app/components/AIStageConfigModal";
import { AIRunModal } from "./app/components/AIRunModal";
import { ImageLightbox } from "./app/components/ImageLightbox";

function App() {
  type PendingConfigMode = "import" | "edit";
  const initialRoute: RouteState =
    typeof window !== "undefined"
      ? parseHashRoute(window.location.hash)
      : { section: "list", settingsSection: "fields" };
  const [files, setFiles] = useState<FileViewState[]>([]);
  const [activeFileId, setActiveFileId] = useState<string | null>(null);
  const [isUploading, setIsUploading] = useState(false);
  const [isExporting, setIsExporting] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string>("");
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
  const [selectedAIConfigName, setSelectedAIConfigName] = useState<string>(
    DEFAULT_AI_CONFIG_NAME,
  );
  const [draftAIConfigName, setDraftAIConfigName] = useState<string>(
    DEFAULT_AI_CONFIG_NAME,
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
  const [activeAIStageKey, setActiveAIStageKey] =
    useState<AIDetectStageKey>("precheck");
  const [aiDetectElapsedMs, setAIDetectElapsedMs] = useState(0);
  const [isAIRunModalOpen, setIsAIRunModalOpen] = useState(false);
  const [aiBatchTask, setAIBatchTask] = useState<AIBatchTaskState>(
    INITIAL_AI_BATCH_TASK,
  );
  const [aiBatchConcurrency, setAIBatchConcurrency] = useState<number>(
    DEFAULT_AI_BATCH_CONCURRENCY,
  );
  const [selectedRowId, setSelectedRowId] = useState<string | null>(null);
  const [batchSelectedRowIds, setBatchSelectedRowIds] = useState<string[]>([]);
  const [pendingFile, setPendingFile] = useState<ParsedFile | null>(null);
  const [pendingSelectedDisplayKeys, setPendingSelectedDisplayKeys] = useState<
    string[]
  >([]);
  const [pendingSelectedFilterKeys, setPendingSelectedFilterKeys] = useState<
    string[]
  >([]);
  const [pendingEditableColumnKeys, setPendingEditableColumnKeys] = useState<
    string[]
  >([]);
  const [pendingConfigNotice, setPendingConfigNotice] = useState<string>("");
  const [pendingConfigMode, setPendingConfigMode] =
    useState<PendingConfigMode>("import");
  const [showHiddenFields, setShowHiddenFields] = useState(false);
  const [activeSection, setActiveSection] = useState<MainSection>(
    initialRoute.section,
  );
  const [activeSettingsSection, setActiveSettingsSection] =
    useState<SettingsSection>(initialRoute.settingsSection);
  const [isSidebarCollapsed, setIsSidebarCollapsed] = useState(false);
  const [listPage, setListPage] = useState(1);
  const [listPageSize, setListPageSize] = useState<number>(
    LIST_PAGE_SIZE_OPTIONS[1],
  );
  const [latexRenderOverrides, setLatexRenderOverrides] = useState<
    Record<string, boolean>
  >({});
  const [previewImageSrc, setPreviewImageSrc] = useState<string | null>(null);
  const [theme, setTheme] = useState<"dark" | "light">(() => {
    if (typeof window !== "undefined") {
      return (localStorage.getItem("theme") as "dark" | "light") || "dark";
    }
    return "dark";
  });
  const uploadInputRef = useRef<HTMLInputElement>(null);
  const persistTimersRef = useRef<Record<string, number>>({});
  const pendingPersistRef = useRef<Record<string, FileViewState>>({});
  const aiStreamAbortRef = useRef<AbortController | null>(null);
  const aiBatchAbortRef = useRef<AbortController | null>(null);
  const aiDetectStartedAtRef = useRef<number | null>(null);

  useEffect(() => {
    document.documentElement.setAttribute("data-theme", theme);
    localStorage.setItem("theme", theme);
  }, [theme]);

  useEffect(() => {
    let disposed = false;
    const controller = new AbortController();

    const loadPersistedFiles = async () => {
      try {
        const response = await fetch("/api/files", {
          signal: controller.signal,
        });
        if (!response.ok) {
          return;
        }

        const payload = (await response.json()) as { files?: unknown };
        const rawFiles = Array.isArray(payload.files) ? payload.files : [];
        const restoredFiles = rawFiles
          .map((item) => normalizeLoadedFileState(item))
          .filter((item): item is FileViewState => item !== null);

        if (disposed || restoredFiles.length === 0) {
          return;
        }

        setFiles((previous) => {
          if (previous.length === 0) {
            return restoredFiles;
          }

          const merged = new Map<string, FileViewState>();
          restoredFiles.forEach((file) => merged.set(file.fileId, file));
          previous.forEach((file) => merged.set(file.fileId, file));
          return Array.from(merged.values());
        });
        setActiveFileId((previous) => previous ?? restoredFiles[0].fileId);
      } catch {
        // Ignore load errors and keep empty startup state.
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
      aiStreamAbortRef.current?.abort();
      aiStreamAbortRef.current = null;
      aiBatchAbortRef.current?.abort();
      aiBatchAbortRef.current = null;
      Object.values(persistTimersRef.current).forEach((timerId) => {
        window.clearTimeout(timerId);
      });
      persistTimersRef.current = {};
      pendingPersistRef.current = {};
    };
  }, []);

  const toggleTheme = () => {
    setTheme((prev) => (prev === "dark" ? "light" : "dark"));
  };

  const navigateToSection = (
    section: MainSection,
    settingsSection: SettingsSection = activeSettingsSection,
    options?: { replace?: boolean },
  ) => {
    const nextHash = buildHashRoute(section, settingsSection);
    if (typeof window !== "undefined" && window.location.hash !== nextHash) {
      if (options?.replace) {
        window.history.replaceState(null, "", nextHash);
        setActiveSection(section);
        setActiveSettingsSection(settingsSection);
      } else {
        window.location.hash = nextHash;
      }
      return;
    }

    setActiveSection(section);
    setActiveSettingsSection(settingsSection);
  };

  const activeFile = useMemo(
    () => files.find((item) => item.fileId === activeFileId) ?? null,
    [files, activeFileId],
  );

  useEffect(() => {
    if (typeof window === "undefined") {
      return;
    }

    const syncRouteState = () => {
      const nextRoute = parseHashRoute(window.location.hash);
      setActiveSection(nextRoute.section);
      setActiveSettingsSection(nextRoute.settingsSection);
    };

    window.addEventListener("hashchange", syncRouteState);
    syncRouteState();

    return () => {
      window.removeEventListener("hashchange", syncRouteState);
    };
  }, []);

  useEffect(() => {
    if (!activeFile) {
      navigateToSection("list", activeSettingsSection, { replace: true });
    }
  }, [activeFile]);

  useEffect(() => {
    if (!activeFile) {
      const nextConfig = createDefaultAIDetectConfig();
      setAIConfigList([
        {
          name: DEFAULT_AI_CONFIG_NAME,
          config: nextConfig,
        },
      ]);
      setSelectedAIConfigName(DEFAULT_AI_CONFIG_NAME);
      setDraftAIConfigName(DEFAULT_AI_CONFIG_NAME);
      setAIConfig(nextConfig);
      setDraftAIConfig(cloneAIDetectConfig(nextConfig));
      setAIConfigFormMessage("");
      setAIThinkingText("");
      setAIResultText("");
      setAIResultMessage("");
      setBatchSelectedRowIds([]);
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
              config: normalizeLoadedAIDetectConfig(payload.config),
            },
          ];
        }

        const normalizedConfigs = normalizeNamedAIDetectConfigsForColumns(
          loadedConfigs,
          activeFile.columns,
        );
        const activeConfigName = pickAIConfigName(
          normalizedConfigs,
          payload.activeConfigName,
        );
        const activeConfig =
          normalizedConfigs.find((item) => item.name === activeConfigName)
            ?.config ?? normalizedConfigs[0].config;

        if (disposed) {
          return;
        }
        setAIConfigList(normalizedConfigs);
        setSelectedAIConfigName(activeConfigName);
        setDraftAIConfigName(activeConfigName);
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
        setSelectedAIConfigName(DEFAULT_AI_CONFIG_NAME);
        setDraftAIConfigName(DEFAULT_AI_CONFIG_NAME);
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
    const nextConfigList =
      normalizedConfigs.length > 0
        ? normalizedConfigs
        : [
            {
              name: DEFAULT_AI_CONFIG_NAME,
              config: normalizeAIDetectConfigForColumns(
                createDefaultAIDetectConfig(),
                activeFile.columns,
              ),
            },
          ];
    const nextSelectedName = pickAIConfigName(
      nextConfigList,
      selectedAIConfigName,
    );
    const nextSelectedConfig =
      nextConfigList.find((item) => item.name === nextSelectedName)?.config ??
      nextConfigList[0].config;

    setAIConfigList(nextConfigList);
    setSelectedAIConfigName(nextSelectedName);
    setAIConfig(nextSelectedConfig);
    setDraftAIConfigName((previous) =>
      nextConfigList.some((item) => item.name === previous)
        ? previous
        : nextSelectedName,
    );
    setDraftAIConfig((previous) =>
      normalizeAIDetectConfigForColumns(previous, activeFile.columns),
    );
  }, [activeFile?.fileId, activeFile?.columns]);

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

  const filterColumns = useMemo(() => {
    if (!activeFile) {
      return [];
    }
    return activeFile.selectedFilterColumnKeys
      .map((key) => activeFile.columns.find((column) => column.key === key))
      .filter((column): column is ParsedColumn => Boolean(column));
  }, [activeFile]);
  const filterOptionsMap = useMemo(() => {
    if (!activeFile) {
      return new Map<string, string[]>();
    }
    const map = new Map<string, string[]>();
    filterColumns.forEach((column) => {
      map.set(column.key, getDistinctOptions(activeFile.rows, column.key));
    });
    return map;
  }, [activeFile, filterColumns]);

  const displayColumns = useMemo(() => {
    if (!activeFile) {
      return [];
    }
    return activeFile.columns.filter((column) => {
      return activeFile.selectedDisplayColumnKeys.includes(column.key);
    });
  }, [activeFile]);

  const hiddenColumns = useMemo(() => {
    if (!activeFile) {
      return [];
    }
    return activeFile.columns.filter((column) => {
      return !activeFile.selectedDisplayColumnKeys.includes(column.key);
    });
  }, [activeFile]);

  const aiSubmitFieldColumns = useMemo(
    () => (activeFile ? activeFile.columns : []),
    [activeFile],
  );

  const activeStageConfig = useMemo(
    () => aiConfig.stages[activeAIStageKey],
    [aiConfig, activeAIStageKey],
  );
  const activeProfile = useMemo(() => {
    const profiles = aiConfig.profiles ?? [];
    const matched = profiles.find(
      (item) => item.name === activeStageConfig?.profileName,
    );
    return matched ?? profiles[0] ?? null;
  }, [aiConfig.profiles, activeStageConfig?.profileName]);

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
  const batchSelectedRowIdSet = useMemo(
    () => new Set(batchSelectedRowIds),
    [batchSelectedRowIds],
  );

  const visibleRows = useMemo(() => {
    if (!activeFile) {
      return [];
    }

    return activeFile.rows.filter((row) => {
      for (const column of filterColumns) {
        const filterValue =
          activeFile.columnFilterValues[column.key] ?? ALL_FILTER_VALUE;
        if (filterValue === ALL_FILTER_VALUE) {
          continue;
        }
        const value = getCellText(row, column.key).trim();
        if (value !== filterValue) {
          return false;
        }
      }
      return true;
    });
  }, [activeFile, filterColumns]);

  const totalListPages = Math.max(
    1,
    Math.ceil(visibleRows.length / listPageSize) || 1,
  );
  const paginatedRows = useMemo(() => {
    const start = (listPage - 1) * listPageSize;
    return visibleRows.slice(start, start + listPageSize);
  }, [visibleRows, listPage, listPageSize]);

  useEffect(() => {
    setListPage(1);
  }, [
    activeFile?.fileId,
    activeFile?.selectedFilterColumnKeys,
    activeFile?.columnFilterValues,
    listPageSize,
  ]);

  useEffect(() => {
    if (listPage > totalListPages) {
      setListPage(totalListPages);
    }
  }, [listPage, totalListPages]);

  useEffect(() => {
    if (!activeFile || visibleRows.length === 0) {
      setSelectedRowId(null);
      return;
    }

    if (
      selectedRowId !== null &&
      !visibleRows.some((row) => row.rowId === selectedRowId)
    ) {
      setSelectedRowId(null);
    }
  }, [activeFile, visibleRows, selectedRowId]);

  const selectedRow = useMemo(
    () => visibleRows.find((row) => row.rowId === selectedRowId) ?? null,
    [visibleRows, selectedRowId],
  );
  const aiRequestPreview = useMemo(() => {
    if (!activeFile || !selectedRow) {
      return "";
    }
    const fields = buildAIDetectFieldsForRow(
      activeFile.columns,
      selectedRow,
      activeStageConfig?.submitFieldKeys ?? [],
    );
    return JSON.stringify(
      {
        configName: selectedAIConfigName,
        stageKey: activeAIStageKey,
        stageTitle: AI_STAGE_LABELS[activeAIStageKey]?.shortTitle ?? "",
        profileName: activeProfile?.name ?? "",
        provider: activeProfile?.profile.provider,
        model: activeProfile?.profile.model,
        reasoningEffort: activeProfile?.profile.reasoningEffort,
        retryCount: activeProfile?.profile.retryCount,
        prompt: activeStageConfig?.prompt,
        fields,
      },
      null,
      2,
    );
  }, [
    activeFile,
    selectedRow,
    activeStageConfig,
    activeAIStageKey,
    activeProfile,
    selectedAIConfigName,
  ]);

  const openRowDetail = (rowId: string) => {
    setSelectedRowId(rowId);
    navigateToSection("detail");
  };

  const activeRowIndex = selectedRow
    ? visibleRows.findIndex((row) => row.rowId === selectedRow.rowId)
    : -1;
  const previousRow =
    activeRowIndex > 0 ? visibleRows[activeRowIndex - 1] : null;
  const nextRow =
    activeRowIndex >= 0 && activeRowIndex < visibleRows.length - 1
      ? visibleRows[activeRowIndex + 1]
      : null;

  useEffect(() => {
    if (!activeFile) {
      setBatchSelectedRowIds([]);
      return;
    }

    const visibleIdSet = new Set(visibleRows.map((row) => row.rowId));
    setBatchSelectedRowIds((previous) =>
      previous.filter((rowId) => visibleIdSet.has(rowId)),
    );
  }, [activeFile?.fileId, visibleRows]);

  const persistFileState = (file: FileViewState) => {
    fetch(`/api/files/${encodeURIComponent(file.fileId)}/state`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ state: file }),
    }).catch(() => {});
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

  const patchActiveFile = (updater: (file: FileViewState) => FileViewState) => {
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

  const updateRowAIResult = (
    fileId: string,
    rowId: string,
    stageKey: AIDetectStageKey,
    resultText: string,
  ) => {
    let nextFileToPersist: FileViewState | null = null;
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
        nextFileToPersist = nextFile;
        return nextFile;
      }),
    );

    if (nextFileToPersist) {
      schedulePersistFileState(nextFileToPersist);
    }
  };

  const persistColumnPrefs = (file: FileViewState) => {
    fetch(`/api/column-prefs/${encodeURIComponent(file.fileName)}`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        fieldSignature: getFieldSignature(file.columns),
        displayKeys: file.selectedDisplayColumnKeys,
        editableKeys: file.selectedEditableColumnKeys,
        filterKeys: file.selectedFilterColumnKeys,
      }),
    }).catch(() => {});
  };

  const resetPendingConfigState = () => {
    setPendingFile(null);
    setPendingSelectedDisplayKeys([]);
    setPendingSelectedFilterKeys([]);
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
    setPendingSelectedFilterKeys(activeFile.selectedFilterColumnKeys);
    setPendingEditableColumnKeys(activeFile.selectedEditableColumnKeys);
    setPendingConfigNotice("");
    setPendingConfigMode("edit");
    navigateToSection("settings", "fields");
  };

  const syncActiveAIConfigState = (nextConfig: AIDetectConfig) => {
    setAIConfig(nextConfig);
    setAIConfigList((previous) =>
      previous.map((item) =>
        item.name === selectedAIConfigName
          ? {
              ...item,
              config: nextConfig,
            }
          : item,
      ),
    );
  };

  const onSelectAIConfigForRun = (configName: string) => {
    if (!activeFile) {
      return;
    }
    const matched = aiConfigList.find((item) => item.name === configName);
    if (!matched) {
      return;
    }

    const normalized = normalizeAIDetectConfigForColumns(
      matched.config,
      activeFile.columns,
    );
    setSelectedAIConfigName(configName);
    setAIConfig(normalized);
    setAIConfigList((previous) =>
      previous.map((item) =>
        item.name === configName
          ? {
              ...item,
              config: normalized,
            }
          : item,
      ),
    );
    if (!isAIStageConfigModalOpen && !isAIProfileModalOpen) {
      setDraftAIConfigName(configName);
      setDraftAIConfig(cloneAIDetectConfig(normalized));
    }
    setAIConfigFormMessage("");

    fetch(`/api/ai-config/${encodeURIComponent(activeFile.fileName)}/active`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        name: configName,
      }),
    }).catch(() => {});
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
    setDraftAIConfigName(selectedAIConfigName);
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
    setDraftAIConfigName(selectedAIConfigName);
    setAIConfigFormMessage("");
    setIsAIStageConfigModalOpen(false);
  };

  const onCancelAIProfileModal = () => {
    setDraftAIConfig(cloneAIDetectConfig(aiConfig));
    setDraftAIConfigName(selectedAIConfigName);
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

    const nextConfigName = normalizeAIConfigName(draftAIConfigName);
    if (draftAIConfigName.trim().length === 0) {
      setAIConfigFormMessage("配置名称不能为空");
      return;
    }

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
      if (profile.provider === "openai") {
        if (profile.url.trim().length === 0) {
          setAIConfigFormMessage(
            `【${profileName}】OpenAI 兼容接口 URL 不能为空`,
          );
          return;
        }
        if (profile.apiKey.trim().length === 0) {
          setAIConfigFormMessage(`【${profileName}】OpenAI API Key 不能为空`);
          return;
        }
      }
      if (profile.provider === "gemini") {
        if (profile.url.trim().length === 0) {
          setAIConfigFormMessage(
            `【${profileName}】Gemini 接口 Endpoint 不能为空`,
          );
          return;
        }
        if (profile.apiKey.trim().length === 0) {
          setAIConfigFormMessage(`【${profileName}】Gemini API Key 不能为空`);
          return;
        }
      }
    }

    // Skip stage validation when saving from profile modal
    if (!skipStageValidation) {
      for (const stageKey of AI_STAGE_ORDER) {
        const stageConfig = nextConfig.stages[stageKey];
        const stageLabel = AI_STAGE_LABELS[stageKey]?.shortTitle ?? stageKey;

        if (!profileNameSet.has(stageConfig.profileName)) {
          setAIConfigFormMessage(`【${stageLabel}】请选择有效的接口配置`);
          return;
        }
        if (stageConfig.submitFieldKeys.length === 0) {
          setAIConfigFormMessage(`【${stageLabel}】请至少选择一个提交回答字段`);
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
      setSelectedAIConfigName(nextConfigName);
      setAIConfig(nextConfig);
      setDraftAIConfigName(nextConfigName);
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

  // Save profile config only (skip stage validation)
  const onSaveAIProfileConfig = () => onSaveAIConfig(true);

  // Save stage config (full validation)
  const onSaveAIStageConfig = () => onSaveAIConfig(false);

  const onRunAIDetect = async () => {
    if (!activeFile || !selectedRow) {
      return;
    }
    if (isAIBatchRunning) {
      setAIResultMessage("批量 AI 任务运行中，暂不可发起单条回答");
      return;
    }

    const normalizedConfig = normalizeAIDetectConfigForColumns(
      aiConfig,
      activeFile.columns,
    );
    syncActiveAIConfigState(normalizedConfig);
    const runningConfigName = selectedAIConfigName;
    const stageConfig = normalizedConfig.stages[activeAIStageKey];
    const stageLabel = AI_STAGE_LABELS[activeAIStageKey]?.shortTitle ?? "";
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
    if (profile.provider === "openai") {
      if (profile.url.trim().length === 0) {
        setAIResultMessage("请先配置 OpenAI 兼容接口 URL");
        return;
      }
      if (profile.apiKey.trim().length === 0) {
        setAIResultMessage("请先配置 OpenAI API Key");
        return;
      }
    }
    if (profile.provider === "gemini") {
      if (profile.url.trim().length === 0) {
        setAIResultMessage("请先配置 Gemini 接口 Endpoint");
        return;
      }
      if (profile.apiKey.trim().length === 0) {
        setAIResultMessage("请先配置 Gemini API Key");
        return;
      }
    }
    if (stageConfig.submitFieldKeys.length === 0) {
      setAIResultMessage("请先在 AI 配置中选择提交回答字段");
      return;
    }
    if (stageConfig.prompt.trim().length === 0) {
      setAIResultMessage("请先配置 Prompt");
      return;
    }

    const fields = buildAIDetectFieldsForRow(
      activeFile.columns,
      selectedRow,
      stageConfig.submitFieldKeys,
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
      setAIThinkingText(streamResult.thinkingText);
      const composedText = composeAISaveText(
        streamResult.answerText,
        streamResult.thinkingText,
      );
      if (composedText.trim().length === 0) {
        setAIResultMessage("AI 返回为空");
      } else {
        updateRowAIResult(
          activeFile.fileId,
          selectedRow.rowId,
          activeAIStageKey,
          composedText,
        );
        setAIResultMessage(
          `AI 回答完成（配置：${runningConfigName}${stageLabel ? ` / ${stageLabel}` : ""}），已写入 AI 检测结果`,
        );
      }
    } catch (error) {
      if (controller.signal.aborted) {
        setAIResultMessage("AI 回答已取消");
      } else {
        const message = error instanceof Error ? error.message : "AI 回答失败";
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

  const applyBatchAIResultsToFile = (
    fileId: string,
    stageKey: AIDetectStageKey,
    resultMap: Map<string, string>,
  ) => {
    if (resultMap.size === 0) {
      return;
    }

    let nextFileToPersist: FileViewState | null = null;
    setFiles((previous) =>
      previous.map((file) => {
        if (file.fileId !== fileId) {
          return file;
        }

        const nextRows = file.rows.map((row) => {
          const result = resultMap.get(row.rowId);
          if (result === undefined) {
            return row;
          }
          return {
            ...row,
            aiResults: {
              ...(row.aiResults ?? {}),
              [stageKey]: result,
            },
          };
        });

        const nextFile: FileViewState = {
          ...file,
          rows: nextRows,
        };
        nextFileToPersist = nextFile;
        return nextFile;
      }),
    );

    if (nextFileToPersist) {
      schedulePersistFileState(nextFileToPersist);
    }
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
    const runningConfigName = selectedAIConfigName;
    const stageConfig = normalizedConfig.stages[activeAIStageKey];
    const stageLabel = AI_STAGE_LABELS[activeAIStageKey]?.shortTitle ?? "";
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
    if (profile.provider === "openai") {
      if (profile.url.trim().length === 0) {
        setErrorMessage("请先配置 OpenAI 兼容接口 URL");
        return;
      }
      if (profile.apiKey.trim().length === 0) {
        setErrorMessage("请先配置 OpenAI API Key");
        return;
      }
    }
    if (profile.provider === "gemini") {
      if (profile.url.trim().length === 0) {
        setErrorMessage("请先配置 Gemini 接口 Endpoint");
        return;
      }
      if (profile.apiKey.trim().length === 0) {
        setErrorMessage("请先配置 Gemini API Key");
        return;
      }
    }
    if (stageConfig.submitFieldKeys.length === 0) {
      setErrorMessage("请先在 AI 配置中选择提交回答字段");
      return;
    }
    if (stageConfig.prompt.trim().length === 0) {
      setErrorMessage("请先配置 Prompt");
      return;
    }

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
    const resultMap = new Map<string, string>();
    let nextCursor = 0;
    const requestedConcurrency =
      normalizeAIBatchConcurrency(aiBatchConcurrency);
    const workerCount = Math.min(requestedConcurrency, targetRows.length);

    aiBatchAbortRef.current?.abort();
    const controller = new AbortController();
    aiBatchAbortRef.current = controller;
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
          ? `已选择 ${targetRows.length} 条，并发 ${workerCount} 线程`
          : `并发 ${workerCount} 线程`,
    });
    setErrorMessage("");
    setAIResultMessage("");

    const runWorker = async () => {
      while (!controller.signal.aborted) {
        const currentIndex = nextCursor;
        nextCursor += 1;
        if (currentIndex >= targetRows.length) {
          return;
        }

        const row = targetRows[currentIndex];
        try {
          const fields = buildAIDetectFieldsForRow(
            targetColumns,
            row,
            stageConfig.submitFieldKeys,
          );
          if (fields.length === 0) {
            throw new Error("没有可提交的回答字段");
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
            { signal: controller.signal },
          );
          const text = composeAISaveText(
            streamResult.answerText,
            streamResult.thinkingText,
          );

          if (text.trim().length === 0) {
            throw new Error("AI 返回为空");
          }
          resultMap.set(row.rowId, text);
          setAIBatchTask((previous) => ({
            ...previous,
            completed: previous.completed + 1,
            success: previous.success + 1,
          }));
        } catch {
          if (controller.signal.aborted) {
            return;
          }
          setAIBatchTask((previous) => ({
            ...previous,
            completed: previous.completed + 1,
            failed: previous.failed + 1,
          }));
        }
      }
    };

    try {
      await Promise.all(Array.from({ length: workerCount }, () => runWorker()));

      if (controller.signal.aborted) {
        return;
      }

      applyBatchAIResultsToFile(targetFileId, activeAIStageKey, resultMap);

      setAIBatchTask((previous) => ({
        ...previous,
        status: "completed",
        message: `结果已写入 AI 检测结果（配置：${runningConfigName}${stageLabel ? ` / ${stageLabel}` : ""}）`,
      }));
      setErrorMessage("");
    } catch (error) {
      if (controller.signal.aborted) {
        return;
      }

      const message =
        error instanceof Error ? error.message : "批量 AI 回答任务执行失败";
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
    }
  };

  const onRunSelectedBatchAIAnswer = async () => {
    await onRunBatchAIAnswer(batchSelectedRowIds);
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
        exportColumns.map((column) => row.values[column.key]?.value ?? ""),
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

  const onUploadClick = () => {
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
      const formData = new FormData();
      formData.append("file", selected);
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

      const parsed = (await response.json()) as ParsedFile;
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
      let initialFilterKeys = normalizeFilterSelection(parsed.columns);
      let shouldShowColumnModal = true;
      let nextPendingNotice = "";

      try {
        const prefsRes = await fetch(
          `/api/column-prefs/${encodeURIComponent(parsed.fileName)}`,
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
            const normalizedFilterKeys = normalizeFilterSelection(
              parsed.columns,
              prefsData.config.filterKeys,
            );
            const currentSignature = getFieldSignature(parsed.columns);
            if (prefsData.config.fieldSignature === currentSignature) {
              const nextFile = toViewState(
                parsed,
                normalizedSaved.displayKeys,
                normalizedSaved.editableKeys,
                normalizedFilterKeys,
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
              initialFilterKeys = normalizedFilterKeys;
            }
          }
        }
      } catch {
        // Ignore and fall back to default selection
      }

      if (shouldShowColumnModal) {
        setPendingFile(parsed);
        setPendingSelectedDisplayKeys(initialDisplayKeys);
        setPendingSelectedFilterKeys(initialFilterKeys);
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

  const onTogglePendingFilterColumn = (columnKey: string) => {
    if (!pendingFile) {
      return;
    }
    setPendingSelectedFilterKeys((previous) => {
      const exists = previous.includes(columnKey);
      return exists
        ? previous.filter((key) => key !== columnKey)
        : [...previous, columnKey];
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

  const onPendingClearFilterColumns = () => {
    setPendingSelectedFilterKeys([]);
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
          pendingSelectedFilterKeys,
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
      pendingSelectedFilterKeys,
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
        ? file.selectedDisplayColumnKeys.filter((key) => key !== columnKey)
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
        selectedFilterColumnKeys: file.selectedFilterColumnKeys,
        columnFilterValues: file.columnFilterValues,
      };
      persistColumnPrefs(nextFile);
      return nextFile;
    });
  };

  const onColumnFilterChange = (columnKey: string, value: string) => {
    patchActiveFile((file) => ({
      ...file,
      columnFilterValues: {
        ...file.columnFilterValues,
        [columnKey]: value,
      },
    }));
  };

  const onToggleBatchRowSelection = (rowId: string) => {
    setBatchSelectedRowIds((previous) => {
      if (previous.includes(rowId)) {
        return previous.filter((item) => item !== rowId);
      }
      return [...previous, rowId];
    });
  };

  const onSelectAllBatchRows = () => {
    setBatchSelectedRowIds(visibleRows.map((row) => row.rowId));
  };

  const onClearBatchRows = () => {
    setBatchSelectedRowIds([]);
  };

  const getLatexToggleKey = (columnKey: string) =>
    activeFileId ? `${activeFileId}::${columnKey}` : columnKey;

  const onToggleLatexRender = (columnKey: string) => {
    const key = getLatexToggleKey(columnKey);
    setLatexRenderOverrides((previous) => ({
      ...previous,
      [key]: !(previous[key] ?? false),
    }));
  };

  const renderReadonlyCell = (
    row: ParsedRow,
    column: ParsedColumn,
    cell: ParsedCell | undefined,
    shouldRenderLatex: boolean,
  ) => {
    if (!cell) {
      return <span className="empty-text">-</span>;
    }

    if (cell.type === "image" && cell.src) {
      return (
        <div className="image-cell">
          <img
            src={cell.src}
            alt={cell.value || "Excel图片"}
            onClick={() => setPreviewImageSrc(cell.src!)}
            onError={() => {
              logUIImageRenderError(row.rowId, column.title, cell.src ?? "");
            }}
          />
          {cell.value ? <span>{cell.value}</span> : null}
        </div>
      );
    }

    const textValue = cell.value ?? "";
    if (cell.type === "text" && textValue.length > 0) {
      const hasLatex = hasLatexSyntax(textValue);
      const autoDisplayLatex = shouldAutoDisplayLatex(textValue);
      if (hasLatex && shouldRenderLatex) {
        return (
          <LatexRenderer value={textValue} forceDisplay={autoDisplayLatex} />
        );
      }
      return hasLatex ? (
        <div className="latex-plain">{textValue}</div>
      ) : (
        <div className="plain-text-value">{textValue}</div>
      );
    }

    return cell.value ? (
      <div className="plain-text-value">{cell.value}</div>
    ) : (
      <span className="empty-text">-</span>
    );
  };

  const renderCellContent = (
    row: ParsedRow,
    column: ParsedColumn,
    shouldRenderLatex = true,
  ) => {
    const cell = row.values[column.key];
    if (!column.editable) {
      return renderReadonlyCell(row, column, cell, shouldRenderLatex);
    }

    const currentValue = cell?.value ?? "";

    if (isQualifiedColumnTitle(column.title)) {
      const stableOptions = ["", "合格", "不合格"];
      const shouldAppendCurrent =
        currentValue.length > 0 && !stableOptions.includes(currentValue);
      return (
        <select
          className="qualified-select"
          value={currentValue}
          onChange={(event) =>
            onEditCell(row.rowId, column.key, event.target.value)
          }
        >
          <option value="">未填写</option>
          <option value="合格">合格</option>
          <option value="不合格">不合格</option>
          {shouldAppendCurrent ? (
            <option value={currentValue}>{currentValue}</option>
          ) : null}
        </select>
      );
    }

    if (isOpensourceColumnTitle(column.title)) {
      const stableOptions = ["", "是", "否"];
      const shouldAppendCurrent =
        currentValue.length > 0 && !stableOptions.includes(currentValue);
      return (
        <select
          className="qualified-select"
          value={currentValue}
          onChange={(event) =>
            onEditCell(row.rowId, column.key, event.target.value)
          }
        >
          <option value="">未填写</option>
          <option value="是">是</option>
          <option value="否">否</option>
          {shouldAppendCurrent ? (
            <option value={currentValue}>{currentValue}</option>
          ) : null}
        </select>
      );
    }

    if (isInspectorColumnTitle(column.title)) {
      return (
        <input
          className="inspector-input"
          value={currentValue}
          onChange={(event) =>
            onEditCell(row.rowId, column.key, event.target.value)
          }
          placeholder="请输入质检员"
        />
      );
    }

    if (isFeedbackColumnTitle(column.title)) {
      return (
        <textarea
          className="feedback-input"
          value={currentValue}
          onChange={(event) =>
            onEditCell(row.rowId, column.key, event.target.value)
          }
          placeholder="请输入质检反馈意见"
        />
      );
    }

    if (cell?.type === "image" && cell.src) {
      return (
        <div className="image-cell">
          <img
            src={cell.src}
            alt={cell.value || "Excel图片"}
            onClick={() => setPreviewImageSrc(cell.src!)}
            onError={() => {
              logUIImageRenderError(row.rowId, column.title, cell.src ?? "");
            }}
          />
          <input
            className="editable-text-input"
            value={currentValue}
            onChange={(event) =>
              onEditCell(row.rowId, column.key, event.target.value)
            }
            placeholder={`请输入${column.title}`}
          />
        </div>
      );
    }

    return (
      <input
        className="editable-text-input"
        value={currentValue}
        onChange={(event) =>
          onEditCell(row.rowId, column.key, event.target.value)
        }
        placeholder={`请输入${column.title}`}
      />
    );
  };

  const renderDetailField = (column: ParsedColumn, isHidden = false) => {
    if (!selectedRow) return null;
    const isRequired = column.editable;
    const isChecked = !isHidden;
    const cell = selectedRow.values[column.key];
    const hasLatex =
      !column.editable &&
      !isHidden &&
      cell?.type === "text" &&
      typeof cell.value === "string" &&
      hasLatexSyntax(cell.value);
    const latexToggleKey = getLatexToggleKey(column.key);
    const isLatexRenderingEnabled =
      latexRenderOverrides[latexToggleKey] ?? false;

    return (
      <div
        key={`${selectedRow.rowId}_${column.key}`}
        className={`detail-field ${isHidden ? "hidden-field" : ""}`}
      >
        <div className="detail-label">
          <button
            type="button"
            className={`field-toggle ${isRequired ? "locked" : ""} ${isChecked ? "checked" : ""}`}
            onClick={() => {
              if (!isRequired) {
                onToggleDisplayColumn(column.key);
              }
            }}
            title={
              isRequired
                ? "可编辑字段必须展示"
                : isHidden
                  ? "点击显示此字段"
                  : "点击隐藏此字段"
            }
          />
          <div className="field-name-wrap">
            <span className="field-name">{column.title}</span>
            {hasLatex ? (
              <label
                className="latex-toggle"
                title="控制该字段是否按 LaTeX 公式渲染"
              >
                <input
                  type="checkbox"
                  checked={isLatexRenderingEnabled}
                  onChange={() => onToggleLatexRender(column.key)}
                  aria-label={`${column.title} 的 LaTeX 渲染开关`}
                />
                <span>LaTeX渲染</span>
              </label>
            ) : null}
          </div>
          {column.editable ? (
            <span className="field-badge badge-editable">可编辑</span>
          ) : null}
          {isRequired ? (
            <span className="field-badge badge-locked">必显</span>
          ) : null}
        </div>
        {!isHidden ? (
          <div className="detail-value">
            {renderCellContent(selectedRow, column, isLatexRenderingEnabled)}
          </div>
        ) : null}
      </div>
    );
  };

  const renderListReadonlyCell = (row: ParsedRow, column: ParsedColumn) => {
    const cell = row.values[column.key];
    // Hide filename text for image columns on list page (show "-" instead)
    const isImageColumn = /图片/.test(column.title);
    if (isImageColumn && cell?.type === "text") {
      return <span className="empty-text">-</span>;
    }
    return renderReadonlyCell(
      row,
      {
        ...column,
        editable: false,
      },
      cell,
      false,
    );
  };
  const getListCellTitle = (row: ParsedRow, column: ParsedColumn) =>
    getCellText(row, column.key).trim();

  const listRangeStart =
    visibleRows.length === 0 ? 0 : (listPage - 1) * listPageSize + 1;
  const listRangeEnd = Math.min(listPage * listPageSize, visibleRows.length);
  const routePathLabel = buildHashRoute(activeSection, activeSettingsSection);
  const pageTitle =
    activeSection === "list"
      ? "题目列表"
      : activeSection === "detail"
        ? "题目详情"
        : activeSettingsSection === "ai"
          ? "AI 设置"
          : "字段设置";
  const pageDescription =
    activeSection === "list"
      ? `当前文件共 ${visibleRows.length} 条，正在展示 ${listRangeStart}-${listRangeEnd} 条。`
      : activeSection === "detail"
        ? selectedRow
          ? `当前查看第 ${activeRowIndex + 1} 条，支持字段编辑与 AI 回答。`
          : "请先在题目列表中选择一条记录。"
        : activeSettingsSection === "ai"
          ? "管理接口配置与阶段任务配置，控制提示词与结果保存字段。"
          : "管理详情页字段展示和可编辑字段。";

  return (
    <div className="app-shell">
      <HeaderBar
        files={files}
        activeFileId={activeFileId}
        onSelectFile={setActiveFileId}
        onRemoveFile={onRemoveFile}
        errorMessage={errorMessage}
        aiBatchTask={aiBatchTask}
        aiBatchProgressPercent={aiBatchProgressPercent}
        isAIBatchRunning={isAIBatchRunning}
        theme={theme}
        onToggleTheme={toggleTheme}
        onOpenAIStageConfigModal={onOpenAIStageConfigModal}
        onExportFile={onExportFile}
        onUploadClick={onUploadClick}
        uploadInputRef={uploadInputRef}
        onUploadFile={onUploadFile}
        isExporting={isExporting}
        isUploading={isUploading}
        aiConfigLoading={aiConfigLoading}
        activeFile={activeFile}
      />

      {/* ─── Main Content ─── */}
      <main
        className={`main-content main-workspace ${isSidebarCollapsed ? "sidebar-collapsed" : ""}`}
      >
        <WorkspaceSidebar
          isCollapsed={isSidebarCollapsed}
          activeSection={activeSection}
          activeSettingsSection={activeSettingsSection}
          activeFile={activeFile}
          onToggle={() => setIsSidebarCollapsed((previous) => !previous)}
          onNavigate={navigateToSection}
        />

        <section className="workspace-main">
          {!activeFile ? (
            <section className="placeholder workspace-placeholder">
              <div className="placeholder-icon">
                <IconFile />
              </div>
              <h2>等待文件导入</h2>
              <p>
                点击右上角「导入
                Excel」按钮，导入后可在左侧切换列表、详情与设置。
              </p>
            </section>
          ) : (
            <>
              <section className="workspace-topbar">
                <div className="workspace-topbar-head">
                  <div className="workspace-topbar-copy">
                    <span className="workspace-route">{routePathLabel}</span>
                    <h2>{pageTitle}</h2>
                    <p>{pageDescription}</p>
                  </div>
                  <div className="workspace-topbar-meta">
                    <span>{activeFile.fileName}</span>
                    {activeSection === "list" ? (
                      <>
                        <span>字段 {activeFile.columns.length}</span>
                        <span>已勾选 {batchSelectedRowIds.length}</span>
                      </>
                    ) : null}
                    {activeSection === "detail" ? (
                      <>
                        <span>展示字段 {displayColumns.length}</span>
                        <span>隐藏字段 {hiddenColumns.length}</span>
                      </>
                    ) : null}
                    {activeSection === "settings" ? (
                      <>
                        <span>
                          当前分区{" "}
                          {activeSettingsSection === "ai" ? "AI" : "字段"}
                        </span>
                        <span>当前配置 {selectedAIConfigName}</span>
                      </>
                    ) : null}
                  </div>
                </div>

                <div className="toolbar page-toolbar">
                  {activeSection !== "settings"
                    ? filterColumns.map((column) => {
                        const options = filterOptionsMap.get(column.key) ?? [];
                        return (
                          <div className="filter-group" key={column.key}>
                            <label htmlFor={`filter-${column.key}`}>
                              {column.title}
                            </label>
                            <select
                              id={`filter-${column.key}`}
                              value={
                                activeFile.columnFilterValues[column.key] ??
                                ALL_FILTER_VALUE
                              }
                              onChange={(event) =>
                                onColumnFilterChange(
                                  column.key,
                                  event.target.value,
                                )
                              }
                            >
                              <option value={ALL_FILTER_VALUE}>
                                {ALL_FILTER_VALUE}
                              </option>
                              {options.map((item) => (
                                <option key={item} value={item}>
                                  {item}
                                </option>
                              ))}
                            </select>
                          </div>
                        );
                      })
                    : null}
                  <div className="toolbar-spacer" />
                  {activeSection === "list" ? (
                    <div className="toolbar-actions">
                      <div className="batch-toolbar">
                        <label className="batch-control">
                          <span>批量并发</span>
                          <input
                            type="number"
                            min={MIN_AI_BATCH_CONCURRENCY}
                            max={MAX_AI_BATCH_CONCURRENCY}
                            step={1}
                            value={aiBatchConcurrency}
                            onChange={(event) =>
                              setAIBatchConcurrency(
                                normalizeAIBatchConcurrency(
                                  Number(event.target.value),
                                ),
                              )
                            }
                            disabled={isAIBatchRunning}
                          />
                        </label>
                        <div className="batch-buttons">
                          <button
                            type="button"
                            className="btn"
                            onClick={onSelectAllBatchRows}
                            disabled={
                              visibleRows.length === 0 || isAIBatchRunning
                            }
                          >
                            {batchSelectedRowIds.length ===
                              visibleRows.length && visibleRows.length > 0
                              ? "取消全选"
                              : "全选可见"}
                          </button>
                          <button
                            type="button"
                            className="btn"
                            onClick={onClearBatchRows}
                            disabled={
                              batchSelectedRowIds.length === 0 ||
                              isAIBatchRunning
                            }
                          >
                            清空勾选
                          </button>
                          <button
                            type="button"
                            className="btn btn-primary"
                            onClick={onRunSelectedBatchAIAnswer}
                            disabled={
                              aiConfigLoading ||
                              isAIDetecting ||
                              isAIBatchRunning ||
                              batchSelectedRowIds.length === 0
                            }
                          >
                            {isAIBatchRunning
                              ? "AI批量回答中..."
                              : `批量回答已选 ${batchSelectedRowIds.length} 条`}
                          </button>
                        </div>
                      </div>
                    </div>
                  ) : null}
                  {activeSection === "detail" ? (
                    <div className="toolbar-actions">
                      <button
                        type="button"
                        className="btn"
                        onClick={() => navigateToSection("list")}
                      >
                        返回列表
                      </button>
                      <button
                        type="button"
                        className="btn"
                        onClick={() =>
                          previousRow && openRowDetail(previousRow.rowId)
                        }
                        disabled={!previousRow}
                      >
                        上一题
                      </button>
                      <button
                        type="button"
                        className="btn"
                        onClick={() => nextRow && openRowDetail(nextRow.rowId)}
                        disabled={!nextRow}
                      >
                        下一题
                      </button>
                    </div>
                  ) : null}
                  {activeSection === "settings" ? (
                    <div className="settings-tabs">
                      <button
                        type="button"
                        className={`btn ${activeSettingsSection === "fields" ? "btn-primary" : ""}`}
                        onClick={() => navigateToSection("settings", "fields")}
                      >
                        字段设置
                      </button>
                      <button
                        type="button"
                        className={`btn ${activeSettingsSection === "ai" ? "btn-primary" : ""}`}
                        onClick={() => navigateToSection("settings", "ai")}
                      >
                        AI 设置
                      </button>
                    </div>
                  ) : null}
                </div>
              </section>

              <section className="workspace-view">
                {activeSection === "list" ? (
                  <section className="page-panel">
                    <ListPage
                      activeFile={activeFile}
                      visibleRows={visibleRows}
                      paginatedRows={paginatedRows}
                      listPage={listPage}
                      listPageSize={listPageSize}
                      totalListPages={totalListPages}
                      listPageSizeOptions={LIST_PAGE_SIZE_OPTIONS}
                      batchSelectedRowIdSet={batchSelectedRowIdSet}
                      selectedRowId={selectedRowId}
                      onToggleBatchRowSelection={onToggleBatchRowSelection}
                      onOpenRowDetail={openRowDetail}
                      onPageChange={setListPage}
                      onPageSizeChange={setListPageSize}
                      getCellTitle={getListCellTitle}
                      renderListReadonlyCell={renderListReadonlyCell}
                    />
                  </section>
                ) : null}

                {activeSection === "detail" ? (
                  <section className="page-panel detail-page-panel">
                    <DetailPage
                      selectedRow={selectedRow}
                      displayColumns={displayColumns}
                      hiddenColumns={hiddenColumns}
                      showHiddenFields={showHiddenFields}
                      onToggleHiddenFields={() =>
                        setShowHiddenFields((previous) => !previous)
                      }
                      onOpenAIRunModal={() => setIsAIRunModalOpen(true)}
                      renderDetailField={renderDetailField}
                      aiResults={selectedRow?.aiResults}
                    />
                  </section>
                ) : null}

                {activeSection === "settings" ? (
                  <section className="page-panel settings-page-panel">
                    <SettingsPage
                      activeSettingsSection={activeSettingsSection}
                      activeFile={activeFile}
                      displayColumns={displayColumns}
                      aiConfigList={aiConfigList}
                      selectedAIConfigName={selectedAIConfigName}
                      aiConfig={aiConfig}
                      onOpenActiveFileConfig={onOpenActiveFileConfig}
                      onOpenAIStageConfigModal={onOpenAIStageConfigModal}
                      onOpenAIProfileModal={onOpenAIProfileModal}
                    />
                  </section>
                ) : null}
              </section>
            </>
          )}
        </section>
      </main>
      {/* ─── Column Selection Modal ─── */}
      <ColumnConfigModal
        pendingFile={pendingFile}
        pendingConfigMode={pendingConfigMode}
        pendingConfigNotice={pendingConfigNotice}
        pendingSelectedDisplayKeys={pendingSelectedDisplayKeys}
        pendingSelectedFilterKeys={pendingSelectedFilterKeys}
        pendingEditableColumnKeys={pendingEditableColumnKeys}
        onPendingSelectAllDisplayColumns={onPendingSelectAllDisplayColumns}
        onPendingClearDisplayColumns={onPendingClearDisplayColumns}
        onPendingClearFilterColumns={onPendingClearFilterColumns}
        onPendingClearEditableColumns={onPendingClearEditableColumns}
        onTogglePendingDisplayColumn={onTogglePendingDisplayColumn}
        onTogglePendingFilterColumn={onTogglePendingFilterColumn}
        onTogglePendingEditableColumn={onTogglePendingEditableColumn}
        onCancelPendingFile={onCancelPendingFile}
        onConfirmPendingFile={onConfirmPendingFile}
      />

      {/* ─── AI Stage Config Modal ─── */}
      <AIStageConfigModal
        isOpen={isAIStageConfigModalOpen}
        activeFile={activeFile}
        aiConfigFormMessage={aiConfigFormMessage}
        aiConfigList={aiConfigList}
        draftAIConfigName={draftAIConfigName}
        setDraftAIConfigName={setDraftAIConfigName}
        draftAIConfig={draftAIConfig}
        setDraftAIConfig={setDraftAIConfig}
        aiSubmitFieldColumns={aiSubmitFieldColumns}
        aiConfigSaving={aiConfigSaving}
        onToggleDraftAISubmitField={onToggleDraftAISubmitField}
        onCancel={onCancelAIStageConfigModal}
        onSave={onSaveAIStageConfig}
      />
      {/* ─── AI Profile Modal ─── */}
      <AIProfileModal
        isOpen={isAIProfileModalOpen}
        activeFile={activeFile}
        aiConfigFormMessage={aiConfigFormMessage}
        aiConfigList={aiConfigList}
        draftAIConfigName={draftAIConfigName}
        setDraftAIConfigName={setDraftAIConfigName}
        draftAIConfig={draftAIConfig}
        setDraftAIConfig={setDraftAIConfig}
        aiConfigSaving={aiConfigSaving}
        onCancel={onCancelAIProfileModal}
        onSave={onSaveAIProfileConfig}
      />
      <AIRunModal
        isOpen={isAIRunModalOpen}
        rowId={selectedRow?.rowId}
        aiConfigList={aiConfigList}
        selectedAIConfigName={selectedAIConfigName}
        onSelectAIConfigForRun={onSelectAIConfigForRun}
        aiStageKey={activeAIStageKey}
        onSelectAIStage={setActiveAIStageKey}
        aiConfigLoading={aiConfigLoading}
        isAIDetecting={isAIDetecting}
        isAIBatchRunning={isAIBatchRunning}
        aiDetectElapsedText={aiDetectElapsedText}
        canRunAIDetect={
          Boolean(selectedRow) &&
          !isAIDetecting &&
          !aiConfigLoading &&
          !isAIBatchRunning
        }
        onRunAIDetect={onRunAIDetect}
        aiRetryCount={activeProfile?.profile.retryCount ?? 0}
        aiMergedStreamText={aiMergedStreamText}
        onAIResultTextChange={(value) => {
          setAIThinkingText("");
          setAIResultText(value);
        }}
        aiResultMessage={aiResultMessage}
        aiRequestPreview={aiRequestPreview}
        onClose={() => setIsAIRunModalOpen(false)}
      />

      {/* ─── Image Lightbox ─── */}
      <ImageLightbox
        src={previewImageSrc}
        onClose={() => setPreviewImageSrc(null)}
      />
    </div>
  );
}

export default App;
