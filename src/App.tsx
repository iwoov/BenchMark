import {
    useEffect,
    useMemo,
    useRef,
    useState,
    type CSSProperties,
    type MouseEvent as ReactMouseEvent,
} from "react";
import {
    AI_CLEANING_TOOL_LABELS,
    AI_CLEANING_TOOL_ORDER,
    AI_RUN_ALL_KEY,
    AI_RUN_ALL_LABEL,
    AI_RUN_STAGE_ORDER,
    AI_STAGE_LABELS,
    LIST_PAGE_SIZE_OPTIONS,
    MAX_AI_BATCH_CONCURRENCY,
    MIN_AI_BATCH_CONCURRENCY,
} from "./app/constants";
import {
    normalizeAIBatchConcurrency,
    parseAIResultJSON,
} from "./app/ai-helpers";
import { getLevelColumnKey } from "./app/file-helpers";
import { HeaderBar } from "./app/components/HeaderBar";
import { WorkspaceSidebar } from "./app/components/WorkspaceSidebar";
import { DashboardPage } from "./app/components/DashboardPage";
import { ListPage } from "./app/components/ListPage";
import { DetailPage } from "./app/components/DetailPage";
import { SettingsPage } from "./app/components/SettingsPage";
import { ColumnConfigModal } from "./app/components/ColumnConfigModal";
import { AIProfileModal } from "./app/components/AIProfileModal";
import { AIRouteModal } from "./app/components/AIRouteModal";
import { AIStageConfigModal } from "./app/components/AIStageConfigModal";
import { AIRunModal } from "./app/components/AIRunModal";
import { AIChatConfigModal } from "./app/components/AIChatConfigModal";
import { AICleaningConfigModal } from "./app/components/AICleaningConfigModal";
import { AIChatSidebar } from "./app/components/AIChatSidebar";
import { ImageLightbox } from "./app/components/ImageLightbox";
import { FilterConfigModal } from "./app/components/FilterConfigModal";
import { IconFile, IconMessageSquare } from "./app/icons";
import { useTheme } from "./app/hooks/useTheme";
import { getInitialRoute, useRouteState } from "./app/hooks/useRouteState";
import { useFileStore } from "./app/hooks/useFileStore";
import { useListView } from "./app/hooks/useListView";
import { useAIManager } from "./app/hooks/useAIManager";
import { useCellRenderers } from "./app/hooks/useCellRenderers";

function App() {
    const initialRoute = getInitialRoute();
    const [selectedRowId, setSelectedRowId] = useState<string | null>(
        initialRoute.rowId ?? null,
    );
    const [errorMessage, setErrorMessage] = useState<string>("");
    const [showHiddenFields, setShowHiddenFields] = useState(false);
    const [isSidebarCollapsed, setIsSidebarCollapsed] = useState(false);
    const [isDetailChatSidebarHidden, setIsDetailChatSidebarHidden] =
        useState(true);
    const [detailChatSidebarWidth, setDetailChatSidebarWidth] = useState(360);
    const [isFilterModalOpen, setIsFilterModalOpen] = useState(false);
    const [latexRenderOverrides, setLatexRenderOverrides] = useState<
        Record<string, boolean>
    >({});
    const [previewImageSrc, setPreviewImageSrc] = useState<string | null>(null);
    const detailChatResizeRef = useRef<{
        active: boolean;
        startX: number;
        startWidth: number;
    }>({
        active: false,
        startX: 0,
        startWidth: 360,
    });

    const { theme, toggleTheme } = useTheme();
    const { activeSection, activeSettingsSection, navigateToSection } =
        useRouteState({
            initialRoute,
            onRowIdChange: setSelectedRowId,
        });

    const {
        files,
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
    } = useFileStore({ navigateToSection, setErrorMessage });

    const {
        listPage,
        listPageSize,
        setListPage,
        setListPageSize,
        totalListPages,
        paginatedRows,
        visibleRows,
        displayColumns,
        hiddenColumns,
        selectedRow,
        previousRow,
        nextRow,
        batchSelectedRowIds,
        batchSelectedRowIdSet,
        onToggleBatchRowSelection,
        onSelectAllBatchRows,
        onSelectCurrentPageBatchRows,
        onClearBatchRows,
    } = useListView({
        activeFile,
        selectedRowId,
        setSelectedRowId,
        defaultPageSize: LIST_PAGE_SIZE_OPTIONS[2],
    });

    const {
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
        isAICleaningConfigModalOpen,
        isAIRunModalOpen,
        setIsAIRunModalOpen,
        aiBatchTask,
        aiBatchProgressPercent,
        isAIBatchRunning,
        aiBatchConcurrency,
        setAIBatchConcurrency,
        activeAIRunKey,
        setActiveAIRunKey,
        modalStageKey,
        modalFieldLabels,
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
        isAICleaning,
        activeAICleaningToolKey,
        aiCleaningElapsedText,
        aiCleaningStreamText,
        aiCleaningStatusMessage,
        onOpenAIStageConfigModal,
        onOpenAIProfileModal,
        onOpenAIRouteModal,
        onOpenAIChatConfigModal,
        onOpenAICleaningConfigModal,
        onCancelAIStageConfigModal,
        onCancelAIProfileModal,
        onCancelAIRouteModal,
        onCancelAIChatConfigModal,
        onCancelAICleaningConfigModal,
        onToggleDraftAISubmitField,
        onToggleDraftAIChatSubmitField,
        onToggleDraftAICleaningSubmitField,
        onUpdateDraftAICleaningOutputMapping,
        onSaveAIStageConfig,
        onSaveAIProfileConfig,
        onSaveAIRouteConfig,
        onSaveAIChatConfig,
        onSaveAICleaningConfig,
        onRunAICleaning,
        onRunAIDetect,
        onRunAllAIDetect,
        onRunBatchAIAnswer,
        onSendAIChatMessage,
        onClearAIChatSession,
        openAIRunModalForStage,
        onAIResultTextChange,
    } = useAIManager({
        activeFile,
        activeFileId,
        selectedRow,
        selectedRowId,
        setErrorMessage,
        navigateToSection,
        updateRowAIResult,
        updateRowCleaningResult,
        persistFileState,
        flushPendingAIResults,
        latestFileStateRef,
    });

    useEffect(() => {
        if (initialLoadComplete && !activeFile) {
            navigateToSection("dashboard", activeSettingsSection, null, {
                replace: true,
            });
        }
    }, [
        initialLoadComplete,
        activeFile,
        activeSettingsSection,
        navigateToSection,
    ]);

    useEffect(() => {
        const handleMouseMove = (event: MouseEvent) => {
            if (!detailChatResizeRef.current.active) {
                return;
            }
            const deltaX = detailChatResizeRef.current.startX - event.clientX;
            const nextWidth = Math.min(
                560,
                Math.max(280, detailChatResizeRef.current.startWidth + deltaX),
            );
            setDetailChatSidebarWidth(nextWidth);
        };

        const handleMouseUp = () => {
            detailChatResizeRef.current.active = false;
        };

        window.addEventListener("mousemove", handleMouseMove);
        window.addEventListener("mouseup", handleMouseUp);
        return () => {
            window.removeEventListener("mousemove", handleMouseMove);
            window.removeEventListener("mouseup", handleMouseUp);
        };
    }, []);

    const aiSubmitFieldColumns = useMemo(
        () => (activeFile ? activeFile.columns : []),
        [activeFile],
    );

    const openRowDetail = (rowId: string) => {
        setSelectedRowId(rowId);
        navigateToSection("list", activeSettingsSection, rowId);
    };

    const onRunSelectedBatchAIAnswer = async () => {
        await onRunBatchAIAnswer(batchSelectedRowIds);
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

    const level3TagsFieldKey = useMemo(
        () =>
            aiConfig.cleaning.generate_level3_tags.outputMappings.find(
                (item) => item.outputKey === "tags",
            )?.targetFieldKey ?? "",
        [aiConfig.cleaning.generate_level3_tags.outputMappings],
    );
    const onRemoveLevel3Tag = async (tag: string) => {
        if (!selectedRow || !activeFile) {
            return;
        }
        const currentResult =
            selectedRow.cleaningResults?.generate_level3_tags ?? null;
        const parsed =
            parseAIResultJSON(currentResult?.parsedJsonText ?? "") ??
            parseAIResultJSON(currentResult?.responseText ?? "");
        if (!parsed || !Array.isArray(parsed.tags)) {
            return;
        }
        const nextTags = parsed.tags
            .filter((item): item is string => typeof item === "string")
            .map((item) => item.trim())
            .filter((item) => item.length > 0 && item !== tag);
        const nextParsed = {
            ...parsed,
            tags: nextTags,
        };
        const nextParsedJsonText = JSON.stringify(nextParsed);
        const nextResponseText = JSON.stringify(nextParsed, null, 2);
        const response = await fetch(
            `/api/files/${encodeURIComponent(activeFile.fileId)}/cleaning-results/generate_level3_tags`,
            {
                method: "PUT",
                headers: {
                    "Content-Type": "application/json",
                },
                body: JSON.stringify({
                    rowId: selectedRow.rowId,
                    fileName: activeFile.fileName,
                    responseText: nextResponseText,
                    parsedJsonText: nextParsedJsonText,
                }),
            },
        );
        if (!response.ok) {
            const payload = (await response.json().catch(() => ({}))) as {
                message?: string;
            };
            setErrorMessage(payload.message ?? "删除标签失败");
            return;
        }

        const mappedFieldValues =
            level3TagsFieldKey.trim().length > 0
                ? { [level3TagsFieldKey]: nextTags.join(", ") }
                : undefined;
        updateRowCleaningResult(
            activeFile.fileId,
            selectedRow.rowId,
            "generate_level3_tags",
            {
                responseText: nextResponseText,
                parsedJsonText: nextParsedJsonText,
                updatedAt: new Date().toISOString(),
            },
            mappedFieldValues,
        );
    };

    const onAddLevel3Tag = async (tag: string) => {
        if (!selectedRow || !activeFile) {
            return;
        }
        const nextTag = tag.trim();
        if (nextTag.length === 0) {
            return;
        }
        const currentResult =
            selectedRow.cleaningResults?.generate_level3_tags ?? null;
        const parsed =
            parseAIResultJSON(currentResult?.parsedJsonText ?? "") ??
            parseAIResultJSON(currentResult?.responseText ?? "") ??
            {};
        const currentTags = Array.isArray(parsed.tags)
            ? parsed.tags
                  .filter((item): item is string => typeof item === "string")
                  .map((item) => item.trim())
                  .filter((item) => item.length > 0)
            : [];
        if (currentTags.includes(nextTag)) {
            return;
        }
        const nextParsed = {
            ...(parsed && typeof parsed === "object" ? parsed : {}),
            tags: [...currentTags, nextTag],
        };
        const nextParsedJsonText = JSON.stringify(nextParsed);
        const nextResponseText = JSON.stringify(nextParsed, null, 2);
        const response = await fetch(
            `/api/files/${encodeURIComponent(activeFile.fileId)}/cleaning-results/generate_level3_tags`,
            {
                method: "PUT",
                headers: {
                    "Content-Type": "application/json",
                },
                body: JSON.stringify({
                    rowId: selectedRow.rowId,
                    fileName: activeFile.fileName,
                    responseText: nextResponseText,
                    parsedJsonText: nextParsedJsonText,
                }),
            },
        );
        if (!response.ok) {
            const payload = (await response.json().catch(() => ({}))) as {
                message?: string;
            };
            const message = payload.message ?? "添加标签失败";
            setErrorMessage(message);
            throw new Error(message);
        }

        const mappedFieldValues =
            level3TagsFieldKey.trim().length > 0
                ? { [level3TagsFieldKey]: [...currentTags, nextTag].join(", ") }
                : undefined;
        updateRowCleaningResult(
            activeFile.fileId,
            selectedRow.rowId,
            "generate_level3_tags",
            {
                responseText: nextResponseText,
                parsedJsonText: nextParsedJsonText,
                updatedAt: new Date().toISOString(),
            },
            mappedFieldValues,
        );
    };

    const biochemLevel1FieldKey = useMemo(() => {
        const mappedFieldKey =
            aiConfig.cleaning.biochem_level1_refine.outputMappings.find(
                (item) => item.outputKey === "discipline",
            )?.targetFieldKey ?? "";
        if (mappedFieldKey.trim().length > 0) {
            return mappedFieldKey.trim();
        }
        if (!activeFile) {
            return "";
        }
        return getLevelColumnKey(activeFile.columns, "level1") ?? "";
    }, [activeFile, aiConfig.cleaning.biochem_level1_refine.outputMappings]);

    const level1ColumnKey = useMemo(
        () =>
            activeFile
                ? (getLevelColumnKey(activeFile.columns, "level1") ?? "")
                : "",
        [activeFile],
    );

    const onUpdateBiochemLevel1Discipline = async (discipline: string) => {
        if (!selectedRow || !activeFile) {
            return;
        }

        const nextDiscipline = discipline.trim();
        if (nextDiscipline.length === 0) {
            return;
        }

        const currentResult =
            selectedRow.cleaningResults?.biochem_level1_refine ?? null;
        const parsed =
            parseAIResultJSON(currentResult?.parsedJsonText ?? "") ??
            parseAIResultJSON(currentResult?.responseText ?? "") ??
            {};
        const nextParsed = {
            ...(parsed && typeof parsed === "object" ? parsed : {}),
            discipline: nextDiscipline,
            confidence:
                parsed && typeof parsed === "object" && "confidence" in parsed
                    ? parsed.confidence
                    : "",
            reason:
                parsed && typeof parsed === "object" && "reason" in parsed
                    ? parsed.reason
                    : "",
        };
        const nextParsedJsonText = JSON.stringify(nextParsed);
        const nextResponseText = JSON.stringify(nextParsed, null, 2);
        const response = await fetch(
            `/api/files/${encodeURIComponent(activeFile.fileId)}/cleaning-results/biochem_level1_refine`,
            {
                method: "PUT",
                headers: {
                    "Content-Type": "application/json",
                },
                body: JSON.stringify({
                    rowId: selectedRow.rowId,
                    fileName: activeFile.fileName,
                    responseText: nextResponseText,
                    parsedJsonText: nextParsedJsonText,
                }),
            },
        );
        if (!response.ok) {
            const payload = (await response.json().catch(() => ({}))) as {
                message?: string;
            };
            const message = payload.message ?? "更新生化 Level1 失败";
            setErrorMessage(message);
            throw new Error(message);
        }

        const mappedFieldValues: Record<string, string> = {};
        if (biochemLevel1FieldKey.length > 0) {
            mappedFieldValues[biochemLevel1FieldKey] = nextDiscipline;
        }
        if (level1ColumnKey.length > 0) {
            mappedFieldValues[level1ColumnKey] = nextDiscipline;
        }
        updateRowCleaningResult(
            activeFile.fileId,
            selectedRow.rowId,
            "biochem_level1_refine",
            {
                responseText: nextResponseText,
                parsedJsonText: nextParsedJsonText,
                updatedAt: new Date().toISOString(),
            },
            Object.keys(mappedFieldValues).length > 0
                ? mappedFieldValues
                : undefined,
        );
    };

    const { renderDetailField, renderListReadonlyCell, getListCellTitle } =
        useCellRenderers({
            selectedRow,
            level3TagsFieldKey,
            latexRenderOverrides,
            onToggleLatexRender,
            onToggleDisplayColumn,
            onEditCell,
            getLatexToggleKey,
            setPreviewImageSrc,
        });

    const isDetailView = activeSection === "list" && selectedRowId !== null;
    const isDetailChatSidebarVisible =
        isDetailView && activeFile !== null && !isDetailChatSidebarHidden;
    const showWorkspaceTopbar = isDetailView || activeSection !== "dashboard";

    const startResizeDetailChatSidebar = (
        event: ReactMouseEvent<HTMLButtonElement>,
    ) => {
        detailChatResizeRef.current = {
            active: true,
            startX: event.clientX,
            startWidth: detailChatSidebarWidth,
        };
        event.preventDefault();
    };

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
                onExportFile={onExportFile}
                onUploadClick={onUploadClick}
                uploadInputRef={uploadInputRef}
                onUploadFile={onUploadFile}
                isExporting={isExporting}
                isUploading={isUploading}
                activeFile={activeFile}
            />

            {/* ─── Main Content ─── */}
            <main
                className={`main-content main-workspace ${isSidebarCollapsed ? "sidebar-collapsed" : ""} ${isDetailChatSidebarVisible ? "detail-chat-visible" : ""}`}
                style={
                    {
                        "--detail-chat-sidebar-width": `${detailChatSidebarWidth}px`,
                    } as CSSProperties
                }
            >
                <WorkspaceSidebar
                    isCollapsed={isSidebarCollapsed}
                    activeSection={activeSection}
                    activeSettingsSection={activeSettingsSection}
                    activeFile={activeFile}
                    onToggle={() =>
                        setIsSidebarCollapsed((previous) => !previous)
                    }
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
                            {showWorkspaceTopbar ? (
                                <section className="workspace-topbar">
                                    <div className="toolbar page-toolbar">
                                        {activeSection === "list" &&
                                        !isDetailView ? (
                                            <div className="list-toolbar">
                                                <div className="filter-bar">
                                                    <button
                                                        type="button"
                                                        className={`btn ${activeFile.filterConditions.length > 0 ? "btn-primary" : ""}`}
                                                        onClick={() =>
                                                            setIsFilterModalOpen(
                                                                true,
                                                            )
                                                        }
                                                    >
                                                        {activeFile
                                                            .filterConditions
                                                            .length > 0
                                                            ? `筛选条件 ${activeFile.filterConditions.length}`
                                                            : "添加筛选"}
                                                    </button>
                                                    {activeFile.filterConditions
                                                        .length > 0 ? (
                                                        <button
                                                            type="button"
                                                            className="btn"
                                                            onClick={
                                                                onClearFilterConditions
                                                            }
                                                        >
                                                            清空筛选
                                                        </button>
                                                    ) : null}
                                                </div>
                                                <div className="batch-bar">
                                                    <div className="batch-bar-controls">
                                                        <label className="batch-control">
                                                            <span>AI工具</span>
                                                            <select
                                                                value={
                                                                    activeAIRunKey
                                                                }
                                                                onChange={(
                                                                    event,
                                                                ) =>
                                                                    setActiveAIRunKey(
                                                                        event
                                                                            .target
                                                                            .value as typeof activeAIRunKey,
                                                                    )
                                                                }
                                                                disabled={
                                                                    aiConfigLoading ||
                                                                    isAIDetecting ||
                                                                    isAICleaning ||
                                                                    isAIBatchRunning
                                                                }
                                                            >
                                                                <optgroup label="AI检测">
                                                                    {AI_RUN_STAGE_ORDER.map(
                                                                        (
                                                                            stageKey,
                                                                        ) => (
                                                                            <option
                                                                                key={
                                                                                    stageKey
                                                                                }
                                                                                value={
                                                                                    stageKey
                                                                                }
                                                                            >
                                                                                {stageKey ===
                                                                                AI_RUN_ALL_KEY
                                                                                    ? AI_RUN_ALL_LABEL
                                                                                    : (AI_STAGE_LABELS[
                                                                                          stageKey
                                                                                      ]
                                                                                          ?.shortTitle ??
                                                                                      stageKey)}
                                                                            </option>
                                                                        ),
                                                                    )}
                                                                </optgroup>
                                                                <optgroup label="数据清洗">
                                                                    {AI_CLEANING_TOOL_ORDER.map(
                                                                        (
                                                                            toolKey,
                                                                        ) => (
                                                                            <option
                                                                                key={
                                                                                    toolKey
                                                                                }
                                                                                value={
                                                                                    toolKey
                                                                                }
                                                                            >
                                                                                {
                                                                                    AI_CLEANING_TOOL_LABELS[
                                                                                        toolKey
                                                                                    ]
                                                                                        .shortTitle
                                                                                }
                                                                            </option>
                                                                        ),
                                                                    )}
                                                                </optgroup>
                                                            </select>
                                                        </label>
                                                        <label className="batch-control">
                                                            <span>
                                                                批量并发
                                                            </span>
                                                            <input
                                                                type="number"
                                                                min={
                                                                    MIN_AI_BATCH_CONCURRENCY
                                                                }
                                                                max={
                                                                    MAX_AI_BATCH_CONCURRENCY
                                                                }
                                                                step={1}
                                                                value={
                                                                    aiBatchConcurrency
                                                                }
                                                                onChange={(
                                                                    event,
                                                                ) =>
                                                                    setAIBatchConcurrency(
                                                                        normalizeAIBatchConcurrency(
                                                                            Number(
                                                                                event
                                                                                    .target
                                                                                    .value,
                                                                            ),
                                                                        ),
                                                                    )
                                                                }
                                                                disabled={
                                                                    isAIBatchRunning
                                                                }
                                                            />
                                                        </label>
                                                    </div>
                                                    <div className="batch-bar-actions">
                                                        <button
                                                            type="button"
                                                            className="btn"
                                                            onClick={
                                                                onSelectCurrentPageBatchRows
                                                            }
                                                            disabled={
                                                                paginatedRows.length ===
                                                                    0 ||
                                                                isAIBatchRunning
                                                            }
                                                        >
                                                            仅选当前页
                                                        </button>
                                                        <button
                                                            type="button"
                                                            className="btn"
                                                            onClick={
                                                                onSelectAllBatchRows
                                                            }
                                                            disabled={
                                                                visibleRows.length ===
                                                                    0 ||
                                                                isAIBatchRunning
                                                            }
                                                        >
                                                            {batchSelectedRowIds.length ===
                                                                visibleRows.length &&
                                                            visibleRows.length >
                                                                0
                                                                ? "取消全选"
                                                                : "全选筛选结果"}
                                                        </button>
                                                        <button
                                                            type="button"
                                                            className="btn"
                                                            onClick={
                                                                onClearBatchRows
                                                            }
                                                            disabled={
                                                                batchSelectedRowIds.length ===
                                                                    0 ||
                                                                isAIBatchRunning
                                                            }
                                                        >
                                                            清空勾选
                                                        </button>
                                                        <button
                                                            type="button"
                                                            className="btn btn-primary"
                                                            onClick={
                                                                onRunSelectedBatchAIAnswer
                                                            }
                                                            disabled={
                                                                aiConfigLoading ||
                                                                isAIDetecting ||
                                                                isAIChatting ||
                                                                isAICleaning ||
                                                                isAIBatchRunning ||
                                                                batchSelectedRowIds.length ===
                                                                    0
                                                            }
                                                        >
                                                            {isAIBatchRunning
                                                                ? "AI批量运行中..."
                                                                : `批量运行已选 ${batchSelectedRowIds.length} 条`}
                                                        </button>
                                                    </div>
                                                </div>
                                            </div>
                                        ) : null}
                                        {isDetailView ? (
                                            <div className="toolbar-actions">
                                                <button
                                                    type="button"
                                                    className="btn"
                                                    onClick={() =>
                                                        navigateToSection(
                                                            "list",
                                                        )
                                                    }
                                                >
                                                    返回列表
                                                </button>
                                                <button
                                                    type="button"
                                                    className="btn"
                                                    onClick={() =>
                                                        previousRow &&
                                                        openRowDetail(
                                                            previousRow.rowId,
                                                        )
                                                    }
                                                    disabled={!previousRow}
                                                >
                                                    上一题
                                                </button>
                                                <button
                                                    type="button"
                                                    className="btn"
                                                    onClick={() =>
                                                        nextRow &&
                                                        openRowDetail(
                                                            nextRow.rowId,
                                                        )
                                                    }
                                                    disabled={!nextRow}
                                                >
                                                    下一题
                                                </button>
                                                {!isDetailChatSidebarVisible ? (
                                                    <button
                                                        type="button"
                                                        className="btn btn-ghost detail-chat-open-btn detail-chat-open-toolbar"
                                                        onClick={() =>
                                                            setIsDetailChatSidebarHidden(
                                                                false,
                                                            )
                                                        }
                                                        aria-label="显示 AI 聊天"
                                                        title="显示 AI 聊天"
                                                    >
                                                        <IconMessageSquare />
                                                    </button>
                                                ) : null}
                                            </div>
                                        ) : null}
                                        {activeSection === "settings" ? (
                                            <div className="settings-tabs">
                                                <button
                                                    type="button"
                                                    className={`btn ${activeSettingsSection === "fields" ? "btn-primary" : ""}`}
                                                    onClick={() =>
                                                        navigateToSection(
                                                            "settings",
                                                            "fields",
                                                        )
                                                    }
                                                >
                                                    字段设置
                                                </button>
                                                <button
                                                    type="button"
                                                    className={`btn ${activeSettingsSection === "statistics" ? "btn-primary" : ""}`}
                                                    onClick={() =>
                                                        navigateToSection(
                                                            "settings",
                                                            "statistics",
                                                        )
                                                    }
                                                >
                                                    统计设置
                                                </button>
                                                <button
                                                    type="button"
                                                    className={`btn ${activeSettingsSection === "ai" ? "btn-primary" : ""}`}
                                                    onClick={() =>
                                                        navigateToSection(
                                                            "settings",
                                                            "ai",
                                                        )
                                                    }
                                                >
                                                    AI 设置
                                                </button>
                                            </div>
                                        ) : null}
                                    </div>
                                </section>
                            ) : null}

                            <section className="workspace-view">
                                {activeSection === "dashboard" ? (
                                    <section className="page-panel">
                                        <DashboardPage
                                            files={files}
                                            activeFile={activeFile}
                                            onSelectFile={setActiveFileId}
                                            onOpenStatisticsSettings={() =>
                                                navigateToSection(
                                                    "settings",
                                                    "statistics",
                                                )
                                            }
                                        />
                                    </section>
                                ) : null}

                                {activeSection === "list" && !isDetailView ? (
                                    <section className="page-panel">
                                        <ListPage
                                            activeFile={activeFile}
                                            visibleRows={visibleRows}
                                            paginatedRows={paginatedRows}
                                            listPage={listPage}
                                            listPageSize={listPageSize}
                                            totalListPages={totalListPages}
                                            listPageSizeOptions={
                                                LIST_PAGE_SIZE_OPTIONS
                                            }
                                            batchSelectedRowIdSet={
                                                batchSelectedRowIdSet
                                            }
                                            selectedRowId={selectedRowId}
                                            rowStreamProgress={
                                                rowStreamProgress
                                            }
                                            isAIBatchRunning={isAIBatchRunning}
                                            activeAIRunKey={activeAIRunKey}
                                            rowBatchStatuses={rowBatchStatuses}
                                            onToggleBatchRowSelection={
                                                onToggleBatchRowSelection
                                            }
                                            onOpenRowDetail={openRowDetail}
                                            onPageChange={setListPage}
                                            onPageSizeChange={setListPageSize}
                                            getCellTitle={getListCellTitle}
                                            renderListReadonlyCell={
                                                renderListReadonlyCell
                                            }
                                        />
                                    </section>
                                ) : null}

                                {isDetailView ? (
                                    <section className="page-panel detail-page-panel">
                                        <DetailPage
                                            selectedRow={selectedRow}
                                            level1Options={
                                                activeFile?.level1Options ?? []
                                            }
                                            displayColumns={displayColumns}
                                            hiddenColumns={hiddenColumns}
                                            showHiddenFields={showHiddenFields}
                                            onToggleHiddenFields={() =>
                                                setShowHiddenFields(
                                                    (previous) => !previous,
                                                )
                                            }
                                            onOpenAIRunModal={
                                                openAIRunModalForStage
                                            }
                                            onRunAllAIDetect={onRunAllAIDetect}
                                            canRunAllAIDetect={canRunAIDetect}
                                            runAllTimerText={runAllTimerText}
                                            runAllStageTimers={
                                                runAllStageTimers
                                            }
                                            renderDetailField={
                                                renderDetailField
                                            }
                                            aiResults={selectedRow?.aiResults}
                                            cleaningResults={
                                                selectedRow?.cleaningResults
                                            }
                                            isAICleaning={isAICleaning}
                                            activeAICleaningToolKey={
                                                activeAICleaningToolKey
                                            }
                                            aiCleaningElapsedText={
                                                aiCleaningElapsedText
                                            }
                                            aiCleaningStreamText={
                                                aiCleaningStreamText
                                            }
                                            aiCleaningStatusMessage={
                                                aiCleaningStatusMessage
                                            }
                                            onAddLevel3Tag={onAddLevel3Tag}
                                            onRemoveLevel3Tag={
                                                onRemoveLevel3Tag
                                            }
                                            onUpdateBiochemLevel1Discipline={
                                                onUpdateBiochemLevel1Discipline
                                            }
                                            onRunAICleaning={onRunAICleaning}
                                            onToggleRowEnabled={
                                                onToggleRowEnabled
                                            }
                                        />
                                    </section>
                                ) : null}

                                {activeSection === "settings" ? (
                                    <section className="page-panel settings-page-panel">
                                        <SettingsPage
                                            activeSettingsSection={
                                                activeSettingsSection
                                            }
                                            activeFile={activeFile}
                                            displayColumns={displayColumns}
                                            aiConfigList={aiConfigList}
                                            aiConfig={aiConfig}
                                            onOpenActiveFileConfig={
                                                onOpenActiveFileConfig
                                            }
                                            onToggleStatisticsField={
                                                onToggleStatisticsField
                                            }
                                            onSetStatisticsChartType={
                                                onSetStatisticsChartType
                                            }
                                            onOpenAIStageConfigModal={
                                                onOpenAIStageConfigModal
                                            }
                                            onOpenAIProfileModal={
                                                onOpenAIProfileModal
                                            }
                                            onOpenAIRouteModal={
                                                onOpenAIRouteModal
                                            }
                                            onOpenAIChatConfigModal={
                                                onOpenAIChatConfigModal
                                            }
                                            onOpenAICleaningConfigModal={
                                                onOpenAICleaningConfigModal
                                            }
                                        />
                                    </section>
                                ) : null}
                            </section>
                        </>
                    )}
                </section>

                {isDetailChatSidebarVisible ? (
                    <div className="detail-chat-shell">
                        <button
                            type="button"
                            className="detail-chat-resize-handle"
                            onMouseDown={startResizeDetailChatSidebar}
                            aria-label="拖动调整 AI 聊天侧边栏宽度"
                            title="拖动调整 AI 聊天侧边栏宽度"
                        >
                            <span />
                        </button>
                        <AIChatSidebar
                            routes={aiConfig.routes}
                            activeRouteName={activeChatRouteName}
                            chatMessages={chatMessages}
                            chatInput={chatInput}
                            chatStatusMessage={chatStatusMessage}
                            isAIChatting={isAIChatting}
                            aiChatElapsedText={aiChatElapsedText}
                            onHide={() => setIsDetailChatSidebarHidden(true)}
                            onRouteChange={setActiveChatRouteName}
                            onInputChange={setChatInput}
                            onSend={onSendAIChatMessage}
                            onClear={onClearAIChatSession}
                        />
                    </div>
                ) : null}
            </main>
            {/* ─── Column Selection Modal ─── */}
            <ColumnConfigModal
                pendingFile={pendingFile}
                pendingConfigMode={pendingConfigMode}
                pendingConfigNotice={pendingConfigNotice}
                pendingSelectedDisplayKeys={pendingSelectedDisplayKeys}
                pendingEditableColumnKeys={pendingEditableColumnKeys}
                onPendingSelectAllDisplayColumns={
                    onPendingSelectAllDisplayColumns
                }
                onPendingClearDisplayColumns={onPendingClearDisplayColumns}
                onPendingClearEditableColumns={onPendingClearEditableColumns}
                onTogglePendingDisplayColumn={onTogglePendingDisplayColumn}
                onTogglePendingEditableColumn={onTogglePendingEditableColumn}
                onCancelPendingFile={onCancelPendingFile}
                onConfirmPendingFile={onConfirmPendingFile}
            />
            <FilterConfigModal
                isOpen={isFilterModalOpen}
                activeFile={activeFile}
                onClose={() => setIsFilterModalOpen(false)}
                onSave={onUpdateFilterConditions}
            />

            {/* ─── AI Stage Config Modal ─── */}
            <AIStageConfigModal
                isOpen={isAIStageConfigModalOpen}
                activeFile={activeFile}
                aiConfigFormMessage={aiConfigFormMessage}
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
                draftAIConfig={draftAIConfig}
                setDraftAIConfig={setDraftAIConfig}
                aiConfigSaving={aiConfigSaving}
                onCancel={onCancelAIProfileModal}
                onSave={onSaveAIProfileConfig}
            />
            <AIRouteModal
                isOpen={isAIRouteModalOpen}
                activeFile={activeFile}
                aiConfigFormMessage={aiConfigFormMessage}
                draftAIConfig={draftAIConfig}
                setDraftAIConfig={setDraftAIConfig}
                aiConfigSaving={aiConfigSaving}
                onCancel={onCancelAIRouteModal}
                onSave={onSaveAIRouteConfig}
            />
            <AIChatConfigModal
                isOpen={isAIChatConfigModalOpen}
                activeFile={activeFile}
                aiConfigFormMessage={aiConfigFormMessage}
                draftAIConfig={draftAIConfig}
                setDraftAIConfig={setDraftAIConfig}
                aiSubmitFieldColumns={aiSubmitFieldColumns}
                aiConfigSaving={aiConfigSaving}
                onToggleDraftAIChatSubmitField={onToggleDraftAIChatSubmitField}
                onCancel={onCancelAIChatConfigModal}
                onSave={onSaveAIChatConfig}
            />
            <AICleaningConfigModal
                isOpen={isAICleaningConfigModalOpen}
                activeFile={activeFile}
                aiConfigFormMessage={aiConfigFormMessage}
                draftAIConfig={draftAIConfig}
                setDraftAIConfig={setDraftAIConfig}
                aiSubmitFieldColumns={aiSubmitFieldColumns}
                aiConfigSaving={aiConfigSaving}
                onToggleDraftAICleaningSubmitField={
                    onToggleDraftAICleaningSubmitField
                }
                onUpdateDraftAICleaningOutputMapping={
                    onUpdateDraftAICleaningOutputMapping
                }
                onCancel={onCancelAICleaningConfigModal}
                onSave={onSaveAICleaningConfig}
            />
            <AIRunModal
                isOpen={isAIRunModalOpen}
                rowId={selectedRow?.rowId}
                aiStageKey={modalStageKey}
                isAIDetecting={isAIDetecting}
                aiDetectElapsedText={aiDetectElapsedText}
                canRunAIDetect={canRunAIDetect}
                onRunAIDetect={() => onRunAIDetect(modalStageKey)}
                aiMergedStreamText={aiMergedStreamText}
                onAIResultTextChange={onAIResultTextChange}
                aiResultMessage={aiResultMessage}
                aiFieldLabels={modalFieldLabels}
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
