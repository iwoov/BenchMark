import { useEffect, useMemo, useState } from "react";
import {
    AI_RUN_ALL_KEY,
    AI_RUN_ALL_LABEL,
    AI_RUN_STAGE_ORDER,
    AI_STAGE_LABELS,
    LIST_PAGE_SIZE_OPTIONS,
    MAX_AI_BATCH_CONCURRENCY,
    MIN_AI_BATCH_CONCURRENCY,
} from "./app/constants";
import { normalizeAIBatchConcurrency } from "./app/ai-helpers";
import { HeaderBar } from "./app/components/HeaderBar";
import { WorkspaceSidebar } from "./app/components/WorkspaceSidebar";
import { ListPage } from "./app/components/ListPage";
import { DetailPage } from "./app/components/DetailPage";
import { SettingsPage } from "./app/components/SettingsPage";
import { ColumnConfigModal } from "./app/components/ColumnConfigModal";
import { AIProfileModal } from "./app/components/AIProfileModal";
import { AIRouteModal } from "./app/components/AIRouteModal";
import { AIStageConfigModal } from "./app/components/AIStageConfigModal";
import { AIRunModal } from "./app/components/AIRunModal";
import { ImageLightbox } from "./app/components/ImageLightbox";
import { FilterConfigModal } from "./app/components/FilterConfigModal";
import { IconFile } from "./app/icons";
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
    const [isFilterModalOpen, setIsFilterModalOpen] = useState(false);
    const [latexRenderOverrides, setLatexRenderOverrides] = useState<
        Record<string, boolean>
    >({});
    const [previewImageSrc, setPreviewImageSrc] = useState<string | null>(null);

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
        onEditCell,
        onToggleDisplayColumn,
        onUpdateFilterConditions,
        onClearFilterConditions,
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
        onOpenAIStageConfigModal,
        onOpenAIProfileModal,
        onOpenAIRouteModal,
        onCancelAIStageConfigModal,
        onCancelAIProfileModal,
        onCancelAIRouteModal,
        onToggleDraftAISubmitField,
        onSaveAIStageConfig,
        onSaveAIProfileConfig,
        onSaveAIRouteConfig,
        onRunAIDetect,
        onRunAllAIDetect,
        onRunBatchAIAnswer,
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
        persistFileState,
        flushPendingAIResults,
        latestFileStateRef,
    });

    useEffect(() => {
        if (initialLoadComplete && !activeFile) {
            navigateToSection("list", activeSettingsSection, null, {
                replace: true,
            });
        }
    }, [
        initialLoadComplete,
        activeFile,
        activeSettingsSection,
        navigateToSection,
    ]);

    const aiSubmitFieldColumns = useMemo(
        () => (activeFile ? activeFile.columns : []),
        [activeFile],
    );

    const openRowDetail = (rowId: string) => {
        setSelectedRowId(rowId);
        navigateToSection("detail", activeSettingsSection, rowId);
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

    const { renderDetailField, renderListReadonlyCell, getListCellTitle } =
        useCellRenderers({
            selectedRow,
            latexRenderOverrides,
            onToggleLatexRender,
            onToggleDisplayColumn,
            onEditCell,
            getLatexToggleKey,
            setPreviewImageSrc,
        });

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
                className={`main-content main-workspace ${isSidebarCollapsed ? "sidebar-collapsed" : ""}`}
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
                            <section className="workspace-topbar">
                                <div className="toolbar page-toolbar">
                                    {activeSection === "list" ? (
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
                                                    {activeFile.filterConditions
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
                                                        <span>运行阶段</span>
                                                        <select
                                                            value={
                                                                activeAIRunKey
                                                            }
                                                            onChange={(event) =>
                                                                setActiveAIRunKey(
                                                                    event.target
                                                                        .value as typeof activeAIRunKey,
                                                                )
                                                            }
                                                            disabled={
                                                                aiConfigLoading ||
                                                                isAIDetecting ||
                                                                isAIBatchRunning
                                                            }
                                                        >
                                                            {AI_RUN_STAGE_ORDER.map(
                                                                (stageKey) => (
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
                                                        </select>
                                                    </label>
                                                    <label className="batch-control">
                                                        <span>批量并发</span>
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
                                                            onChange={(event) =>
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
                                                        visibleRows.length > 0
                                                            ? "取消全选"
                                                            : "全选可见"}
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
                                                            isAIBatchRunning ||
                                                            batchSelectedRowIds.length ===
                                                                0
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
                                                onClick={() =>
                                                    navigateToSection("list")
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
                                                    openRowDetail(nextRow.rowId)
                                                }
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

                                {activeSection === "detail" ? (
                                    <section className="page-panel detail-page-panel">
                                        <DetailPage
                                            selectedRow={selectedRow}
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
                                            onOpenAIStageConfigModal={
                                                onOpenAIStageConfigModal
                                            }
                                            onOpenAIProfileModal={
                                                onOpenAIProfileModal
                                            }
                                            onOpenAIRouteModal={
                                                onOpenAIRouteModal
                                            }
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
