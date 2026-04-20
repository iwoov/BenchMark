import type { ChangeEvent, RefObject } from "react";
import type { FileViewState } from "../../types";
import type { AIBatchTaskState } from "../types";
import { getAIBatchTaskStatusText } from "../ai-helpers";
import { IconBrand, IconDownload, IconMoon, IconSun } from "../icons";

interface HeaderBarProps {
    files: FileViewState[];
    activeFileId: string | null;
    onSelectFile: (fileId: string) => void;
    errorMessage: string;
    aiBatchTask: AIBatchTaskState;
    aiBatchProgressPercent: number;
    isAIBatchRunning: boolean;
    theme: "dark" | "light";
    onToggleTheme: () => void;
    onExportFile: () => void;
    uploadInputRef: RefObject<HTMLInputElement>;
    onUploadFile: (event: ChangeEvent<HTMLInputElement>) => void;
    isExporting: boolean;
    activeFile: FileViewState | null;
}

export function HeaderBar({
    files,
    activeFileId,
    onSelectFile,
    errorMessage,
    aiBatchTask,
    aiBatchProgressPercent,
    isAIBatchRunning,
    theme,
    onToggleTheme,
    onExportFile,
    uploadInputRef,
    onUploadFile,
    isExporting,
    activeFile,
}: HeaderBarProps) {
    // Determine data sources that share the same project as activeFile
    const activeProjectId = activeFile?.projectId ?? activeFile?.fileId ?? null;
    const siblingDataSources =
        activeProjectId !== null
            ? files.filter((f) => (f.projectId ?? f.fileId) === activeProjectId)
            : [];
    const showDataSourcePicker = siblingDataSources.length > 1;

    return (
        <header className="header-bar">
            <div className="header-inner">
                <div className="header-brand">
                    <div className="brand-icon">
                        <IconBrand />
                    </div>
                    <div className="header-brand-copy">
                        <h1>质检工作台</h1>
                        <span>项目驱动的数据质检与评测工作区</span>
                    </div>
                </div>

                <div className="header-project-picker">
                    <select
                        className="header-project-select"
                        value={activeProjectId ?? ""}
                        onChange={(event) => {
                            // Switch to the first datasource of the selected project
                            const targetProjectId = event.target.value;
                            const first = files.find(
                                (f) =>
                                    (f.projectId ?? f.fileId) ===
                                    targetProjectId,
                            );
                            if (first) {
                                onSelectFile(first.fileId);
                            }
                        }}
                        disabled={files.length === 0}
                    >
                        <option value="" disabled>
                            {files.length === 0 ? "暂无项目" : "请选择项目"}
                        </option>
                        {/* Deduplicate by projectId */}
                        {Array.from(
                            new Map(
                                files.map((f) => [f.projectId ?? f.fileId, f]),
                            ).values(),
                        ).map((f) => (
                            <option
                                key={f.projectId ?? f.fileId}
                                value={f.projectId ?? f.fileId}
                            >
                                {f.fileName}
                            </option>
                        ))}
                    </select>
                    {showDataSourcePicker ? (
                        <select
                            className="header-project-select header-datasource-select"
                            value={activeFileId ?? ""}
                            onChange={(event) =>
                                onSelectFile(event.target.value)
                            }
                        >
                            {siblingDataSources.map((f) => (
                                <option key={f.fileId} value={f.fileId}>
                                    {f.dataSourceName
                                        ? f.dataSourceName
                                        : `数据源 ${f.fileId.slice(0, 6)}`}
                                </option>
                            ))}
                        </select>
                    ) : null}
                </div>

                <div className="header-actions">
                    {errorMessage ? (
                        <span className="error-text">{errorMessage}</span>
                    ) : null}
                    {aiBatchTask.total > 0 ? (
                        <div
                            className={`ai-batch-status ${isAIBatchRunning ? "running" : "completed"}`}
                        >
                            <div className="ai-batch-status-head">
                                <span>
                                    {getAIBatchTaskStatusText(aiBatchTask)}
                                </span>
                                <strong>
                                    {aiBatchTask.completed}/{aiBatchTask.total}
                                </strong>
                            </div>
                            <div className="ai-batch-progress">
                                <div
                                    className="ai-batch-progress-bar"
                                    style={{
                                        width: `${aiBatchProgressPercent}%`,
                                    }}
                                />
                            </div>
                            <div className="ai-batch-counts">
                                <span className="ai-batch-success">
                                    成功 {aiBatchTask.success}
                                </span>
                                <span className="ai-batch-failed">
                                    失败 {aiBatchTask.failed}
                                </span>
                                <span>{aiBatchProgressPercent}%</span>
                            </div>
                            {aiBatchTask.fileName ? (
                                <div
                                    className="ai-batch-file"
                                    title={aiBatchTask.fileName}
                                >
                                    {`任务文件：${aiBatchTask.fileName}`}
                                </div>
                            ) : null}
                            {aiBatchTask.message ? (
                                <div className="ai-batch-message">
                                    {aiBatchTask.message}
                                </div>
                            ) : null}
                        </div>
                    ) : null}
                    <button
                        type="button"
                        className="theme-toggle"
                        onClick={onToggleTheme}
                        title={
                            theme === "dark" ? "切换浅色主题" : "切换深色主题"
                        }
                    >
                        {theme === "dark" ? <IconSun /> : <IconMoon />}
                    </button>
                    <button
                        type="button"
                        className="btn"
                        onClick={onExportFile}
                        disabled={isExporting || !activeFile}
                    >
                        <IconDownload />
                        {isExporting ? "导出中..." : "导出 Excel"}
                    </button>
                    <input
                        ref={uploadInputRef}
                        type="file"
                        accept=".xls,.xlsx,.json"
                        className="hidden-input"
                        onChange={onUploadFile}
                    />
                </div>
            </div>
        </header>
    );
}
