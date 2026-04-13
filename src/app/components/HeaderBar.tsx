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
                        value={activeFileId ?? ""}
                        onChange={(event) => onSelectFile(event.target.value)}
                        disabled={files.length === 0}
                    >
                        <option value="" disabled>
                            {files.length === 0 ? "暂无项目" : "请选择项目"}
                        </option>
                        {files.map((file) => (
                            <option key={file.fileId} value={file.fileId}>
                                {file.fileName}
                            </option>
                        ))}
                    </select>
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
