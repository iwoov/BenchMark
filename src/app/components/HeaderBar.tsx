import type { ChangeEvent, RefObject } from "react";
import type { FileViewState } from "../../types";
import type { AIBatchTaskState } from "../types";
import { getAIBatchTaskStatusText } from "../ai-helpers";
import { IconDownload, IconMoon, IconSun, IconUpload } from "../icons";

interface HeaderBarProps {
  files: FileViewState[];
  activeFileId: string | null;
  onSelectFile: (fileId: string) => void;
  onRemoveFile: (fileId: string) => void;
  errorMessage: string;
  aiBatchTask: AIBatchTaskState;
  aiBatchProgressPercent: number;
  isAIBatchRunning: boolean;
  theme: "dark" | "light";
  onToggleTheme: () => void;
  onOpenAIStageConfigModal: () => void;
  onExportFile: () => void;
  onUploadClick: () => void;
  uploadInputRef: RefObject<HTMLInputElement>;
  onUploadFile: (event: ChangeEvent<HTMLInputElement>) => void;
  isExporting: boolean;
  isUploading: boolean;
  aiConfigLoading: boolean;
  activeFile: FileViewState | null;
}

export function HeaderBar({
  files,
  activeFileId,
  onSelectFile,
  onRemoveFile,
  errorMessage,
  aiBatchTask,
  aiBatchProgressPercent,
  isAIBatchRunning,
  theme,
  onToggleTheme,
  onOpenAIStageConfigModal,
  onExportFile,
  onUploadClick,
  uploadInputRef,
  onUploadFile,
  isExporting,
  isUploading,
  aiConfigLoading,
  activeFile,
}: HeaderBarProps) {
  return (
    <header className="header-bar">
      <div className="header-inner">
        <div className="header-brand">
          <div className="brand-icon">
            <svg viewBox="0 0 24 24">
              <path
                d="M9 2L4 7v13a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V4a2 2 0 0 0-2-2H9zm0 0v5H4m4 4h8m-8 4h8m-8 4h4"
                fill="none"
                stroke="white"
                strokeWidth="1.5"
                strokeLinecap="round"
                strokeLinejoin="round"
              />
            </svg>
          </div>
          <h1>质检工作台</h1>
        </div>

        <div className="file-tabs">
          {files.map((file) => (
            <div
              key={file.fileId}
              className={`file-tab ${file.fileId === activeFileId ? "active" : ""}`}
            >
              <button
                type="button"
                style={{
                  all: "unset",
                  cursor: "pointer",
                  display: "contents",
                }}
                onClick={() => onSelectFile(file.fileId)}
              >
                {file.fileName}
              </button>
              <span className="tab-badge">{file.rows.length}</span>
              <button
                type="button"
                className="tab-close"
                onClick={(event) => {
                  event.stopPropagation();
                  onRemoveFile(file.fileId);
                }}
                title="关闭"
              >
                ×
              </button>
            </div>
          ))}
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
                <span>{getAIBatchTaskStatusText(aiBatchTask)}</span>
                <strong>
                  {aiBatchTask.completed}/{aiBatchTask.total}
                </strong>
              </div>
              <div className="ai-batch-progress">
                <div
                  className="ai-batch-progress-bar"
                  style={{ width: `${aiBatchProgressPercent}%` }}
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
                <div className="ai-batch-file" title={aiBatchTask.fileName}>
                  {`任务文件：${aiBatchTask.fileName}`}
                </div>
              ) : null}
              {aiBatchTask.message ? (
                <div className="ai-batch-message">{aiBatchTask.message}</div>
              ) : null}
            </div>
          ) : null}
          <button
            type="button"
            className="theme-toggle"
            onClick={onToggleTheme}
            title={theme === "dark" ? "切换浅色主题" : "切换深色主题"}
          >
            {theme === "dark" ? <IconSun /> : <IconMoon />}
          </button>
          <button
            type="button"
            className="btn"
            onClick={onOpenAIStageConfigModal}
            disabled={!activeFile || aiConfigLoading}
          >
            AI阶段配置
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
          <button
            type="button"
            className="btn btn-primary"
            onClick={onUploadClick}
            disabled={isUploading}
          >
            <IconUpload />
            {isUploading ? "解析中..." : "导入 Excel"}
          </button>
          <input
            ref={uploadInputRef}
            type="file"
            accept=".xls,.xlsx"
            className="hidden-input"
            onChange={onUploadFile}
          />
        </div>
      </div>
    </header>
  );
}
