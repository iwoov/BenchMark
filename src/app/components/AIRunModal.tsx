import type { AIDetectStageKey } from "../../types";
import { AI_STAGE_LABELS } from "../constants";

interface AIRunModalProps {
  isOpen: boolean;
  rowId?: string;
  aiStageKey: AIDetectStageKey;
  isAIDetecting: boolean;
  aiDetectElapsedText: string;
  canRunAIDetect: boolean;
  onRunAIDetect: () => void;
  aiMergedStreamText: string;
  onAIResultTextChange: (value: string) => void;
  aiResultMessage: string;
  aiFieldLabels: string[];
  onClose: () => void;
}

export function AIRunModal({
  isOpen,
  rowId,
  aiStageKey,
  isAIDetecting,
  aiDetectElapsedText,
  canRunAIDetect,
  onRunAIDetect,
  aiMergedStreamText,
  onAIResultTextChange,
  aiResultMessage,
  aiFieldLabels,
  onClose,
}: AIRunModalProps) {
  if (!isOpen) {
    return null;
  }
  const stageLabel = AI_STAGE_LABELS[aiStageKey]?.shortTitle ?? aiStageKey;
  const stageTitle = AI_STAGE_LABELS[aiStageKey]?.title ?? aiStageKey;

  return (
    <div className="column-modal-mask">
      <div className="column-modal ai-run-modal">
        <div className="ai-preview-head">
          <div className="ai-preview-title">
            <h3>AI 检测</h3>
            {rowId ? <p>{`当前记录：${rowId}`}</p> : null}
          </div>
          <button
            type="button"
            className="ai-preview-close"
            onClick={onClose}
            aria-label="关闭弹窗"
          >
            ×
          </button>
        </div>

        <div className="ai-run-stage-info">
          <span className="ai-run-stage-title">{stageTitle}</span>
          <div className="ai-run-field-tags">
            <span className="ai-run-field-label">提交字段：</span>
            <div className="ai-run-field-list">
              {aiFieldLabels.length > 0 ? (
                aiFieldLabels.map((label) => (
                  <span key={label} className="ai-run-field-tag">
                    {label}
                  </span>
                ))
              ) : (
                <span className="ai-run-field-empty">暂无字段</span>
              )}
            </div>
          </div>
        </div>

        <div className="ai-run-panel">
          <textarea
            className="ai-preview-textarea"
            value={aiMergedStreamText}
            onChange={(event) => onAIResultTextChange(event.target.value)}
            placeholder="点击运行按钮后，这里会显示响应内容，可手动复制。"
          />
        </div>

        {aiResultMessage ? (
          <div className="ai-stream-message">{aiResultMessage}</div>
        ) : null}

        <div className="column-modal-footer">
          <button
            type="button"
            className="btn btn-primary"
            onClick={onRunAIDetect}
            disabled={!canRunAIDetect}
          >
            {isAIDetecting
              ? `AI回答中 ${aiDetectElapsedText}`
              : `运行 ${stageLabel}`}
          </button>
          <button type="button" className="btn" onClick={onClose}>
            关闭
          </button>
        </div>
      </div>
    </div>
  );
}
