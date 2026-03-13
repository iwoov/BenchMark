import { useEffect, useState } from "react";
import type { AIDetectStageKey } from "../../types";
import { AI_STAGE_LABELS, AI_STAGE_ORDER } from "../constants";

type AIRunTab = "response" | "request";

interface AIRunModalProps {
  isOpen: boolean;
  rowId?: string;
  aiStageKey: AIDetectStageKey;
  onSelectAIStage: (stageKey: AIDetectStageKey) => void;
  aiConfigLoading: boolean;
  isAIDetecting: boolean;
  isAIBatchRunning: boolean;
  aiDetectElapsedText: string;
  canRunAIDetect: boolean;
  onRunAIDetect: () => void;
  aiRetryCount: number;
  aiMergedStreamText: string;
  onAIResultTextChange: (value: string) => void;
  aiResultMessage: string;
  aiRequestPreview: string;
  onClose: () => void;
}

export function AIRunModal({
  isOpen,
  rowId,
  aiStageKey,
  onSelectAIStage,
  aiConfigLoading,
  isAIDetecting,
  isAIBatchRunning,
  aiDetectElapsedText,
  canRunAIDetect,
  onRunAIDetect,
  aiRetryCount,
  aiMergedStreamText,
  onAIResultTextChange,
  aiResultMessage,
  aiRequestPreview,
  onClose,
}: AIRunModalProps) {
  const [activeTab, setActiveTab] = useState<AIRunTab>("response");

  useEffect(() => {
    if (isOpen) {
      setActiveTab("response");
    }
  }, [isOpen]);

  if (!isOpen) {
    return null;
  }

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

        <div className="ai-run-controls">
          <label className="ai-run-config">
            <span>运行阶段</span>
            <select
              value={aiStageKey}
              onChange={(event) =>
                onSelectAIStage(event.target.value as AIDetectStageKey)
              }
              disabled={aiConfigLoading || isAIDetecting || isAIBatchRunning}
            >
              {AI_STAGE_ORDER.map((stageKey) => (
                <option key={stageKey} value={stageKey}>
                  {AI_STAGE_LABELS[stageKey]?.shortTitle ?? stageKey}
                </option>
              ))}
            </select>
          </label>
          <button
            type="button"
            className="btn btn-primary"
            onClick={onRunAIDetect}
            disabled={!canRunAIDetect}
          >
            {isAIDetecting ? `AI回答中 ${aiDetectElapsedText}` : "运行AI阶段"}
          </button>
          <div className="ai-result-target">
            <span>重试：</span>
            <strong className="ai-retry-count">{aiRetryCount}次</strong>
          </div>
        </div>

        <div className="ai-run-tabs">
          <button
            type="button"
            className={`btn ${activeTab === "response" ? "btn-primary" : ""}`}
            onClick={() => setActiveTab("response")}
          >
            响应结果
          </button>
          <button
            type="button"
            className={`btn ${activeTab === "request" ? "btn-primary" : ""}`}
            onClick={() => setActiveTab("request")}
          >
            请求详情
          </button>
        </div>

        <div className="ai-run-panel">
          {activeTab === "response" ? (
            <textarea
              className="ai-preview-textarea"
              value={aiMergedStreamText}
              onChange={(event) => onAIResultTextChange(event.target.value)}
              placeholder="点击“运行AI阶段”后，这里会显示响应内容，可手动复制。"
            />
          ) : (
            <pre className="ai-preview-content">
              {aiRequestPreview.trim().length > 0
                ? aiRequestPreview
                : "暂无内容"}
            </pre>
          )}
        </div>

        {aiResultMessage ? (
          <div className="ai-stream-message">{aiResultMessage}</div>
        ) : null}

        <div className="column-modal-footer">
          <button type="button" className="btn" onClick={onClose}>
            关闭
          </button>
        </div>
      </div>
    </div>
  );
}
