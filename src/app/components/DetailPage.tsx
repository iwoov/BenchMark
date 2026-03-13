import type { ReactNode } from "react";
import type { AIDetectStageKey, ParsedColumn, ParsedRow } from "../../types";
import { IconChevron } from "../icons";
import { AI_STAGE_LABELS, AI_STAGE_ORDER } from "../constants";

interface AIResultParsed {
  // Pre-check
  is_valid?: boolean;
  reason?: string;
  requires_image?: boolean;
  // Context Audit
  is_consistent?: boolean;
  missing_info?: string;
  // Independent Solving
  analysis?: Record<string, string>;
  final_answer?: string;
  // Final Verdict
  status?: "Pass" | "Fail";
  discrepancy_detail?: string;
  [key: string]: unknown;
}

function parseAIResultJSON(content: string): AIResultParsed | null {
  if (!content || content.trim().length === 0) {
    return null;
  }
  // Try to extract JSON from the content (may contain thinking process before JSON)
  const jsonMatch = content.match(/\{[\s\S]*\}/);
  if (!jsonMatch) {
    return null;
  }
  try {
    return JSON.parse(jsonMatch[0]) as AIResultParsed;
  } catch {
    return null;
  }
}

function renderPrecheckResult(parsed: AIResultParsed) {
  const isValid = parsed.is_valid;
  const reason = parsed.reason?.trim() || "";
  const requiresImage = parsed.requires_image;

  return (
    <div className="ai-result-formatted">
      <div className="ai-result-status-row">
        <span
          className={`ai-result-badge ${isValid ? "badge-valid" : "badge-invalid"}`}
        >
          {isValid ? "✓ 合格" : "✗ 不合格"}
        </span>
        {typeof requiresImage === "boolean" && (
          <span
            className={`ai-result-badge ${requiresImage ? "badge-image-required" : "badge-image-optional"}`}
          >
            {requiresImage ? "🖼️ 需要图片" : "无需图片"}
          </span>
        )}
      </div>
      {!isValid && reason && reason !== "无" && (
        <div className="ai-result-reason">
          <strong>不合格原因：</strong>
          <span>{reason}</span>
        </div>
      )}
    </div>
  );
}

function renderContextAuditResult(parsed: AIResultParsed) {
  const isConsistent = parsed.is_consistent;
  const missingInfo = parsed.missing_info?.trim() || "";

  return (
    <div className="ai-result-formatted">
      <div className="ai-result-status-row">
        <span
          className={`ai-result-badge ${isConsistent ? "badge-valid" : "badge-invalid"}`}
        >
          {isConsistent ? "✓ 一致" : "✗ 不一致"}
        </span>
      </div>
      {!isConsistent && missingInfo && missingInfo !== "无" && (
        <div className="ai-result-reason">
          <strong>缺失/矛盾信息：</strong>
          <span>{missingInfo}</span>
        </div>
      )}
    </div>
  );
}

function renderIndependentSolvingResult(parsed: AIResultParsed) {
  const finalAnswer = parsed.final_answer?.trim() || "无法确定";

  return (
    <div className="ai-result-formatted">
      <div className="ai-result-status-row">
        <span className="ai-result-badge badge-answer">
          📝 答案：{finalAnswer}
        </span>
      </div>
    </div>
  );
}

function renderFinalVerdictResult(parsed: AIResultParsed) {
  const status = parsed.status;
  const isPass = status === "Pass";
  const discrepancy = parsed.discrepancy_detail?.trim() || "";

  return (
    <div className="ai-result-formatted">
      <div className="ai-result-status-row">
        <span
          className={`ai-result-badge ${isPass ? "badge-valid" : "badge-invalid"}`}
        >
          {isPass ? "✓ Pass" : "✗ Fail"}
        </span>
      </div>
      {!isPass && discrepancy && discrepancy !== "无" && (
        <div className="ai-result-reason">
          <strong>差异说明：</strong>
          <span>{discrepancy}</span>
        </div>
      )}
    </div>
  );
}

function renderAIResultContent(stageKey: AIDetectStageKey, content: string) {
  const trimmed = content.trim();
  if (trimmed.length === 0) {
    return <span className="ai-result-empty">暂无结果</span>;
  }

  const parsed = parseAIResultJSON(trimmed);
  if (!parsed) {
    return <pre className="ai-result-raw">{trimmed}</pre>;
  }

  switch (stageKey) {
    case "precheck":
      if (typeof parsed.is_valid === "boolean") {
        return renderPrecheckResult(parsed);
      }
      break;
    case "context_audit":
      if (typeof parsed.is_consistent === "boolean") {
        return renderContextAuditResult(parsed);
      }
      break;
    case "independent_solving":
      if (parsed.final_answer !== undefined) {
        return renderIndependentSolvingResult(parsed);
      }
      break;
    case "final_verdict":
      if (parsed.status === "Pass" || parsed.status === "Fail") {
        return renderFinalVerdictResult(parsed);
      }
      break;
  }

  // Fallback: show raw content
  return <pre className="ai-result-raw">{trimmed}</pre>;
}

interface DetailPageProps {
  selectedRow: ParsedRow | null;
  displayColumns: ParsedColumn[];
  hiddenColumns: ParsedColumn[];
  showHiddenFields: boolean;
  onToggleHiddenFields: () => void;
  onOpenAIRunModal: () => void;
  renderDetailField: (column: ParsedColumn, isHidden: boolean) => ReactNode;
  aiResults?: Partial<Record<AIDetectStageKey, string>>;
}

export function DetailPage({
  selectedRow,
  displayColumns,
  hiddenColumns,
  showHiddenFields,
  onToggleHiddenFields,
  onOpenAIRunModal,
  renderDetailField,
  aiResults,
}: DetailPageProps) {
  if (!selectedRow) {
    return (
      <div className="record-list-empty">请先在题目列表页选择一条记录</div>
    );
  }

  return (
    <section className="record-detail standalone-record-detail">
      <div className="record-detail-header">
        <h3>字段详情</h3>
        <span>点击字段左侧勾选框可控制显示/隐藏</span>
      </div>
      <div className="record-detail-ai-toolbar">
        <div className="record-detail-ai-actions">
          <button
            type="button"
            className="btn btn-primary"
            onClick={onOpenAIRunModal}
          >
            运行AI检测
          </button>
        </div>
        <div className="record-detail-ai-results">
          <div className="record-detail-ai-results-head">
            <strong>AI检测结果</strong>
            <span>四阶段结果仅保存到数据库</span>
          </div>
          <div className="record-detail-ai-results-grid">
            {AI_STAGE_ORDER.map((stageKey) => {
              const label = AI_STAGE_LABELS[stageKey];
              const content = aiResults?.[stageKey] ?? "";
              return (
                <div key={stageKey} className="record-detail-ai-result-card">
                  <div className="record-detail-ai-result-title">
                    <span>{label.shortTitle}</span>
                    <small>{label.title}</small>
                  </div>
                  <div className="record-detail-ai-result-body">
                    {renderAIResultContent(stageKey, content)}
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>
      <div className="detail-fields">
        {displayColumns.map((column) => renderDetailField(column, false))}
        {hiddenColumns.length > 0 ? (
          <div className="hidden-fields-section">
            <button
              type="button"
              className={`hidden-fields-toggle ${showHiddenFields ? "expanded" : ""}`}
              onClick={onToggleHiddenFields}
            >
              <IconChevron />
              <span>{hiddenColumns.length} 个已隐藏字段</span>
            </button>
            {showHiddenFields ? (
              <div className="hidden-fields-list">
                {hiddenColumns.map((column) => renderDetailField(column, true))}
              </div>
            ) : null}
          </div>
        ) : null}
      </div>
    </section>
  );
}
