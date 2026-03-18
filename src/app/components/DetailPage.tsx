import type { ReactNode } from "react";
import type { AIDetectStageKey, ParsedColumn, ParsedRow } from "../../types";
import {
    extractAIResultFinalAnswer,
    parseAIResultJSON,
    readBooleanLike,
} from "../ai-helpers";
import { IconChevron } from "../icons";
import {
    AI_RUN_ALL_LABEL,
    AI_STAGE_LABELS,
    AI_STAGE_ORDER,
} from "../constants";

function readTextValue(value: unknown): string {
    return typeof value === "string" ? value.trim() : "";
}

function hasMeaningfulText(value: string): boolean {
    return value.length > 0 && value !== "无";
}

function stringifyAnalysis(value: unknown): string {
    if (typeof value === "string") {
        return value.trim();
    }
    if (!value || typeof value !== "object") {
        return "";
    }
    return Object.entries(value)
        .map(([key, item]) =>
            typeof item === "string" && item.trim().length > 0
                ? `${key}: ${item.trim()}`
                : "",
        )
        .filter((item) => item.length > 0)
        .join("\n");
}

function renderInfoBlock(
    label: string,
    value: string,
    tone: "danger" | "warning" | "neutral" = "neutral",
) {
    if (!hasMeaningfulText(value)) {
        return null;
    }

    return (
        <div className={`ai-result-note note-${tone}`}>
            <strong>{label}：</strong>
            <span>{value}</span>
        </div>
    );
}

function getVerdictBadgeClass(verdict: string): string {
    if (verdict === "优秀") {
        return "badge-valid";
    }
    if (verdict === "需修改解析") {
        return "badge-warning";
    }
    return "badge-invalid";
}

function getRiskBadgeClass(level: string): string {
    if (level === "低") {
        return "badge-valid";
    }
    if (level === "中") {
        return "badge-warning";
    }
    return "badge-invalid";
}

function renderPrecheckResult(parsed: Record<string, unknown>) {
    const isValid = readBooleanLike(parsed.is_valid);
    const invalidReason =
        readTextValue(parsed.invalid_reason) || readTextValue(parsed.reason);
    const requiresImage = readBooleanLike(parsed.requires_image);
    const hasAbsoluteWords = readBooleanLike(parsed.has_absolute_words);
    const absoluteWordsDetails = readTextValue(parsed.absolute_words_details);
    const isAllSelectedWarning = readBooleanLike(
        parsed.is_all_selected_warning,
    );

    return (
        <div className="ai-result-formatted">
            <div className="ai-result-status-row">
                {isValid !== null && (
                    <span
                        className={`ai-result-badge ${isValid ? "badge-valid" : "badge-invalid"}`}
                    >
                        {isValid ? "✓ 题目有效" : "✗ 题目无效"}
                    </span>
                )}
                {requiresImage !== null && (
                    <span
                        className={`ai-result-badge ${requiresImage ? "badge-image-required" : "badge-image-optional"}`}
                    >
                        {requiresImage ? "需要图片" : "无需图片"}
                    </span>
                )}
                {hasAbsoluteWords !== null && (
                    <span
                        className={`ai-result-badge ${hasAbsoluteWords ? "badge-warning" : "badge-neutral"}`}
                    >
                        {hasAbsoluteWords ? "存在绝对化表述" : "无绝对化表述"}
                    </span>
                )}
                {isAllSelectedWarning !== null && (
                    <span
                        className={`ai-result-badge ${isAllSelectedWarning ? "badge-warning" : "badge-neutral"}`}
                    >
                        {isAllSelectedWarning ? "全选陷阱预警" : "非全选预警"}
                    </span>
                )}
            </div>
            {isValid === false
                ? renderInfoBlock("无效原因", invalidReason, "danger")
                : null}
            {hasAbsoluteWords === true
                ? renderInfoBlock(
                      "绝对化表述详情",
                      absoluteWordsDetails,
                      "warning",
                  )
                : null}
        </div>
    );
}

function renderLegacyContextAuditResult(parsed: Record<string, unknown>) {
    const isConsistent = readBooleanLike(parsed.is_consistent);
    const missingInfo = readTextValue(parsed.missing_info);

    return (
        <div className="ai-result-formatted">
            <div className="ai-result-status-row">
                {isConsistent !== null ? (
                    <span
                        className={`ai-result-badge ${isConsistent ? "badge-valid" : "badge-invalid"}`}
                    >
                        {isConsistent ? "✓ 一致" : "✗ 不一致"}
                    </span>
                ) : null}
            </div>
            {isConsistent === false
                ? renderInfoBlock("缺失/矛盾信息", missingInfo, "danger")
                : null}
        </div>
    );
}

function renderSubjectivityResult(parsed: Record<string, unknown>) {
    const isObjective = readBooleanLike(parsed.is_objective);
    const riskLevel = readTextValue(parsed.subjectivity_risk_level);
    const analysis = readTextValue(parsed.analysis);

    return (
        <div className="ai-result-formatted">
            <div className="ai-result-status-row">
                {isObjective !== null ? (
                    <span
                        className={`ai-result-badge ${isObjective ? "badge-valid" : "badge-warning"}`}
                    >
                        {isObjective ? "客观题" : "存在主观性风险"}
                    </span>
                ) : null}
                {hasMeaningfulText(riskLevel) ? (
                    <span
                        className={`ai-result-badge ${getRiskBadgeClass(riskLevel)}`}
                    >
                        {`风险：${riskLevel}`}
                    </span>
                ) : null}
            </div>
            {renderInfoBlock(
                "分析",
                analysis,
                isObjective === false ? "warning" : "neutral",
            )}
        </div>
    );
}

function renderIndependentSolvingResult(
    parsed: Record<string, unknown>,
    content: string,
) {
    const finalAnswer = extractAIResultFinalAnswer(content) ?? "无法确定";
    const canBeSolved = readBooleanLike(parsed.can_be_solved);
    const unsolvableReason = readTextValue(parsed.unsolvable_reason);
    const reasoning =
        readTextValue(parsed.ai_reasoning_step_by_step) ||
        stringifyAnalysis(parsed.analysis);

    return (
        <div className="ai-result-formatted">
            <div className="ai-result-status-row">
                <span className="ai-result-badge badge-answer">
                    {`答案：${finalAnswer}`}
                </span>
                {canBeSolved !== null ? (
                    <span
                        className={`ai-result-badge ${canBeSolved ? "badge-valid" : "badge-invalid"}`}
                    >
                        {canBeSolved ? "可解" : "不可解"}
                    </span>
                ) : null}
            </div>
            {renderInfoBlock(
                "推理过程",
                reasoning,
                canBeSolved === false ? "warning" : "neutral",
            )}
            {canBeSolved === false
                ? renderInfoBlock("缺失信息", unsolvableReason, "danger")
                : null}
        </div>
    );
}

function renderLegacyFinalVerdictResult(parsed: Record<string, unknown>) {
    const status = readTextValue(parsed.status);
    const isPass = status === "Pass";
    const discrepancy = readTextValue(parsed.discrepancy_detail);

    return (
        <div className="ai-result-formatted">
            <div className="ai-result-status-row">
                <span
                    className={`ai-result-badge ${isPass ? "badge-valid" : "badge-invalid"}`}
                >
                    {isPass ? "✓ Pass" : "✗ Fail"}
                </span>
            </div>
            {isPass ? null : renderInfoBlock("差异说明", discrepancy, "danger")}
        </div>
    );
}

function renderDeepAlignmentResult(parsed: Record<string, unknown>) {
    const isAnswerConsistent = readBooleanLike(parsed.is_answer_consistent);
    const hasExtraInfo = readBooleanLike(parsed.has_extra_info);
    const extraInfoDetails = readTextValue(parsed.extra_info_details);
    const isLogicForced = readBooleanLike(parsed.is_logic_forced);
    const logicFlawDetails = readTextValue(parsed.logic_flaw_details);
    const finalVerdict = readTextValue(parsed.final_verdict);

    return (
        <div className="ai-result-formatted">
            <div className="ai-result-status-row">
                {isAnswerConsistent !== null ? (
                    <span
                        className={`ai-result-badge ${isAnswerConsistent ? "badge-valid" : "badge-invalid"}`}
                    >
                        {isAnswerConsistent ? "答案一致" : "答案不一致"}
                    </span>
                ) : null}
                {hasExtraInfo !== null ? (
                    <span
                        className={`ai-result-badge ${hasExtraInfo ? "badge-warning" : "badge-valid"}`}
                    >
                        {hasExtraInfo ? "存在额外信息" : "无额外信息"}
                    </span>
                ) : null}
                {isLogicForced !== null ? (
                    <span
                        className={`ai-result-badge ${isLogicForced ? "badge-invalid" : "badge-valid"}`}
                    >
                        {isLogicForced ? "逻辑存疑" : "逻辑自洽"}
                    </span>
                ) : null}
                {hasMeaningfulText(finalVerdict) ? (
                    <span
                        className={`ai-result-badge ${getVerdictBadgeClass(finalVerdict)}`}
                    >
                        {finalVerdict}
                    </span>
                ) : null}
            </div>
            {hasExtraInfo === true
                ? renderInfoBlock("额外信息详情", extraInfoDetails, "warning")
                : null}
            {isLogicForced === true
                ? renderInfoBlock("逻辑问题", logicFlawDetails, "danger")
                : null}
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
            if (readBooleanLike(parsed.is_valid) !== null) {
                return renderPrecheckResult(parsed);
            }
            break;
        case "context_audit":
            if (
                readBooleanLike(parsed.is_objective) !== null ||
                readTextValue(parsed.subjectivity_risk_level).length > 0
            ) {
                return renderSubjectivityResult(parsed);
            }
            if (readBooleanLike(parsed.is_consistent) !== null) {
                return renderLegacyContextAuditResult(parsed);
            }
            break;
        case "independent_solving":
            if (
                extractAIResultFinalAnswer(trimmed) ||
                readTextValue(parsed.ai_reasoning_step_by_step).length > 0 ||
                readBooleanLike(parsed.can_be_solved) !== null ||
                readTextValue(stringifyAnalysis(parsed.analysis)).length > 0
            ) {
                return renderIndependentSolvingResult(parsed, trimmed);
            }
            break;
        case "final_verdict":
            if (
                readBooleanLike(parsed.is_answer_consistent) !== null ||
                readBooleanLike(parsed.has_extra_info) !== null ||
                readBooleanLike(parsed.is_logic_forced) !== null ||
                readTextValue(parsed.final_verdict).length > 0
            ) {
                return renderDeepAlignmentResult(parsed);
            }
            if (parsed.status === "Pass" || parsed.status === "Fail") {
                return renderLegacyFinalVerdictResult(parsed);
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
    onOpenAIRunModal: (stageKey: AIDetectStageKey) => void;
    onRunAllAIDetect: () => void;
    canRunAllAIDetect: boolean;
    runAllTimerText?: string;
    runAllStageTimers?: Partial<Record<AIDetectStageKey, string>>;
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
    onRunAllAIDetect,
    canRunAllAIDetect,
    runAllTimerText,
    runAllStageTimers,
    renderDetailField,
    aiResults,
}: DetailPageProps) {
    if (!selectedRow) {
        return (
            <div className="record-list-empty">
                请先在题目列表页选择一条记录
            </div>
        );
    }

    return (
        <section className="record-detail standalone-record-detail">
            <div className="record-detail-ai-toolbar">
                <div className="record-detail-ai-header">
                    <strong className="record-detail-ai-title">
                        AI自动化检测
                    </strong>
                    <button
                        type="button"
                        className="btn btn-primary"
                        onClick={onRunAllAIDetect}
                        disabled={!canRunAllAIDetect}
                    >
                        {runAllTimerText
                            ? `执行全部中 ${runAllTimerText}`
                            : AI_RUN_ALL_LABEL}
                    </button>
                </div>
                <div className="record-detail-ai-results">
                    <div className="record-detail-ai-results-grid">
                        {AI_STAGE_ORDER.map((stageKey) => {
                            const label = AI_STAGE_LABELS[stageKey];
                            const content = aiResults?.[stageKey] ?? "";
                            const stageTimer = runAllStageTimers?.[stageKey];
                            const hasResult = content.trim().length > 0;
                            const isRunAllRunning = Boolean(runAllTimerText);
                            const buttonLabel = hasResult
                                ? "查看"
                                : isRunAllRunning
                                  ? (stageTimer ?? "00:00")
                                  : "运行";
                            const buttonAriaLabel = hasResult
                                ? `查看 ${label.shortTitle}`
                                : isRunAllRunning
                                  ? `运行中 ${label.shortTitle}`
                                  : `运行 ${label.shortTitle}`;
                            const isButtonDisabled =
                                !hasResult && isRunAllRunning;
                            return (
                                <div
                                    key={stageKey}
                                    className="record-detail-ai-result-card"
                                >
                                    <div className="record-detail-ai-result-title">
                                        <div className="record-detail-ai-result-title-text">
                                            <span>{label.shortTitle}</span>
                                            <small>{label.title}</small>
                                        </div>
                                        <button
                                            type="button"
                                            className="btn btn-ghost ai-stage-run-btn"
                                            aria-label={buttonAriaLabel}
                                            onClick={() =>
                                                onOpenAIRunModal(stageKey)
                                            }
                                            disabled={isButtonDisabled}
                                        >
                                            {buttonLabel}
                                        </button>
                                    </div>
                                    <div className="record-detail-ai-result-body">
                                        {renderAIResultContent(
                                            stageKey,
                                            content,
                                        )}
                                    </div>
                                </div>
                            );
                        })}
                    </div>
                </div>
            </div>
            <div className="record-detail-header">
                <h3>字段详情</h3>
                <span>点击字段左侧勾选框可控制显示/隐藏</span>
            </div>
            <div className="detail-fields">
                {displayColumns.map((column) =>
                    renderDetailField(column, false),
                )}
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
                                {hiddenColumns.map((column) =>
                                    renderDetailField(column, true),
                                )}
                            </div>
                        ) : null}
                    </div>
                ) : null}
            </div>
        </section>
    );
}
