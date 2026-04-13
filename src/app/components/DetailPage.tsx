import {
    useEffect,
    useMemo,
    useRef,
    useState,
    type ChangeEvent,
    type CSSProperties,
    type MouseEvent as ReactMouseEvent,
    type ReactNode,
} from "react";
import type {
    AICleaningToolKey,
    AICleaningToolResult,
    AIDetectStageKey,
    AIEvaluationAttemptResult,
    AIEvaluationTaskConfig,
    ParsedColumn,
    ParsedRow,
} from "../../types";
import {
    extractAIResultFinalAnswer,
    parseAIResultJSON,
    readBooleanLike,
} from "../ai-helpers";
import { normalizeHeaderTitle } from "../file-helpers";
import { IconChevron, IconEdit, IconPlus } from "../icons";
import {
    AI_CLEANING_TOOL_LABELS,
    AI_CLEANING_TOOL_ORDER,
    AI_RUN_ALL_LABEL,
    AI_STAGE_LABELS,
    AI_STAGE_ORDER,
} from "../constants";

type DetailWorkspacePanelKey = "quality" | "cleaning" | "evaluation";

const DETAIL_WORKSPACE_PANEL_STORAGE_KEY =
    "benchmark:detail-workspace-panel";
const DETAIL_CLEANING_TOOL_STORAGE_KEY = "benchmark:detail-cleaning-tool";

const DETAIL_WORKSPACE_PANELS: Array<{
    key: DetailWorkspacePanelKey;
    title: string;
    description: string;
}> = [
    {
        key: "quality",
        title: "数据质检",
        description: "检查题目质量、逻辑风险和答案一致性。",
    },
    {
        key: "cleaning",
        title: "数据清洗",
        description: "在清洗工具之间切换，查看或更新结构化结果。",
    },
    {
        key: "evaluation",
        title: "数据评测",
        description: "按评测任务切换查看与运行结果。",
    },
];

function isDetailWorkspacePanelKey(
    value: string,
): value is DetailWorkspacePanelKey {
    return DETAIL_WORKSPACE_PANELS.some((panel) => panel.key === value);
}

function isAICleaningToolKey(value: string): value is AICleaningToolKey {
    return AI_CLEANING_TOOL_ORDER.some((toolKey) => toolKey === value);
}

function readTextValue(value: unknown): string {
    return typeof value === "string" ? value.trim() : "";
}

function hasMeaningfulText(value: string): boolean {
    return value.length > 0 && value !== "无";
}

function formatJsonText(value: string | undefined, fallback: string): string {
    const source = value?.trim() || fallback.trim();
    if (!source) {
        return "";
    }
    try {
        return JSON.stringify(JSON.parse(source), null, 2);
    } catch {
        return source;
    }
}

function getEvaluationVerdictLabel(value: string): string {
    if (value === "correct") {
        return "正确";
    }
    if (value === "incorrect") {
        return "错误";
    }
    if (value === "undetermined") {
        return "无法判断";
    }
    return value || "-";
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

function renderTagLine(tags: string[], onRemoveTag?: (tag: string) => void) {
    if (tags.length === 0) {
        return null;
    }

    return (
        <div className="ai-inline-tag-row">
            <strong>标签：</strong>
            <div className="ai-inline-tag-list">
                {tags.map((tag) => (
                    <span key={tag} className="ai-inline-tag">
                        <span>{tag}</span>
                        {onRemoveTag ? (
                            <button
                                type="button"
                                className="ai-inline-tag-remove"
                                onClick={() => onRemoveTag(tag)}
                                aria-label={`删除标签 ${tag}`}
                                title={`删除标签 ${tag}`}
                            >
                                ×
                            </button>
                        ) : null}
                    </span>
                ))}
            </div>
        </div>
    );
}

function GenerateLevel3TagsResult({
    rowId,
    content,
    onRemoveTag,
    onAddTag,
}: {
    rowId: string;
    content: string;
    onRemoveTag?: (tag: string) => void;
    onAddTag?: (tag: string) => Promise<void>;
}) {
    const parsed = parseAIResultJSON(content);
    const representationMethod = readTextValue(parsed?.representation_method);
    const representationType = readTextValue(parsed?.representation_type);
    const parsedTags = Array.isArray(parsed?.tags)
        ? parsed.tags
              .filter((item): item is string => typeof item === "string")
              .map((item) => item.trim())
              .filter((item) => item.length > 0)
        : [];
    const [isAdding, setIsAdding] = useState(false);
    const [isSaving, setIsSaving] = useState(false);
    const [pendingTag, setPendingTag] = useState("");
    const [saveMessage, setSaveMessage] = useState("");
    const inputRef = useRef<HTMLInputElement>(null);

    useEffect(() => {
        setIsAdding(false);
        setIsSaving(false);
        setPendingTag("");
        setSaveMessage("");
    }, [content, rowId]);

    useEffect(() => {
        if (isAdding) {
            inputRef.current?.focus();
        }
    }, [isAdding]);

    const submitTag = async () => {
        const nextTag = pendingTag.trim();
        if (nextTag.length === 0) {
            return;
        }
        if (parsedTags.includes(nextTag)) {
            setSaveMessage("标签已存在");
            return;
        }
        if (!onAddTag) {
            return;
        }

        setIsSaving(true);
        setSaveMessage("");
        try {
            await onAddTag(nextTag);
            setPendingTag("");
            setIsAdding(false);
            setSaveMessage("已添加标签");
        } catch (error) {
            setSaveMessage(
                error instanceof Error ? error.message : "添加标签失败",
            );
        } finally {
            setIsSaving(false);
        }
    };

    return (
        <div className="ai-result-formatted">
            <div className="ai-result-status-row">
                {hasMeaningfulText(representationType) ? (
                    <span className="ai-result-badge badge-valid">
                        {representationType}
                    </span>
                ) : null}
            </div>
            <div className="ai-inline-tag-row">
                <strong>标签：</strong>
                <div className="ai-inline-tag-list">
                    {parsedTags.map((tag) => (
                        <span key={tag} className="ai-inline-tag">
                            <span>{tag}</span>
                            {onRemoveTag ? (
                                <button
                                    type="button"
                                    className="ai-inline-tag-remove"
                                    onClick={() => onRemoveTag(tag)}
                                    aria-label={`删除标签 ${tag}`}
                                    title={`删除标签 ${tag}`}
                                >
                                    ×
                                </button>
                            ) : null}
                        </span>
                    ))}
                    {isAdding ? (
                        <span className="ai-inline-tag ai-inline-tag-editor">
                            <input
                                ref={inputRef}
                                type="text"
                                value={pendingTag}
                                onChange={(event) =>
                                    setPendingTag(event.target.value)
                                }
                                onBlur={() => {
                                    setIsAdding(false);
                                    setPendingTag("");
                                    setSaveMessage("");
                                }}
                                onKeyDown={(event) => {
                                    if (event.key === "Enter") {
                                        event.preventDefault();
                                        void submitTag();
                                    }
                                    if (event.key === "Escape") {
                                        setIsAdding(false);
                                        setPendingTag("");
                                        setSaveMessage("");
                                    }
                                }}
                                placeholder="新标签"
                                disabled={isSaving}
                            />
                        </span>
                    ) : null}
                    <button
                        type="button"
                        className="btn btn-ghost ai-inline-tag-add"
                        onClick={() => {
                            setSaveMessage("");
                            setIsAdding((previous) => !previous);
                        }}
                        disabled={isSaving}
                        aria-label="添加标签"
                        title="添加标签"
                    >
                        <IconPlus />
                    </button>
                </div>
            </div>
            {saveMessage ? (
                <div className="record-detail-ai-cleaning-message">
                    {saveMessage}
                </div>
            ) : null}
            {renderInfoBlock("表征方法", representationMethod, "neutral")}
        </div>
    );
}

function BiochemLevel1Result({
    rowId,
    content,
    level1Options,
    onSaveDiscipline,
}: {
    rowId: string;
    content: string;
    level1Options: string[];
    onSaveDiscipline?: (discipline: string) => Promise<void>;
}) {
    const parsed = parseAIResultJSON(content);
    const discipline = readTextValue(parsed?.discipline);
    const confidence = readTextValue(parsed?.confidence);
    const reason = readTextValue(parsed?.reason);
    const [isEditing, setIsEditing] = useState(false);
    const [isSaving, setIsSaving] = useState(false);
    const [pendingDiscipline, setPendingDiscipline] = useState(discipline);
    const [saveMessage, setSaveMessage] = useState("");

    const selectableOptions = useMemo(() => {
        const items = [...level1Options];
        if (
            discipline.length > 0 &&
            !items.some((item) => item.trim() === discipline)
        ) {
            items.unshift(discipline);
        }
        return Array.from(
            new Set(
                items
                    .map((item) => item.trim())
                    .filter((item) => item.length > 0),
            ),
        );
    }, [discipline, level1Options]);

    useEffect(() => {
        setIsEditing(false);
        setIsSaving(false);
        setPendingDiscipline(discipline);
        setSaveMessage("");
    }, [discipline, rowId]);

    const handleDisciplineChange = async (
        event: ChangeEvent<HTMLSelectElement>,
    ) => {
        const nextValue = event.target.value.trim();
        setPendingDiscipline(nextValue);
        if (
            nextValue.length === 0 ||
            nextValue === discipline ||
            !onSaveDiscipline
        ) {
            setIsEditing(false);
            setSaveMessage("");
            return;
        }

        setIsSaving(true);
        setSaveMessage("");
        try {
            await onSaveDiscipline(nextValue);
            setIsEditing(false);
            setSaveMessage("已自动保存");
        } catch (error) {
            setSaveMessage(
                error instanceof Error ? error.message : "保存 discipline 失败",
            );
        } finally {
            setIsSaving(false);
        }
    };

    return (
        <div className="ai-result-formatted">
            <div className="ai-result-inline-edit-row">
                <div className="ai-result-inline-edit-label">discipline</div>
                <button
                    type="button"
                    className="btn btn-ghost ai-inline-edit-trigger"
                    onClick={() => {
                        setPendingDiscipline(discipline);
                        setSaveMessage("");
                        setIsEditing((previous) => !previous);
                    }}
                    disabled={isSaving || selectableOptions.length === 0}
                    aria-label="编辑生化 Level1 discipline"
                    title={
                        selectableOptions.length > 0
                            ? "编辑 discipline"
                            : "暂无可选的 level1"
                    }
                >
                    <IconEdit />
                </button>
            </div>
            <div className="ai-result-inline-edit-value">
                {isEditing ? (
                    <select
                        value={pendingDiscipline}
                        onChange={handleDisciplineChange}
                        disabled={isSaving || selectableOptions.length === 0}
                    >
                        <option value="">请选择 Level1</option>
                        {selectableOptions.map((item) => (
                            <option key={item} value={item}>
                                {item}
                            </option>
                        ))}
                    </select>
                ) : hasMeaningfulText(discipline) ? (
                    <span className="ai-result-badge badge-valid">
                        {discipline}
                    </span>
                ) : (
                    <span className="ai-result-empty">未设置</span>
                )}
            </div>
            <div className="ai-result-status-row">
                {hasMeaningfulText(confidence) ? (
                    <span
                        className={`ai-result-badge ${getRiskBadgeClass(confidence)}`}
                    >
                        {`置信度：${confidence}`}
                    </span>
                ) : null}
            </div>
            {saveMessage ? (
                <div className="record-detail-ai-cleaning-message">
                    {saveMessage}
                </div>
            ) : null}
            {renderInfoBlock("判断依据", reason, "neutral")}
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

function formatSuperiorAnswer(value: string): string {
    if (value === "expert") {
        return "专家更准确";
    }
    if (value === "ai") {
        return "AI更准确";
    }
    if (value === "tie") {
        return "双方各有问题";
    }
    return value;
}

function getSuperiorAnswerBadgeClass(value: string): string {
    if (value === "tie") {
        return "badge-warning";
    }
    return "badge-neutral";
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
    const superiorAnswer = readTextValue(parsed.superior_answer);
    const inconsistencyAnalysis = readTextValue(parsed.inconsistency_analysis);
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
                {isAnswerConsistent === false &&
                hasMeaningfulText(superiorAnswer) ? (
                    <span
                        className={`ai-result-badge ${getSuperiorAnswerBadgeClass(superiorAnswer)}`}
                    >
                        {formatSuperiorAnswer(superiorAnswer)}
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
            {isAnswerConsistent === false
                ? renderInfoBlock(
                      "不一致分析",
                      inconsistencyAnalysis,
                      "warning",
                  )
                : null}
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

function renderAICleaningResultContent(
    toolKey: AICleaningToolKey,
    content: string,
    options?: {
        onRemoveLevel3Tag?: (tag: string) => void;
    },
) {
    const trimmed = content.trim();
    if (trimmed.length === 0) {
        return <span className="ai-result-empty">暂无结果</span>;
    }

    const parsed = parseAIResultJSON(trimmed);
    if (!parsed) {
        return <pre className="ai-result-raw">{trimmed}</pre>;
    }

    if (toolKey === "generate_level3_tags") {
        const representationMethod = readTextValue(
            parsed.representation_method,
        );
        const representationType = readTextValue(parsed.representation_type);
        const parsedTags = Array.isArray(parsed.tags)
            ? parsed.tags
                  .filter((item): item is string => typeof item === "string")
                  .map((item) => item.trim())
                  .filter((item) => item.length > 0)
            : [];
        return (
            <div className="ai-result-formatted">
                <div className="ai-result-status-row">
                    {hasMeaningfulText(representationType) ? (
                        <span className="ai-result-badge badge-valid">
                            {representationType}
                        </span>
                    ) : null}
                </div>
                {renderTagLine(parsedTags, options?.onRemoveLevel3Tag)}
                {renderInfoBlock("表征方法", representationMethod, "neutral")}
            </div>
        );
    }

    if (toolKey === "biochem_level1_refine") {
        const discipline = readTextValue(parsed.discipline);
        const confidence = readTextValue(parsed.confidence);
        const reason = readTextValue(parsed.reason);
        return (
            <div className="ai-result-formatted">
                <div className="ai-result-status-row">
                    {hasMeaningfulText(discipline) ? (
                        <span className="ai-result-badge badge-valid">
                            {discipline}
                        </span>
                    ) : null}
                    {hasMeaningfulText(confidence) ? (
                        <span
                            className={`ai-result-badge ${getRiskBadgeClass(confidence)}`}
                        >
                            {`置信度：${confidence}`}
                        </span>
                    ) : null}
                </div>
                {renderInfoBlock("判断依据", reason, "neutral")}
            </div>
        );
    }

    return <pre className="ai-result-raw">{trimmed}</pre>;
}

const QUESTION_FIELD_TITLE_ALIASES = [
    "题目",
    "题干",
    "题目文本",
    "问题",
    "question",
] as const;
const OPTION_FIELD_TITLE_ALIASES = [
    "选项",
    "备选项",
    "options",
    "choices",
] as const;
const REASONING_FIELD_TITLE_ALIASES = [
    "解题过程",
    "解析",
    "解答过程",
    "答案解析",
    "原因",
    "判断依据",
    "reasoning",
    "analysis",
    "reason",
    "solution",
] as const;
const ANSWER_FIELD_TITLE_ALIASES = ["答案", "正确答案", "answer"] as const;
const LEVEL3_TAG_FIELD_TITLE_ALIASES = [
    "标签",
    "level3标签",
    "level3 标签",
    "tags",
] as const;

function matchesDetailFieldAlias(
    title: string,
    aliases: readonly string[],
): boolean {
    const normalizedTitle = normalizeHeaderTitle(title);
    return aliases.some((alias) =>
        normalizedTitle.includes(normalizeHeaderTitle(alias)),
    );
}

function matchesExactDetailFieldAlias(
    title: string,
    aliases: readonly string[],
): boolean {
    const normalizedTitle = normalizeHeaderTitle(title);
    return aliases.some(
        (alias) => normalizedTitle === normalizeHeaderTitle(alias),
    );
}

function isAnswerColumn(column: ParsedColumn): boolean {
    return matchesExactDetailFieldAlias(
        column.title,
        ANSWER_FIELD_TITLE_ALIASES,
    );
}

function isReasoningColumn(column: ParsedColumn): boolean {
    return matchesDetailFieldAlias(column.title, REASONING_FIELD_TITLE_ALIASES);
}

function isLevel3TagColumn(column: ParsedColumn): boolean {
    return matchesExactDetailFieldAlias(
        column.title,
        LEVEL3_TAG_FIELD_TITLE_ALIASES,
    );
}

function isSourceColumn(column: ParsedColumn): boolean {
    const normalizedTitle = normalizeHeaderTitle(column.title);
    return (
        normalizedTitle.includes(normalizeHeaderTitle("题目来源")) ||
        normalizedTitle === normalizeHeaderTitle("source")
    );
}

function getProblemTextColumnPriority(column: ParsedColumn): number {
    if (isSourceColumn(column)) {
        return Number.POSITIVE_INFINITY;
    }
    if (matchesDetailFieldAlias(column.title, QUESTION_FIELD_TITLE_ALIASES)) {
        return 0;
    }
    if (matchesDetailFieldAlias(column.title, OPTION_FIELD_TITLE_ALIASES)) {
        return 1;
    }
    if (isReasoningColumn(column)) {
        return 2;
    }
    return Number.POSITIVE_INFINITY;
}

function isProblemTextColumn(column: ParsedColumn): boolean {
    if (isSourceColumn(column)) {
        return false;
    }
    return (
        matchesDetailFieldAlias(column.title, QUESTION_FIELD_TITLE_ALIASES) ||
        matchesDetailFieldAlias(column.title, OPTION_FIELD_TITLE_ALIASES) ||
        isReasoningColumn(column)
    );
}

function isImageColumn(row: ParsedRow, column: ParsedColumn): boolean {
    const cell = row.values[column.key];
    return (
        cell?.type === "image" &&
        typeof cell.src === "string" &&
        cell.src.length > 0
    );
}

function DraggableImagePanel({
    children,
    canDrag,
    zoomLevel,
    panelId,
}: {
    children: ReactNode;
    canDrag: boolean;
    zoomLevel: number;
    panelId: string;
}) {
    const containerRef = useRef<HTMLDivElement | null>(null);
    const dragStateRef = useRef<{
        active: boolean;
        startX: number;
        startY: number;
        scrollLeft: number;
        scrollTop: number;
        target: HTMLElement | null;
        moved: boolean;
    }>({
        active: false,
        startX: 0,
        startY: 0,
        scrollLeft: 0,
        scrollTop: 0,
        target: null,
        moved: false,
    });
    const suppressClickRef = useRef(false);
    const [isDragging, setIsDragging] = useState(false);

    useEffect(() => {
        if (!canDrag) {
            dragStateRef.current.active = false;
            dragStateRef.current.target = null;
            setIsDragging(false);
        }
    }, [canDrag]);

    useEffect(() => {
        const detailValue = containerRef.current?.querySelector(
            ".detail-value",
        ) as HTMLElement | null;
        const imageCell = containerRef.current?.querySelector(
            ".image-cell",
        ) as HTMLElement | null;
        const image = containerRef.current?.querySelector(
            ".image-cell img",
        ) as HTMLImageElement | null;
        const caption = containerRef.current?.querySelector(
            ".image-cell span",
        ) as HTMLElement | null;
        if (!detailValue || !imageCell || !image) {
            return;
        }

        const applyImageSize = () => {
            if (!image.naturalWidth || !image.naturalHeight) {
                return;
            }

            const horizontalPadding = 24;
            const verticalPadding = 20;
            const captionHeight = caption ? caption.offsetHeight + 6 : 0;
            const availableWidth = Math.max(
                120,
                detailValue.clientWidth - horizontalPadding,
            );
            const availableHeight = Math.max(
                120,
                detailValue.clientHeight - verticalPadding - captionHeight,
            );
            const fitScale = Math.min(
                1,
                availableWidth / image.naturalWidth,
                availableHeight / image.naturalHeight,
            );
            const baseWidth = image.naturalWidth * fitScale;
            const baseHeight = image.naturalHeight * fitScale;
            const scaledWidth = baseWidth * (zoomLevel / 100);
            const scaledHeight = baseHeight * (zoomLevel / 100);

            image.style.width = `${scaledWidth}px`;
            image.style.height = `${scaledHeight}px`;
            imageCell.style.minWidth = `${scaledWidth}px`;
            imageCell.style.minHeight = `${scaledHeight + captionHeight}px`;

            if (!canDrag) {
                detailValue.scrollLeft = 0;
                detailValue.scrollTop = 0;
                return;
            }

            detailValue.scrollLeft = Math.max(
                0,
                (detailValue.scrollWidth - detailValue.clientWidth) / 2,
            );
            detailValue.scrollTop = Math.max(
                0,
                (detailValue.scrollHeight - detailValue.clientHeight) / 2,
            );
        };

        applyImageSize();
        image.addEventListener("load", applyImageSize);
        const resizeObserver = new ResizeObserver(() => {
            applyImageSize();
        });
        resizeObserver.observe(detailValue);

        return () => {
            image.removeEventListener("load", applyImageSize);
            resizeObserver.disconnect();
        };
    }, [canDrag, panelId, zoomLevel]);

    useEffect(() => {
        const handleMouseMove = (event: MouseEvent) => {
            if (!dragStateRef.current.active || !dragStateRef.current.target) {
                return;
            }
            const deltaX = event.clientX - dragStateRef.current.startX;
            const deltaY = event.clientY - dragStateRef.current.startY;
            if (Math.abs(deltaX) > 3 || Math.abs(deltaY) > 3) {
                dragStateRef.current.moved = true;
            }
            dragStateRef.current.target.scrollLeft =
                dragStateRef.current.scrollLeft - deltaX;
            dragStateRef.current.target.scrollTop =
                dragStateRef.current.scrollTop - deltaY;
        };

        const handleMouseUp = () => {
            if (!dragStateRef.current.active) {
                return;
            }
            suppressClickRef.current = dragStateRef.current.moved;
            dragStateRef.current.active = false;
            dragStateRef.current.target = null;
            dragStateRef.current.moved = false;
            setIsDragging(false);
        };

        window.addEventListener("mousemove", handleMouseMove);
        window.addEventListener("mouseup", handleMouseUp);
        return () => {
            window.removeEventListener("mousemove", handleMouseMove);
            window.removeEventListener("mouseup", handleMouseUp);
        };
    }, []);

    const handleMouseDown = (event: ReactMouseEvent<HTMLDivElement>) => {
        if (!canDrag || event.button !== 0) {
            return;
        }
        const detailValue = containerRef.current?.querySelector(
            ".detail-value",
        ) as HTMLElement | null;
        if (!detailValue) {
            return;
        }
        dragStateRef.current = {
            active: true,
            startX: event.clientX,
            startY: event.clientY,
            scrollLeft: detailValue.scrollLeft,
            scrollTop: detailValue.scrollTop,
            target: detailValue,
            moved: false,
        };
        setIsDragging(true);
        event.preventDefault();
    };

    const handleClickCapture = (event: ReactMouseEvent<HTMLDivElement>) => {
        if (!suppressClickRef.current) {
            return;
        }

        suppressClickRef.current = false;
        event.preventDefault();
        event.stopPropagation();
    };

    return (
        <div
            ref={containerRef}
            className={`detail-hero-image-panel ${canDrag ? "can-drag" : ""} ${isDragging ? "is-dragging" : ""}`}
            onMouseDown={handleMouseDown}
            onClickCapture={handleClickCapture}
        >
            {children}
        </div>
    );
}

interface DetailPageProps {
    selectedRow: ParsedRow | null;
    level1Options: string[];
    displayColumns: ParsedColumn[];
    hiddenColumns: ParsedColumn[];
    showHiddenFields: boolean;
    onToggleHiddenFields: () => void;
    onOpenAIRunModal: (stageKey: AIDetectStageKey) => void;
    onRunAllAIDetect: () => void;
    canRunAllAIDetect: boolean;
    runAllTimerText?: string;
    runAllStageTimers?: Partial<Record<AIDetectStageKey, string>>;
    renderDetailField: (
        column: ParsedColumn,
        isHidden: boolean,
        options?: { labelActions?: ReactNode },
    ) => ReactNode;
    aiResults?: Partial<Record<AIDetectStageKey, string>>;
    cleaningResults?: Partial<Record<AICleaningToolKey, AICleaningToolResult>>;
    evaluationTasks: AIEvaluationTaskConfig[];
    evaluationResults?: Record<string, AIEvaluationAttemptResult[]>;
    isAICleaning: boolean;
    activeAICleaningToolKey: AICleaningToolKey | null;
    aiCleaningElapsedText: string;
    aiCleaningStreamText: string;
    aiCleaningStatusMessage: string;
    isAIEvaluating: boolean;
    activeAIEvaluationTaskId: string | null;
    aiEvaluationElapsedText: string;
    aiEvaluationStatusMessage: string;
    onAddLevel3Tag?: (tag: string) => Promise<void>;
    onRemoveLevel3Tag?: (tag: string) => void;
    onUpdateBiochemLevel1Discipline?: (discipline: string) => Promise<void>;
    onRunAICleaning: (toolKey: AICleaningToolKey) => void;
    onRunAIEvaluation: (taskId: string) => void;
    onToggleRowEnabled: (rowId: string, enabled: boolean) => void;
}

export function DetailPage({
    selectedRow,
    level1Options,
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
    cleaningResults,
    evaluationTasks,
    evaluationResults,
    isAICleaning,
    activeAICleaningToolKey,
    aiCleaningElapsedText,
    aiCleaningStreamText,
    aiCleaningStatusMessage,
    isAIEvaluating,
    activeAIEvaluationTaskId,
    aiEvaluationElapsedText,
    aiEvaluationStatusMessage,
    onAddLevel3Tag,
    onRemoveLevel3Tag,
    onUpdateBiochemLevel1Discipline,
    onRunAICleaning,
    onRunAIEvaluation,
    onToggleRowEnabled,
}: DetailPageProps) {
    const [detailImageZoom, setDetailImageZoom] = useState(100);
    const [activeWorkspacePanel, setActiveWorkspacePanel] =
        useState<DetailWorkspacePanelKey>(() => {
            if (typeof window === "undefined") {
                return "cleaning";
            }
            const saved = window.localStorage.getItem(
                DETAIL_WORKSPACE_PANEL_STORAGE_KEY,
            );
            return saved && isDetailWorkspacePanelKey(saved)
                ? saved
                : "cleaning";
        });
    const [selectedCleaningToolKey, setSelectedCleaningToolKey] =
        useState<AICleaningToolKey>(() => {
            if (typeof window === "undefined") {
                return AI_CLEANING_TOOL_ORDER[0];
            }
            const saved = window.localStorage.getItem(
                DETAIL_CLEANING_TOOL_STORAGE_KEY,
            );
            return saved && isAICleaningToolKey(saved)
                ? saved
                : AI_CLEANING_TOOL_ORDER[0];
        });
    const [selectedEvaluationTaskId, setSelectedEvaluationTaskId] = useState(
        () => evaluationTasks[0]?.id ?? "",
    );
    const [selectedEvaluationRawAttempt, setSelectedEvaluationRawAttempt] =
        useState<AIEvaluationAttemptResult | null>(null);

    const problemTextColumns = useMemo(
        () =>
            !selectedRow
                ? []
                : [...displayColumns]
                .filter(
                    (column) =>
                        !isImageColumn(selectedRow, column) &&
                        isProblemTextColumn(column),
                )
                .sort(
                    (left, right) =>
                        getProblemTextColumnPriority(left) -
                        getProblemTextColumnPriority(right),
                ),
        [displayColumns, selectedRow],
    );
    const imageColumns = useMemo(
        () =>
            !selectedRow
                ? []
                : displayColumns.filter((column) =>
                      isImageColumn(selectedRow, column),
                  ),
        [displayColumns, selectedRow],
    );
    const sourceColumns = useMemo(
        () => displayColumns.filter((column) => isSourceColumn(column)),
        [displayColumns],
    );
    const heroColumnKeys = useMemo(
        () =>
            new Set(
                [...problemTextColumns, ...imageColumns].map(
                    (column) => column.key,
                ),
            ),
        [imageColumns, problemTextColumns],
    );
    const baseRegularDisplayColumns = useMemo(
        () =>
            displayColumns.filter(
                (column) =>
                    !heroColumnKeys.has(column.key) &&
                    !sourceColumns.some(
                        (sourceColumn) => sourceColumn.key === column.key,
                    ),
            ),
        [displayColumns, heroColumnKeys, sourceColumns],
    );
    const sourceColumnsBelowHero = useMemo(
        () =>
            problemTextColumns.some((column) => isAnswerColumn(column))
                ? sourceColumns
                : [],
        [problemTextColumns, sourceColumns],
    );
    const regularDisplayColumns = useMemo(() => {
        const orderedColumns = [...baseRegularDisplayColumns];
        const tagIndex = orderedColumns.findIndex((column) =>
            isLevel3TagColumn(column),
        );
        const reasoningIndex = orderedColumns.findIndex((column) =>
            isReasoningColumn(column),
        );
        if (tagIndex >= 0 && reasoningIndex >= 0 && tagIndex > reasoningIndex) {
            const [tagColumn] = orderedColumns.splice(tagIndex, 1);
            orderedColumns.splice(reasoningIndex, 0, tagColumn);
        }

        if (sourceColumnsBelowHero.length > 0) {
            return orderedColumns;
        }

        const answerIndex = orderedColumns.findIndex((column) =>
            isAnswerColumn(column),
        );
        if (answerIndex < 0 || sourceColumns.length === 0) {
            return [...orderedColumns, ...sourceColumns];
        }

        const nextColumns = [...orderedColumns];
        nextColumns.splice(answerIndex + 1, 0, ...sourceColumns);
        return nextColumns;
    }, [
        baseRegularDisplayColumns,
        sourceColumns,
        sourceColumnsBelowHero.length,
    ]);
    const hasHeroLayout =
        problemTextColumns.length > 0 || imageColumns.length > 0;

    useEffect(() => {
        setDetailImageZoom(100);
    }, [selectedRow?.rowId]);

    useEffect(() => {
        if (typeof window === "undefined") {
            return;
        }
        window.localStorage.setItem(
            DETAIL_WORKSPACE_PANEL_STORAGE_KEY,
            activeWorkspacePanel,
        );
    }, [activeWorkspacePanel]);

    useEffect(() => {
        if (typeof window === "undefined") {
            return;
        }
        window.localStorage.setItem(
            DETAIL_CLEANING_TOOL_STORAGE_KEY,
            selectedCleaningToolKey,
        );
    }, [selectedCleaningToolKey]);

    useEffect(() => {
        if (
            activeAICleaningToolKey &&
            activeWorkspacePanel === "cleaning" &&
            activeAICleaningToolKey !== selectedCleaningToolKey
        ) {
            setSelectedCleaningToolKey(activeAICleaningToolKey);
        }
    }, [
        activeAICleaningToolKey,
        activeWorkspacePanel,
        selectedCleaningToolKey,
    ]);

    useEffect(() => {
        if (
            activeAIEvaluationTaskId &&
            activeWorkspacePanel === "evaluation" &&
            activeAIEvaluationTaskId !== selectedEvaluationTaskId
        ) {
            setSelectedEvaluationTaskId(activeAIEvaluationTaskId);
            return;
        }
        if (
            selectedEvaluationTaskId &&
            evaluationTasks.some((task) => task.id === selectedEvaluationTaskId)
        ) {
            return;
        }
        setSelectedEvaluationTaskId(evaluationTasks[0]?.id ?? "");
    }, [
        activeAIEvaluationTaskId,
        activeWorkspacePanel,
        evaluationTasks,
        selectedEvaluationTaskId,
    ]);

    const decreaseImageZoom = () => {
        setDetailImageZoom((previous) => Math.max(50, previous - 25));
    };

    const increaseImageZoom = () => {
        setDetailImageZoom((previous) => Math.min(250, previous + 25));
    };

    const activeWorkspaceMeta =
        DETAIL_WORKSPACE_PANELS.find(
            (panel) => panel.key === activeWorkspacePanel,
        ) ?? DETAIL_WORKSPACE_PANELS[1];
    const cleaningToolLabel = AI_CLEANING_TOOL_LABELS[selectedCleaningToolKey];
    const cleaningSavedContent =
        cleaningResults?.[selectedCleaningToolKey]?.responseText ?? "";
    const isSelectedCleaningToolRunning =
        activeAICleaningToolKey === selectedCleaningToolKey && isAICleaning;
    const selectedCleaningContent =
        isSelectedCleaningToolRunning &&
        aiCleaningStreamText.trim().length > 0
            ? aiCleaningStreamText
            : cleaningSavedContent;
    const selectedEvaluationTask =
        evaluationTasks.find((task) => task.id === selectedEvaluationTaskId) ??
        evaluationTasks[0] ??
        null;
    const selectedEvaluationAttempts = selectedEvaluationTask
        ? (evaluationResults?.[selectedEvaluationTask.id] ?? [])
        : [];
    const isSelectedEvaluationTaskRunning =
        selectedEvaluationTask !== null &&
        activeAIEvaluationTaskId === selectedEvaluationTask.id &&
        isAIEvaluating;
    const evaluationAttemptMap = new Map(
        selectedEvaluationAttempts.map((attempt) => [
            attempt.attemptIndex,
            attempt,
        ]),
    );
    const evaluationAttemptCards = selectedEvaluationTask
        ? Array.from(
              { length: selectedEvaluationTask.attemptCount },
              (_, index) => index + 1,
          )
              .map((attemptIndex) => {
                  const attempt = evaluationAttemptMap.get(attemptIndex) ?? null;
                  if (!attempt) {
                      return {
                          attemptIndex,
                          status: isSelectedEvaluationTaskRunning
                              ? "pending"
                              : "empty",
                          finalAnswer: "-",
                          verdict: isSelectedEvaluationTaskRunning
                              ? "进行中"
                              : "未运行",
                          rawAttempt: null,
                      };
                  }
                  const generationParsed =
                      parseAIResultJSON(
                          attempt.generationParsedJsonText ?? "",
                      ) ?? parseAIResultJSON(attempt.generationResponseText);
                  const judgmentParsed =
                      parseAIResultJSON(
                          attempt.judgmentParsedJsonText ?? "",
                      ) ?? parseAIResultJSON(attempt.judgmentResponseText);
                  return {
                      attemptIndex,
                      status: "done",
                      finalAnswer:
                          readTextValue(generationParsed?.final_answer) ||
                          readTextValue(generationParsed?.ai_final_answer) ||
                          "-",
                      verdict:
                          readTextValue(judgmentParsed?.verdict) ||
                          attempt.finalVerdict ||
                          "-",
                      rawAttempt: attempt,
                  };
              })
              .sort((left, right) => right.attemptIndex - left.attemptIndex)
        : [];

    if (!selectedRow) {
        return (
            <div className="record-list-empty">
                请先在题目列表页选择一条记录
            </div>
        );
    }

    return (
        <>
        <section className="record-detail standalone-record-detail">
            <div className="record-detail-status-bar">
                <div className="record-detail-status-copy">
                    <strong>题目状态</strong>
                    <span>
                        {selectedRow.enabled
                            ? "当前题目已启用"
                            : "当前题目已停用"}
                    </span>
                </div>
                <label className="column-config-switch record-detail-enabled-switch">
                    <input
                        type="checkbox"
                        checked={selectedRow.enabled}
                        onChange={(event) =>
                            onToggleRowEnabled(
                                selectedRow.rowId,
                                event.target.checked,
                            )
                        }
                    />
                    <span>{selectedRow.enabled ? "启用" : "停用"}</span>
                </label>
            </div>
            <div className="detail-page-layout">
                <div className="detail-page-main">
                    <div className="record-detail-header">
                        <div>
                            <h3>原始题目</h3>
                            <span>题目内容始终固定展示，右侧工作区围绕它展开。</span>
                        </div>
                    </div>
                    <div className="detail-fields">
                        {hasHeroLayout ? (
                            <section className="detail-problem-layout">
                                <div className="detail-problem-main">
                                    <div
                                        className="detail-problem-main-grid"
                                        style={
                                            {
                                                "--detail-problem-row-count":
                                                    problemTextColumns.length,
                                            } as CSSProperties
                                        }
                                    >
                                        {problemTextColumns.map((column) =>
                                            renderDetailField(column, false),
                                        )}
                                    </div>
                                </div>
                                <div className="detail-problem-side">
                                    {imageColumns.length > 0 ? (
                                        <>
                                            <div
                                                className="detail-problem-image-panels"
                                                style={
                                                    {
                                                        "--detail-image-scale": `${detailImageZoom / 100}`,
                                                    } as CSSProperties
                                                }
                                            >
                                                {imageColumns.map((column) => (
                                                    <DraggableImagePanel
                                                        key={`${selectedRow.rowId}_${column.key}_hero`}
                                                        canDrag={
                                                            detailImageZoom >
                                                            100
                                                        }
                                                        zoomLevel={
                                                            detailImageZoom
                                                        }
                                                        panelId={`${selectedRow.rowId}_${column.key}`}
                                                    >
                                                        {renderDetailField(
                                                            column,
                                                            false,
                                                            imageColumns[0]
                                                                ?.key ===
                                                                column.key
                                                                ? {
                                                                      labelActions:
                                                                          (
                                                                              <div className="detail-image-toolbar inline-toolbar">
                                                                                  <div className="detail-image-zoom-controls">
                                                                                      <button
                                                                                          type="button"
                                                                                          className="btn btn-ghost"
                                                                                          onClick={
                                                                                              decreaseImageZoom
                                                                                          }
                                                                                          disabled={
                                                                                              detailImageZoom <=
                                                                                              50
                                                                                          }
                                                                                      >
                                                                                          -
                                                                                      </button>
                                                                                      <span>{`${detailImageZoom}%`}</span>
                                                                                      <button
                                                                                          type="button"
                                                                                          className="btn btn-ghost"
                                                                                          onClick={() =>
                                                                                              setDetailImageZoom(
                                                                                                  100,
                                                                                              )
                                                                                          }
                                                                                      >
                                                                                          100%
                                                                                      </button>
                                                                                      <button
                                                                                          type="button"
                                                                                          className="btn btn-ghost"
                                                                                          onClick={
                                                                                              increaseImageZoom
                                                                                          }
                                                                                          disabled={
                                                                                              detailImageZoom >=
                                                                                              250
                                                                                          }
                                                                                      >
                                                                                          +
                                                                                      </button>
                                                                                  </div>
                                                                              </div>
                                                                          ),
                                                                  }
                                                                : undefined,
                                                        )}
                                                    </DraggableImagePanel>
                                                ))}
                                            </div>
                                        </>
                                    ) : null}
                                </div>
                            </section>
                        ) : null}
                        {sourceColumnsBelowHero.map((column) =>
                            renderDetailField(column, false),
                        )}
                        {regularDisplayColumns.map((column) =>
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
                                    <span>
                                        {hiddenColumns.length} 个已隐藏字段
                                    </span>
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
                </div>
                <aside className="detail-page-workspace">
                    <div className="detail-workspace-shell">
                        <div className="detail-workspace-header">
                            <div className="detail-workspace-header-copy">
                                <span className="detail-workspace-eyebrow">
                                    处理工作区
                                </span>
                                <h3>{activeWorkspaceMeta.title}</h3>
                                <p>{activeWorkspaceMeta.description}</p>
                            </div>
                            {activeWorkspacePanel === "quality" ? (
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
                            ) : null}
                        </div>

                        <div className="detail-workspace-switcher">
                            {DETAIL_WORKSPACE_PANELS.map((panel) => (
                                <button
                                    key={panel.key}
                                    type="button"
                                    className={`detail-workspace-tab ${activeWorkspacePanel === panel.key ? "active" : ""}`}
                                    onClick={() =>
                                        setActiveWorkspacePanel(panel.key)
                                    }
                                >
                                    {panel.title}
                                </button>
                            ))}
                        </div>

                        <div className="detail-workspace-body">
                            {activeWorkspacePanel === "quality" ? (
                                <div className="detail-workspace-stage-list">
                                    {AI_STAGE_ORDER.map((stageKey) => {
                                        const label = AI_STAGE_LABELS[stageKey];
                                        const content =
                                            aiResults?.[stageKey] ?? "";
                                        const hasResult =
                                            content.trim().length > 0;
                                        const stageTimer =
                                            runAllStageTimers?.[stageKey];
                                        const isRunAllRunning =
                                            Boolean(runAllTimerText);
                                        const buttonLabel = hasResult
                                            ? "查看"
                                            : isRunAllRunning
                                              ? (stageTimer ?? "00:00")
                                              : "运行";
                                        const isButtonDisabled =
                                            !hasResult && isRunAllRunning;
                                        return (
                                            <section
                                                key={stageKey}
                                                className="detail-workspace-card"
                                            >
                                                <div className="detail-workspace-card-head">
                                                    <div className="detail-workspace-card-copy">
                                                        <strong>
                                                            {label.shortTitle}
                                                        </strong>
                                                        <span>{label.title}</span>
                                                    </div>
                                                    <button
                                                        type="button"
                                                        className="btn btn-ghost ai-stage-run-btn"
                                                        onClick={() =>
                                                            onOpenAIRunModal(
                                                                stageKey,
                                                            )
                                                        }
                                                        disabled={
                                                            isButtonDisabled
                                                        }
                                                    >
                                                        {buttonLabel}
                                                    </button>
                                                </div>
                                                <div className="detail-workspace-card-body">
                                                    {renderAIResultContent(
                                                        stageKey,
                                                        content,
                                                    )}
                                                </div>
                                            </section>
                                        );
                                    })}
                                </div>
                            ) : null}

                            {activeWorkspacePanel === "cleaning" ? (
                                <div className="detail-workspace-cleaning">
                                    <section className="detail-workspace-card detail-workspace-cleaning-card">
                                        <div className="detail-workspace-card-head detail-workspace-cleaning-head">
                                            <div className="detail-workspace-card-copy">
                                                <strong>
                                                    {cleaningToolLabel.shortTitle}
                                                </strong>
                                                <span>
                                                    {cleaningToolLabel.title}
                                                </span>
                                            </div>
                                            <div className="detail-workspace-cleaning-actions">
                                                <label className="detail-workspace-select">
                                                    <span>清洗工具</span>
                                                    <select
                                                        value={
                                                            selectedCleaningToolKey
                                                        }
                                                        onChange={(event) =>
                                                            setSelectedCleaningToolKey(
                                                                event.target
                                                                    .value as AICleaningToolKey,
                                                            )
                                                        }
                                                        disabled={isAICleaning}
                                                    >
                                                        {AI_CLEANING_TOOL_ORDER.map(
                                                            (toolKey) => (
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
                                                                            .title
                                                                    }
                                                                </option>
                                                            ),
                                                        )}
                                                    </select>
                                                </label>
                                                <button
                                                    type="button"
                                                    className="btn btn-primary"
                                                    onClick={() =>
                                                        onRunAICleaning(
                                                            selectedCleaningToolKey,
                                                        )
                                                    }
                                                    disabled={isAICleaning}
                                                >
                                                    {isSelectedCleaningToolRunning
                                                        ? aiCleaningElapsedText
                                                        : "运行工具"}
                                                </button>
                                            </div>
                                        </div>
                                        <div className="detail-workspace-card-body">
                                            {selectedCleaningToolKey ===
                                            "generate_level3_tags" ? (
                                                <GenerateLevel3TagsResult
                                                    rowId={selectedRow.rowId}
                                                    content={
                                                        selectedCleaningContent
                                                    }
                                                    onAddTag={onAddLevel3Tag}
                                                    onRemoveTag={
                                                        onRemoveLevel3Tag
                                                    }
                                                />
                                            ) : selectedCleaningToolKey ===
                                              "biochem_level1_refine" ? (
                                                <BiochemLevel1Result
                                                    rowId={selectedRow.rowId}
                                                    content={
                                                        selectedCleaningContent
                                                    }
                                                    level1Options={
                                                        level1Options
                                                    }
                                                    onSaveDiscipline={
                                                        onUpdateBiochemLevel1Discipline
                                                    }
                                                />
                                            ) : (
                                                renderAICleaningResultContent(
                                                    selectedCleaningToolKey,
                                                    selectedCleaningContent,
                                                )
                                            )}
                                        </div>
                                        {isSelectedCleaningToolRunning ||
                                        aiCleaningStatusMessage ? (
                                            <div className="detail-workspace-cleaning-footer">
                                                {isSelectedCleaningToolRunning ? (
                                                    <span className="detail-workspace-status-pill">
                                                        {`运行中 ${aiCleaningElapsedText}`}
                                                    </span>
                                                ) : null}
                                                {aiCleaningStatusMessage ? (
                                                    <div className="record-detail-ai-cleaning-message">
                                                        {
                                                            aiCleaningStatusMessage
                                                        }
                                                    </div>
                                                ) : null}
                                            </div>
                                        ) : null}
                                    </section>
                                </div>
                            ) : null}

                            {activeWorkspacePanel === "evaluation" ? (
                                <section className="detail-workspace-card">
                                    <div className="detail-workspace-card-head detail-workspace-cleaning-head">
                                        <div className="detail-workspace-card-copy">
                                            <strong>数据评测</strong>
                                            <span>按评测任务切换查看与运行结果</span>
                                        </div>
                                        <div className="detail-workspace-cleaning-actions">
                                            <label className="detail-workspace-select">
                                                <span>评测任务</span>
                                                <select
                                                    value={
                                                        selectedEvaluationTask?.id ??
                                                        ""
                                                    }
                                                    onChange={(event) =>
                                                        setSelectedEvaluationTaskId(
                                                            event.target.value,
                                                        )
                                                    }
                                                >
                                                    {evaluationTasks.map((task) => (
                                                        <option
                                                            key={task.id}
                                                            value={task.id}
                                                        >
                                                            {task.name}
                                                        </option>
                                                    ))}
                                                </select>
                                            </label>
                                            <button
                                                type="button"
                                                className="btn btn-primary"
                                                onClick={() =>
                                                    selectedEvaluationTask &&
                                                    onRunAIEvaluation(
                                                        selectedEvaluationTask.id,
                                                    )
                                                }
                                                disabled={
                                                    !selectedEvaluationTask ||
                                                    isAIEvaluating
                                                }
                                            >
                                                {isSelectedEvaluationTaskRunning
                                                    ? aiEvaluationElapsedText
                                                    : "运行评测"}
                                            </button>
                                        </div>
                                    </div>
                                    <div className="detail-workspace-card-body detail-workspace-cleaning-layout">
                                        {selectedEvaluationTask ? (
                                            <>
                                                <div className="detail-workspace-cleaning-tools">
                                                    <div className="detail-workspace-cleaning-tool-summary">
                                                        <strong>
                                                            {
                                                                selectedEvaluationTask.name
                                                            }
                                                        </strong>
                                                        <span>{`评测次数 ${selectedEvaluationTask.attemptCount} 次`}</span>
                                                        <span>
                                                            {selectedEvaluationTask.enabled
                                                                ? "已启用"
                                                                : "未启用"}
                                                        </span>
                                                    </div>
                                                    <div className="detail-workspace-cleaning-tool-summary">
                                                        <strong>
                                                            第一步：题目作答
                                                        </strong>
                                                        <span>
                                                            {
                                                                selectedEvaluationTask
                                                                    .answerGeneration
                                                                    .routeName
                                                            }
                                                        </span>
                                                    </div>
                                                    <div className="detail-workspace-cleaning-tool-summary">
                                                        <strong>
                                                            第二步：答案判定
                                                        </strong>
                                                        <span>
                                                            {
                                                                selectedEvaluationTask
                                                                    .answerJudgment
                                                                    .routeName
                                                            }
                                                        </span>
                                                    </div>
                                                    <div className="detail-workspace-cleaning-tool-summary">
                                                        <strong>已保存尝试</strong>
                                                        <span>{`${selectedEvaluationAttempts.length} 次`}</span>
                                                    </div>
                                                </div>
                                                <div className="detail-workspace-cleaning-content">
                                                    {evaluationAttemptCards.length > 0 ? (
                                                        <div className="detail-evaluation-attempt-list">
                                                            {evaluationAttemptCards.map(
                                                                ({
                                                                    attemptIndex,
                                                                    finalAnswer,
                                                                    verdict,
                                                                    rawAttempt,
                                                                    status,
                                                                }) => (
                                                                    <article
                                                                        key={
                                                                            attemptIndex
                                                                        }
                                                                        className="detail-evaluation-attempt-card"
                                                                    >
                                                                        <div className="detail-evaluation-attempt-head">
                                                                            <strong>{`第 ${attemptIndex} 次`}</strong>
                                                                            {rawAttempt ? (
                                                                                <button
                                                                                    type="button"
                                                                                    className="btn btn-ghost"
                                                                                    onClick={() =>
                                                                                        setSelectedEvaluationRawAttempt(
                                                                                            rawAttempt,
                                                                                        )
                                                                                    }
                                                                                >
                                                                                    查看 JSON
                                                                                </button>
                                                                            ) : null}
                                                                        </div>
                                                                        <div className="detail-evaluation-attempt-grid">
                                                                            <div className="detail-evaluation-attempt-item">
                                                                                <span>
                                                                                    最终答案
                                                                                </span>
                                                                                <strong>
                                                                                    {
                                                                                        finalAnswer
                                                                                    }
                                                                                </strong>
                                                                            </div>
                                                                            <div className="detail-evaluation-attempt-item">
                                                                                <span>
                                                                                    判定结果
                                                                                </span>
                                                                                <strong>
                                                                                    {
                                                                                        status ===
                                                                                        "done"
                                                                                            ? getEvaluationVerdictLabel(
                                                                                                  verdict,
                                                                                              )
                                                                                            : verdict
                                                                                    }
                                                                                </strong>
                                                                            </div>
                                                                        </div>
                                                                    </article>
                                                                ),
                                                            )}
                                                        </div>
                                                    ) : (
                                                        <div className="detail-workspace-placeholder">
                                                            <p>
                                                                当前任务还没有评测结果。
                                                            </p>
                                                            <p>
                                                                选择任务后点击“运行评测”即可写入数据库并在这里查看多次结果。
                                                            </p>
                                                        </div>
                                                    )}
                                                </div>
                                            </>
                                        ) : (
                                            <div className="detail-workspace-placeholder">
                                                <p>当前还没有配置任何评测任务。</p>
                                            </div>
                                        )}
                                    </div>
                                    {isSelectedEvaluationTaskRunning ||
                                    aiEvaluationStatusMessage ? (
                                        <div className="detail-workspace-cleaning-footer">
                                            {isSelectedEvaluationTaskRunning ? (
                                                <span className="detail-workspace-status-pill">
                                                    {`运行中 ${aiEvaluationElapsedText}`}
                                                </span>
                                            ) : null}
                                            {aiEvaluationStatusMessage ? (
                                                <div className="record-detail-ai-cleaning-message">
                                                    {
                                                        aiEvaluationStatusMessage
                                                    }
                                                </div>
                                            ) : null}
                                        </div>
                                    ) : null}
                                </section>
                            ) : null}
                        </div>
                    </div>
                </aside>
            </div>
        </section>
        {selectedEvaluationRawAttempt ? (
            <div className="column-modal-mask">
                <div className="column-modal ai-config-modal">
                    <h3>{`第 ${selectedEvaluationRawAttempt.attemptIndex} 次评测原始 JSON`}</h3>
                    <p>
                        {selectedEvaluationTask?.name ?? "评测任务"}
                    </p>
                    <div className="ai-config-form">
                        <div className="ai-config-section">
                            <div className="ai-config-section-title">
                                第一步：题目作答 JSON
                            </div>
                            <pre className="detail-evaluation-json">
                                {formatJsonText(
                                    selectedEvaluationRawAttempt.generationParsedJsonText,
                                    selectedEvaluationRawAttempt.generationResponseText,
                                )}
                            </pre>
                        </div>
                        <div className="ai-config-section">
                            <div className="ai-config-section-title">
                                第二步：答案判定 JSON
                            </div>
                            <pre className="detail-evaluation-json">
                                {formatJsonText(
                                    selectedEvaluationRawAttempt.judgmentParsedJsonText,
                                    selectedEvaluationRawAttempt.judgmentResponseText,
                                )}
                            </pre>
                        </div>
                    </div>
                    <div className="column-modal-footer">
                        <button
                            type="button"
                            className="btn btn-primary"
                            onClick={() => setSelectedEvaluationRawAttempt(null)}
                        >
                            关闭
                        </button>
                    </div>
                </div>
            </div>
        ) : null}
        </>
    );
}
