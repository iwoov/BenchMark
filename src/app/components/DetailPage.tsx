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
    ParsedColumn,
    ParsedRow,
} from "../../types";
import {
    extractAIResultFinalAnswer,
    parseAIResultJSON,
    readBooleanLike,
} from "../ai-helpers";
import { normalizeHeaderTitle } from "../file-helpers";
import { IconChevron, IconEdit } from "../icons";
import {
    AI_CLEANING_TOOL_LABELS,
    AI_CLEANING_TOOL_ORDER,
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
    isAICleaning: boolean;
    activeAICleaningToolKey: AICleaningToolKey | null;
    aiCleaningElapsedText: string;
    aiCleaningStreamText: string;
    aiCleaningStatusMessage: string;
    onRemoveLevel3Tag?: (tag: string) => void;
    onUpdateBiochemLevel1Discipline?: (discipline: string) => Promise<void>;
    onRunAICleaning: (toolKey: AICleaningToolKey) => void;
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
    isAICleaning,
    activeAICleaningToolKey,
    aiCleaningElapsedText,
    aiCleaningStreamText,
    aiCleaningStatusMessage,
    onRemoveLevel3Tag,
    onUpdateBiochemLevel1Discipline,
    onRunAICleaning,
}: DetailPageProps) {
    const [detailImageZoom, setDetailImageZoom] = useState(100);
    const [isAIDetectSectionCollapsed, setIsAIDetectSectionCollapsed] =
        useState(false);
    const [aiDetectCardHeight, setAIDetectCardHeight] = useState(220);
    const aiDetectResizeRef = useRef<{
        active: boolean;
        startY: number;
        startHeight: number;
    }>({
        active: false,
        startY: 0,
        startHeight: 220,
    });

    if (!selectedRow) {
        return (
            <div className="record-list-empty">
                请先在题目列表页选择一条记录
            </div>
        );
    }

    const problemTextColumns = useMemo(
        () =>
            [...displayColumns]
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
            displayColumns.filter((column) =>
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
    }, [selectedRow.rowId]);

    useEffect(() => {
        const handleMouseMove = (event: MouseEvent) => {
            if (!aiDetectResizeRef.current.active) {
                return;
            }
            const deltaY = event.clientY - aiDetectResizeRef.current.startY;
            const nextHeight = Math.min(
                480,
                Math.max(140, aiDetectResizeRef.current.startHeight + deltaY),
            );
            setAIDetectCardHeight(nextHeight);
        };

        const handleMouseUp = () => {
            aiDetectResizeRef.current.active = false;
        };

        window.addEventListener("mousemove", handleMouseMove);
        window.addEventListener("mouseup", handleMouseUp);
        return () => {
            window.removeEventListener("mousemove", handleMouseMove);
            window.removeEventListener("mouseup", handleMouseUp);
        };
    }, []);

    const decreaseImageZoom = () => {
        setDetailImageZoom((previous) => Math.max(50, previous - 25));
    };

    const increaseImageZoom = () => {
        setDetailImageZoom((previous) => Math.min(250, previous + 25));
    };

    const startResizeAIDetectSection = (
        event: ReactMouseEvent<HTMLButtonElement>,
    ) => {
        aiDetectResizeRef.current = {
            active: true,
            startY: event.clientY,
            startHeight: aiDetectCardHeight,
        };
        event.preventDefault();
    };

    return (
        <section className="record-detail standalone-record-detail">
            <div className="record-detail-ai-toolbar">
                <div className="record-detail-ai-header">
                    <strong className="record-detail-ai-title">
                        AI自动化检测
                    </strong>
                    <div className="record-detail-ai-header-actions">
                        <button
                            type="button"
                            className="btn"
                            onClick={() =>
                                setIsAIDetectSectionCollapsed(
                                    (previous) => !previous,
                                )
                            }
                        >
                            {isAIDetectSectionCollapsed ? "展开" : "折叠"}
                        </button>
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
                </div>
                {!isAIDetectSectionCollapsed ? (
                    <div
                        className="record-detail-ai-results"
                        style={
                            {
                                "--ai-detect-card-height": `${aiDetectCardHeight}px`,
                            } as CSSProperties
                        }
                    >
                        <div className="record-detail-ai-results-grid">
                            {AI_STAGE_ORDER.map((stageKey) => {
                                const label = AI_STAGE_LABELS[stageKey];
                                const content = aiResults?.[stageKey] ?? "";
                                const stageTimer =
                                    runAllStageTimers?.[stageKey];
                                const hasResult = content.trim().length > 0;
                                const isRunAllRunning =
                                    Boolean(runAllTimerText);
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
                        <button
                            type="button"
                            className="record-detail-ai-resize-handle"
                            onMouseDown={startResizeAIDetectSection}
                            aria-label="拖动调整 AI 自动化检测高度"
                            title="拖动调整 AI 自动化检测高度"
                        >
                            <span />
                        </button>
                    </div>
                ) : null}
            </div>
            <div className="record-detail-ai-toolbar">
                <div className="record-detail-ai-header">
                    <strong className="record-detail-ai-title">数据清洗</strong>
                </div>
                <div className="record-detail-ai-results record-detail-ai-cleaning-results">
                    <div className="record-detail-ai-results-grid">
                        {AI_CLEANING_TOOL_ORDER.map((toolKey) => {
                            const label = AI_CLEANING_TOOL_LABELS[toolKey];
                            const savedContent =
                                cleaningResults?.[toolKey]?.responseText ?? "";
                            const isActiveTool =
                                activeAICleaningToolKey === toolKey;
                            const displayContent =
                                isActiveTool &&
                                isAICleaning &&
                                aiCleaningStreamText.trim().length > 0
                                    ? aiCleaningStreamText
                                    : savedContent;
                            return (
                                <div
                                    key={toolKey}
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
                                            onClick={() =>
                                                onRunAICleaning(toolKey)
                                            }
                                            disabled={isAICleaning}
                                        >
                                            {isActiveTool && isAICleaning
                                                ? aiCleaningElapsedText
                                                : "运行"}
                                        </button>
                                    </div>
                                    <div className="record-detail-ai-result-body">
                                        {toolKey === "biochem_level1_refine" ? (
                                            <BiochemLevel1Result
                                                rowId={selectedRow.rowId}
                                                content={displayContent}
                                                level1Options={level1Options}
                                                onSaveDiscipline={
                                                    onUpdateBiochemLevel1Discipline
                                                }
                                            />
                                        ) : (
                                            renderAICleaningResultContent(
                                                toolKey,
                                                displayContent,
                                                toolKey ===
                                                    "generate_level3_tags"
                                                    ? {
                                                          onRemoveLevel3Tag,
                                                      }
                                                    : undefined,
                                            )
                                        )}
                                    </div>
                                    {isActiveTool && aiCleaningStatusMessage ? (
                                        <div className="record-detail-ai-cleaning-message">
                                            {aiCleaningStatusMessage}
                                        </div>
                                    ) : null}
                                </div>
                            );
                        })}
                    </div>
                </div>
            </div>
            <div className="detail-page-layout">
                <div className="detail-page-main">
                    <div className="record-detail-header">
                        <h3>字段详情</h3>
                        <span>点击字段左侧勾选框可控制显示/隐藏</span>
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
            </div>
        </section>
    );
}
