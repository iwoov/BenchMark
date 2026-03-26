import {
    useEffect,
    useMemo,
    useRef,
    useState,
    type CSSProperties,
    type MouseEvent as ReactMouseEvent,
    type ReactNode,
} from "react";
import type { AIDetectStageKey, ParsedColumn, ParsedRow } from "../../types";
import {
    extractAIResultFinalAnswer,
    parseAIResultJSON,
    readBooleanLike,
} from "../ai-helpers";
import { getCellImageSources, normalizeHeaderTitle } from "../file-helpers";
import { IconChevron, IconCopy } from "../icons";
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

async function normalizeImageBlobForClipboard(blob: Blob): Promise<Blob> {
    if (blob.type === "image/png") {
        return blob;
    }

    const objectUrl = URL.createObjectURL(blob);
    try {
        const image = await new Promise<HTMLImageElement>((resolve, reject) => {
            const nextImage = new Image();
            nextImage.onload = () => resolve(nextImage);
            nextImage.onerror = () => reject(new Error("image decode failed"));
            nextImage.src = objectUrl;
        });

        const canvas = document.createElement("canvas");
        canvas.width = image.naturalWidth || image.width;
        canvas.height = image.naturalHeight || image.height;
        const context = canvas.getContext("2d");
        if (!context) {
            throw new Error("canvas context unavailable");
        }

        context.drawImage(image, 0, 0);
        const pngBlob = await new Promise<Blob | null>((resolve) => {
            canvas.toBlob(resolve, "image/png");
        });
        if (!pngBlob) {
            throw new Error("png encode failed");
        }

        return pngBlob;
    } finally {
        URL.revokeObjectURL(objectUrl);
    }
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
    "reasoning",
    "analysis",
    "solution",
] as const;
const ANSWER_FIELD_TITLE_ALIASES = ["答案", "正确答案", "answer"] as const;

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
    const [detailImageZoom, setDetailImageZoom] = useState(100);
    const [isCopyingImage, setIsCopyingImage] = useState(false);
    const [copyImageFeedback, setCopyImageFeedback] = useState<{
        tone: "success" | "error";
        text: string;
    } | null>(null);

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
        if (sourceColumnsBelowHero.length > 0) {
            return baseRegularDisplayColumns;
        }

        const answerIndex = baseRegularDisplayColumns.findIndex((column) =>
            isAnswerColumn(column),
        );
        if (answerIndex < 0 || sourceColumns.length === 0) {
            return [...baseRegularDisplayColumns, ...sourceColumns];
        }

        const nextColumns = [...baseRegularDisplayColumns];
        nextColumns.splice(answerIndex + 1, 0, ...sourceColumns);
        return nextColumns;
    }, [
        baseRegularDisplayColumns,
        sourceColumns,
        sourceColumnsBelowHero.length,
    ]);
    const hasHeroLayout =
        problemTextColumns.length > 0 || imageColumns.length > 0;
    const firstImageSrc = useMemo(() => {
        for (const column of imageColumns) {
            const imageSources = getCellImageSources(
                selectedRow.values[column.key],
            );
            if (imageSources.length > 0) {
                return imageSources[0];
            }
        }
        return null;
    }, [imageColumns, selectedRow]);

    useEffect(() => {
        setDetailImageZoom(100);
        setIsCopyingImage(false);
        setCopyImageFeedback(null);
    }, [selectedRow.rowId]);

    useEffect(() => {
        if (!copyImageFeedback) {
            return;
        }

        const timeoutId = window.setTimeout(() => {
            setCopyImageFeedback(null);
        }, 2400);

        return () => window.clearTimeout(timeoutId);
    }, [copyImageFeedback]);

    const decreaseImageZoom = () => {
        setDetailImageZoom((previous) => Math.max(50, previous - 25));
    };

    const increaseImageZoom = () => {
        setDetailImageZoom((previous) => Math.min(250, previous + 25));
    };

    const copyFirstImage = async () => {
        if (!firstImageSrc || isCopyingImage) {
            return;
        }

        try {
            setIsCopyingImage(true);
            setCopyImageFeedback(null);

            const response = await fetch(firstImageSrc);
            if (!response.ok) {
                throw new Error(`copy image failed: ${response.status}`);
            }

            const imageBlob = await response.blob();
            if (
                !navigator.clipboard?.write ||
                typeof ClipboardItem === "undefined"
            ) {
                setCopyImageFeedback({
                    tone: "error",
                    text: window.isSecureContext
                        ? "当前浏览器不支持脚本复制图片，请右键图片复制"
                        : "当前页面不是安全上下文，请用 localhost/https 打开",
                });
                return;
            }

            const clipboardBlob =
                await normalizeImageBlobForClipboard(imageBlob);
            await navigator.clipboard.write([
                new ClipboardItem({
                    "image/png": clipboardBlob,
                }),
            ]);

            setCopyImageFeedback({
                tone: "success",
                text: "已复制第一张图片",
            });
        } catch (error) {
            console.error("[DetailImageCopy] failed", error);
            setCopyImageFeedback({
                tone: "error",
                text: "复制图片失败",
            });
        } finally {
            setIsCopyingImage(false);
        }
    };

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
                                    <div className="detail-problem-side-head">
                                        <div className="detail-problem-side-title">
                                            <strong>题目图片</strong>
                                            {copyImageFeedback ? (
                                                <span
                                                    className={`detail-copy-feedback tone-${copyImageFeedback.tone}`}
                                                    role="status"
                                                >
                                                    {copyImageFeedback.text}
                                                </span>
                                            ) : null}
                                        </div>
                                        <div className="detail-image-toolbar">
                                            <button
                                                type="button"
                                                className="btn btn-ghost detail-image-copy-btn"
                                                onClick={() =>
                                                    void copyFirstImage()
                                                }
                                                disabled={
                                                    !firstImageSrc ||
                                                    isCopyingImage
                                                }
                                                aria-label="复制第一张题目图片"
                                                title="复制第一张题目图片"
                                            >
                                                <IconCopy />
                                            </button>
                                            <div className="detail-image-zoom-controls">
                                                <button
                                                    type="button"
                                                    className="btn btn-ghost"
                                                    onClick={decreaseImageZoom}
                                                    disabled={
                                                        detailImageZoom <= 50
                                                    }
                                                >
                                                    -
                                                </button>
                                                <span>{`${detailImageZoom}%`}</span>
                                                <button
                                                    type="button"
                                                    className="btn btn-ghost"
                                                    onClick={() =>
                                                        setDetailImageZoom(100)
                                                    }
                                                >
                                                    100%
                                                </button>
                                                <button
                                                    type="button"
                                                    className="btn btn-ghost"
                                                    onClick={increaseImageZoom}
                                                    disabled={
                                                        detailImageZoom >= 250
                                                    }
                                                >
                                                    +
                                                </button>
                                            </div>
                                        </div>
                                    </div>
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
                                                canDrag={detailImageZoom > 100}
                                                zoomLevel={detailImageZoom}
                                                panelId={`${selectedRow.rowId}_${column.key}`}
                                            >
                                                {renderDetailField(
                                                    column,
                                                    false,
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
