import { useEffect, type ReactNode, useState } from "react";
import type { ParsedCell, ParsedColumn, ParsedRow } from "../../types";
import {
    getCellText,
    isFeedbackColumnTitle,
    isInspectorColumnTitle,
    isOpensourceColumnTitle,
    isQualifiedColumnTitle,
    logUIImageRenderError,
    normalizeHeaderTitle,
} from "../file-helpers";
import { IconCopy } from "../icons";
import {
    LatexRenderer,
    hasLatexSyntax,
    shouldAutoDisplayLatex,
} from "../latex";

const QUESTION_COPY_FIELD_ALIASES = [
    "题目",
    "题干",
    "题目文本",
    "问题",
    "question",
] as const;
const OPTION_COPY_FIELD_ALIASES = [
    "选项",
    "备选项",
    "options",
    "choices",
] as const;
const REASONING_COPY_FIELD_ALIASES = [
    "解题过程",
    "解析",
    "解答过程",
    "答案解析",
    "reasoning",
    "analysis",
    "solution",
] as const;
const LEVEL3_TAG_FIELD_ALIASES = [
    "标签",
    "level3标签",
    "level3 标签",
    "tags",
] as const;

function matchesCopyFieldAlias(
    title: string,
    aliases: readonly string[],
): boolean {
    const normalizedTitle = normalizeHeaderTitle(title);
    return aliases.some((alias) =>
        normalizedTitle.includes(normalizeHeaderTitle(alias)),
    );
}

function isCopyableProblemTextColumn(column: ParsedColumn): boolean {
    return (
        matchesCopyFieldAlias(column.title, QUESTION_COPY_FIELD_ALIASES) ||
        matchesCopyFieldAlias(column.title, OPTION_COPY_FIELD_ALIASES) ||
        matchesCopyFieldAlias(column.title, REASONING_COPY_FIELD_ALIASES)
    );
}

function isLevel3TagFieldTitle(column: ParsedColumn): boolean {
    const normalizedTitle = normalizeHeaderTitle(column.title);
    return LEVEL3_TAG_FIELD_ALIASES.some(
        (alias) => normalizedTitle === normalizeHeaderTitle(alias),
    );
}

function getExternalUrl(value: string): string | null {
    const trimmed = value.trim();
    if (!/^https?:\/\//i.test(trimmed)) {
        return null;
    }

    try {
        return new URL(trimmed).toString();
    } catch {
        return null;
    }
}

function splitTagValues(value: string): string[] {
    return value
        .split(/,\s*|\n+|，|；|;|\||[ \u3000]{2,}/)
        .map((item) => item.trim())
        .filter((item) => item.length > 0);
}

export const useCellRenderers = ({
    selectedRow,
    level3TagsFieldKey,
    latexRenderOverrides,
    onToggleLatexRender,
    onToggleDisplayColumn,
    onEditCell,
    getLatexToggleKey,
    setPreviewImageSrc,
}: {
    selectedRow: ParsedRow | null;
    level3TagsFieldKey?: string;
    latexRenderOverrides: Record<string, boolean>;
    onToggleLatexRender: (columnKey: string) => void;
    onToggleDisplayColumn: (columnKey: string) => void;
    onEditCell: (rowId: string, columnKey: string, value: string) => void;
    getLatexToggleKey: (columnKey: string) => string;
    setPreviewImageSrc: (value: string | null) => void;
}) => {
    const [copiedFieldKey, setCopiedFieldKey] = useState<string | null>(null);

    useEffect(() => {
        if (!copiedFieldKey) {
            return;
        }

        const timeoutId = window.setTimeout(() => {
            setCopiedFieldKey(null);
        }, 1800);

        return () => window.clearTimeout(timeoutId);
    }, [copiedFieldKey]);

    const renderTextValue = (
        textValue: string,
        shouldRenderLatex: boolean,
        variant: "list" | "detail",
    ) => {
        const hasLatex = hasLatexSyntax(textValue);
        const autoDisplayLatex = shouldAutoDisplayLatex(textValue);
        if (hasLatex && shouldRenderLatex) {
            return (
                <LatexRenderer
                    value={textValue}
                    forceDisplay={autoDisplayLatex}
                />
            );
        }

        const externalUrl = getExternalUrl(textValue);
        if (externalUrl) {
            return (
                <a
                    className={`cell-link cell-link-${variant}`}
                    href={externalUrl}
                    target="_blank"
                    rel="noreferrer noopener"
                    onClick={(event) => event.stopPropagation()}
                    title={externalUrl}
                >
                    {textValue}
                </a>
            );
        }

        return hasLatex ? (
            <div className="latex-plain">{textValue}</div>
        ) : (
            <div className="plain-text-value">{textValue}</div>
        );
    };

    const renderReadonlyCell = (
        row: ParsedRow,
        column: ParsedColumn,
        cell: ParsedCell | undefined,
        shouldRenderLatex: boolean,
        variant: "list" | "detail" = "detail",
    ) => {
        if (!cell) {
            return <span className="empty-text">-</span>;
        }

        if (cell.type === "image" && cell.src) {
            return (
                <div className="image-cell">
                    <img
                        src={cell.src}
                        alt={cell.value || "Excel图片"}
                        onClick={() => setPreviewImageSrc(cell.src!)}
                        onError={() => {
                            logUIImageRenderError(
                                row.rowId,
                                column.title,
                                cell.src ?? "",
                            );
                        }}
                    />
                    {cell.value ? <span>{cell.value}</span> : null}
                </div>
            );
        }

        const textValue = cell.value ?? "";
        if (cell.type === "text" && textValue.length > 0) {
            return renderTextValue(textValue, shouldRenderLatex, variant);
        }

        return cell.value ? (
            <div className="plain-text-value">{cell.value}</div>
        ) : (
            <span className="empty-text">-</span>
        );
    };

    const renderCellContent = (
        row: ParsedRow,
        column: ParsedColumn,
        shouldRenderLatex = true,
    ) => {
        const cell = row.values[column.key];
        if (!column.editable) {
            return renderReadonlyCell(row, column, cell, shouldRenderLatex);
        }

        const currentValue = cell?.value ?? "";

        if (isQualifiedColumnTitle(column.title)) {
            const stableOptions = ["", "合格", "不合格"];
            const shouldAppendCurrent =
                currentValue.length > 0 &&
                !stableOptions.includes(currentValue);
            return (
                <select
                    className="qualified-select"
                    value={currentValue}
                    onChange={(event) =>
                        onEditCell(row.rowId, column.key, event.target.value)
                    }
                >
                    <option value="">未填写</option>
                    <option value="合格">合格</option>
                    <option value="不合格">不合格</option>
                    {shouldAppendCurrent ? (
                        <option value={currentValue}>{currentValue}</option>
                    ) : null}
                </select>
            );
        }

        if (isOpensourceColumnTitle(column.title)) {
            const stableOptions = ["", "是", "否"];
            const shouldAppendCurrent =
                currentValue.length > 0 &&
                !stableOptions.includes(currentValue);
            return (
                <select
                    className="qualified-select"
                    value={currentValue}
                    onChange={(event) =>
                        onEditCell(row.rowId, column.key, event.target.value)
                    }
                >
                    <option value="">未填写</option>
                    <option value="是">是</option>
                    <option value="否">否</option>
                    {shouldAppendCurrent ? (
                        <option value={currentValue}>{currentValue}</option>
                    ) : null}
                </select>
            );
        }

        if (isInspectorColumnTitle(column.title)) {
            return (
                <input
                    className="inspector-input"
                    value={currentValue}
                    onChange={(event) =>
                        onEditCell(row.rowId, column.key, event.target.value)
                    }
                    placeholder="请输入质检员"
                />
            );
        }

        if (isFeedbackColumnTitle(column.title)) {
            return (
                <textarea
                    className="feedback-input"
                    value={currentValue}
                    onChange={(event) =>
                        onEditCell(row.rowId, column.key, event.target.value)
                    }
                    placeholder="请输入质检反馈意见"
                />
            );
        }

        if (isCopyableProblemTextColumn(column)) {
            return (
                <textarea
                    className="editable-textarea-input"
                    value={currentValue}
                    onChange={(event) =>
                        onEditCell(row.rowId, column.key, event.target.value)
                    }
                    placeholder={`请输入${column.title}`}
                    rows={4}
                />
            );
        }

        if (cell?.type === "image" && cell.src) {
            return (
                <div className="image-cell">
                    <img
                        src={cell.src}
                        alt={cell.value || "Excel图片"}
                        onClick={() => setPreviewImageSrc(cell.src!)}
                        onError={() => {
                            logUIImageRenderError(
                                row.rowId,
                                column.title,
                                cell.src ?? "",
                            );
                        }}
                    />
                    <input
                        className="editable-text-input"
                        value={currentValue}
                        onChange={(event) =>
                            onEditCell(
                                row.rowId,
                                column.key,
                                event.target.value,
                            )
                        }
                        placeholder={`请输入${column.title}`}
                    />
                </div>
            );
        }

        return (
            <input
                className="editable-text-input"
                value={currentValue}
                onChange={(event) =>
                    onEditCell(row.rowId, column.key, event.target.value)
                }
                placeholder={`请输入${column.title}`}
            />
        );
    };

    const renderDetailField = (
        column: ParsedColumn,
        isHidden = false,
        options?: {
            labelActions?: ReactNode;
            valueContent?: ReactNode;
        },
    ) => {
        if (!selectedRow) {
            return null;
        }
        const isRequired = column.editable;
        const isChecked = !isHidden;
        const cell = selectedRow.values[column.key];
        const hasLatex =
            !column.editable &&
            !isHidden &&
            cell?.type === "text" &&
            typeof cell.value === "string" &&
            hasLatexSyntax(cell.value);
        const latexToggleKey = getLatexToggleKey(column.key);
        const isLatexRenderingEnabled =
            latexRenderOverrides[latexToggleKey] ?? false;
        const copyText = typeof cell?.value === "string" ? cell.value : "";
        const canCopyFieldText =
            !isHidden &&
            isCopyableProblemTextColumn(column) &&
            copyText.trim().length > 0;
        const isFieldCopied = copiedFieldKey === column.key;
        const isLevel3TagsField =
            !isHidden &&
            ((typeof level3TagsFieldKey === "string" &&
                level3TagsFieldKey.length > 0 &&
                column.key === level3TagsFieldKey) ||
                isLevel3TagFieldTitle(column));
        const tags = isLevel3TagsField ? splitTagValues(copyText) : [];

        const handleCopyFieldText = async () => {
            if (!canCopyFieldText || !navigator.clipboard?.writeText) {
                return;
            }

            try {
                await navigator.clipboard.writeText(copyText);
                setCopiedFieldKey(column.key);
            } catch (error) {
                console.error("[DetailFieldCopy] failed", error);
            }
        };

        const handleRemoveTag = (tag: string) => {
            if (!isLevel3TagsField) {
                return;
            }
            const nextTags = tags.filter((item) => item !== tag);
            onEditCell(selectedRow.rowId, column.key, nextTags.join(", "));
        };

        return (
            <div
                key={`${selectedRow.rowId}_${column.key}`}
                className={`detail-field ${isHidden ? "hidden-field" : ""}`}
            >
                <div className="detail-label">
                    <button
                        type="button"
                        className={`field-toggle ${isRequired ? "locked" : ""} ${isChecked ? "checked" : ""}`}
                        onClick={() => {
                            if (!isRequired) {
                                onToggleDisplayColumn(column.key);
                            }
                        }}
                        title={
                            isRequired
                                ? "可编辑字段必须展示"
                                : isHidden
                                  ? "点击显示此字段"
                                  : "点击隐藏此字段"
                        }
                    />
                    <div className="field-name-wrap">
                        <span className="field-name">{column.title}</span>
                        {hasLatex ? (
                            <label
                                className="latex-toggle"
                                title="控制该字段是否按 LaTeX 公式渲染"
                            >
                                <input
                                    type="checkbox"
                                    checked={isLatexRenderingEnabled}
                                    onChange={() =>
                                        onToggleLatexRender(column.key)
                                    }
                                    aria-label={`${column.title} 的 LaTeX 渲染开关`}
                                />
                                <span>LaTeX渲染</span>
                            </label>
                        ) : null}
                    </div>
                    {canCopyFieldText ? (
                        <div className="detail-label-actions">
                            {isFieldCopied ? (
                                <span
                                    className="field-copy-feedback"
                                    role="status"
                                >
                                    已复制
                                </span>
                            ) : null}
                            <button
                                type="button"
                                className="btn btn-ghost field-copy-btn"
                                onClick={() => void handleCopyFieldText()}
                                aria-label={`复制${column.title}`}
                                title={`复制${column.title}`}
                            >
                                <IconCopy />
                            </button>
                        </div>
                    ) : null}
                    {options?.labelActions ? (
                        <div className="detail-label-actions">
                            {options.labelActions}
                        </div>
                    ) : null}
                    {column.editable ? (
                        <span className="field-badge badge-editable">
                            可编辑
                        </span>
                    ) : null}
                    {isRequired ? (
                        <span className="field-badge badge-locked">必显</span>
                    ) : null}
                </div>
                {!isHidden ? (
                    <div className="detail-value">
                        {options?.valueContent ??
                            (isLevel3TagsField ? (
                                tags.length > 0 ? (
                                    <div className="detail-tag-inline">
                                        <strong>标签：</strong>
                                        <div className="detail-tag-inline-list">
                                            {tags.map((tag) => (
                                                <span
                                                    key={tag}
                                                    className="detail-tag-chip"
                                                >
                                                    <span>{tag}</span>
                                                    <button
                                                        type="button"
                                                        className="detail-tag-delete"
                                                        onClick={() =>
                                                            handleRemoveTag(tag)
                                                        }
                                                        aria-label={`删除标签 ${tag}`}
                                                        title={`删除标签 ${tag}`}
                                                    >
                                                        删除
                                                    </button>
                                                </span>
                                            ))}
                                        </div>
                                    </div>
                                ) : (
                                    <span className="empty-text">-</span>
                                )
                            ) : (
                                renderCellContent(
                                    selectedRow,
                                    column,
                                    isLatexRenderingEnabled,
                                )
                            ))}
                    </div>
                ) : null}
            </div>
        );
    };

    const renderListReadonlyCell = (row: ParsedRow, column: ParsedColumn) => {
        const cell = row.values[column.key];
        const isImageColumn = /图片/.test(column.title);
        if (isImageColumn) {
            const textValue = cell?.value?.trim() || "";
            return textValue ? (
                <div className="plain-text-value">{textValue}</div>
            ) : (
                <span className="empty-text">-</span>
            );
        }
        return renderReadonlyCell(
            row,
            {
                ...column,
                editable: false,
            },
            cell,
            false,
            "list",
        );
    };

    const getListCellTitle = (row: ParsedRow, column: ParsedColumn) =>
        getCellText(row, column.key).trim();

    return {
        renderDetailField,
        renderListReadonlyCell,
        getListCellTitle,
    };
};
