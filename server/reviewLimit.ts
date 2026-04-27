import type { ParsedColumn } from "./types.js";

const QUALIFIED_TITLE_ALIASES = ["是否合格"] as const;
const REVIEW_SUBMISSION_VALUES = new Set(["合格", "不合格"]);

export const MAX_ROW_REVIEW_COUNT = 3;

function normalizeHeaderTitle(value: string): string {
    return value.replace(/\s+/g, "").toLowerCase();
}

export function isQualifiedColumnTitle(columnTitle: string): boolean {
    const normalizedTitle = normalizeHeaderTitle(columnTitle);
    return QUALIFIED_TITLE_ALIASES.some(
        (alias) => normalizeHeaderTitle(alias) === normalizedTitle,
    );
}

function getQualifiedColumnKeys(columns: unknown): string[] {
    if (!Array.isArray(columns)) {
        return [];
    }
    return columns
        .filter(
            (column): column is ParsedColumn =>
                !!column &&
                typeof column === "object" &&
                typeof (column as ParsedColumn).key === "string" &&
                typeof (column as ParsedColumn).title === "string",
        )
        .filter((column) => isQualifiedColumnTitle(column.title))
        .map((column) => column.key);
}

function readCellValue(row: Record<string, unknown>, columnKey: string): string {
    const values = row.values;
    if (!values || typeof values !== "object") {
        return "";
    }
    const cell = (values as Record<string, unknown>)[columnKey];
    if (!cell || typeof cell !== "object") {
        return "";
    }
    const value = (cell as { value?: unknown }).value;
    if (typeof value === "string") {
        return value.trim();
    }
    if (typeof value === "number" || typeof value === "boolean") {
        return String(value).trim();
    }
    return "";
}

function isReviewSubmissionValue(value: string): boolean {
    return REVIEW_SUBMISSION_VALUES.has(value);
}

export function getRowReviewCount(row: Record<string, unknown>): number {
    const rawValue = row.reviewCount;
    if (typeof rawValue === "number" && Number.isFinite(rawValue)) {
        return Math.max(0, Math.trunc(rawValue));
    }
    if (typeof rawValue === "string" && rawValue.trim().length > 0) {
        const parsed = Number(rawValue);
        if (Number.isFinite(parsed)) {
            return Math.max(0, Math.trunc(parsed));
        }
    }
    return 0;
}

export function evaluateRowReviewSubmission(params: {
    columns: unknown;
    previousRow: Record<string, unknown>;
    nextRow: Record<string, unknown>;
}): {
    blocked: boolean;
    isReviewSubmission: boolean;
    nextReviewCount: number;
    reviewCount: number;
} {
    const { columns, previousRow, nextRow } = params;
    const reviewCount = getRowReviewCount(previousRow);
    const qualifiedColumnKeys = getQualifiedColumnKeys(columns);
    const isReviewSubmission = qualifiedColumnKeys.some((columnKey) => {
        const previousValue = readCellValue(previousRow, columnKey);
        const nextValue = readCellValue(nextRow, columnKey);
        return (
            previousValue !== nextValue && isReviewSubmissionValue(nextValue)
        );
    });

    if (!isReviewSubmission) {
        return {
            blocked: false,
            isReviewSubmission: false,
            nextReviewCount: reviewCount,
            reviewCount,
        };
    }

    if (reviewCount >= MAX_ROW_REVIEW_COUNT) {
        return {
            blocked: true,
            isReviewSubmission: true,
            nextReviewCount: reviewCount,
            reviewCount,
        };
    }

    return {
        blocked: false,
        isReviewSubmission: true,
        nextReviewCount: reviewCount + 1,
        reviewCount,
    };
}

export function withRowReviewCount(
    row: Record<string, unknown>,
    reviewCount: number,
): Record<string, unknown> {
    return {
        ...row,
        reviewCount: Math.max(0, Math.trunc(reviewCount)),
    };
}
