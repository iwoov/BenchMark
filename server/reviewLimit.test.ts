import assert from "node:assert/strict";
import test from "node:test";
import {
    MAX_ROW_REVIEW_COUNT,
    evaluateRowReviewSubmission,
    getRowReviewCount,
    isQualifiedColumnTitle,
    withRowReviewCount,
} from "./reviewLimit.js";

function createRow(
    qualifiedValue: string,
    reviewCount = 0,
): Record<string, unknown> {
    return {
        rowId: "row-1",
        enabled: true,
        reviewCount,
        values: {
            qualified: {
                type: "text",
                value: qualifiedValue,
            },
            question: {
                type: "text",
                value: "示例题目",
            },
        },
    };
}

const columns = [
    {
        key: "question",
        title: "题目",
        editable: true,
        required: false,
    },
    {
        key: "qualified",
        title: "是否合格",
        editable: true,
        required: false,
    },
];

test("isQualifiedColumnTitle matches normalized aliases", () => {
    assert.equal(isQualifiedColumnTitle(" 是否 合格 "), true);
    assert.equal(isQualifiedColumnTitle("质检结果"), false);
});

test("evaluateRowReviewSubmission increments count on verdict submit", () => {
    const result = evaluateRowReviewSubmission({
        columns,
        previousRow: createRow(""),
        nextRow: createRow("合格"),
    });

    assert.deepEqual(result, {
        blocked: false,
        isReviewSubmission: true,
        nextReviewCount: 1,
        reviewCount: 0,
    });
});

test("evaluateRowReviewSubmission ignores non-verdict edits", () => {
    const previousRow = createRow("合格", 1);
    const nextRow = {
        ...previousRow,
        values: {
            ...((previousRow.values as Record<string, unknown>) ?? {}),
            question: {
                type: "text",
                value: "更新后的题目",
            },
        },
    };

    const result = evaluateRowReviewSubmission({
        columns,
        previousRow,
        nextRow,
    });

    assert.deepEqual(result, {
        blocked: false,
        isReviewSubmission: false,
        nextReviewCount: 1,
        reviewCount: 1,
    });
});

test("evaluateRowReviewSubmission ignores clearing verdict", () => {
    const result = evaluateRowReviewSubmission({
        columns,
        previousRow: createRow("合格", 2),
        nextRow: createRow("", 999),
    });

    assert.deepEqual(result, {
        blocked: false,
        isReviewSubmission: false,
        nextReviewCount: 2,
        reviewCount: 2,
    });
});

test("evaluateRowReviewSubmission blocks submissions after three reviews", () => {
    const result = evaluateRowReviewSubmission({
        columns,
        previousRow: createRow("不合格", MAX_ROW_REVIEW_COUNT),
        nextRow: createRow("合格"),
    });

    assert.deepEqual(result, {
        blocked: true,
        isReviewSubmission: true,
        nextReviewCount: MAX_ROW_REVIEW_COUNT,
        reviewCount: MAX_ROW_REVIEW_COUNT,
    });
});

test("withRowReviewCount persists normalized count", () => {
    const nextRow = withRowReviewCount(createRow("合格"), 2.8);
    assert.equal(getRowReviewCount(nextRow), 2);
});
