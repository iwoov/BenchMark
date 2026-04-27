import assert from "node:assert/strict";
import test from "node:test";
import { mergeImportedFileState } from "./importState.js";
import type { ParsedWorkbook } from "./types.js";

test("mergeImportedFileState preserves reviewCount for updated rows", () => {
    const imported: ParsedWorkbook = {
        fileId: "file-1",
        fileName: "updated.json",
        columns: [
            {
                key: "uuid",
                title: "uuid",
                editable: false,
                required: true,
            },
            {
                key: "question",
                title: "题目",
                editable: true,
                required: false,
            },
        ],
        rows: [
            {
                rowId: "row-1",
                enabled: true,
                values: {
                    uuid: { type: "text", value: "row-1" },
                    question: { type: "text", value: "新题目内容" },
                },
            },
        ],
        level1Options: [],
        level2Options: [],
    };

    const existingState = {
        fileId: "file-1",
        fileName: "current.json",
        columns: [
            {
                key: "uuid",
                title: "uuid",
                editable: false,
                required: true,
            },
            {
                key: "question",
                title: "题目",
                editable: true,
                required: false,
            },
        ],
        rows: [
            {
                rowId: "row-1",
                enabled: true,
                reviewCount: 3,
                values: {
                    uuid: { type: "text", value: "row-1" },
                    question: { type: "text", value: "旧题目内容" },
                },
                aiResults: {
                    precheck: "{\"ok\":true}",
                },
            },
        ],
        level1Options: [],
        level2Options: [],
    };

    const { state, summary } = mergeImportedFileState(imported, {
        existingState,
        projectId: "file-1",
        projectName: "current.json",
        sourceFileName: "updated.json",
    });

    const rows = (state.rows ?? []) as Array<Record<string, unknown>>;
    const row = rows[0] ?? null;
    const values = (row?.values ?? {}) as Record<string, { value?: string }>;
    assert.equal(summary.updatedCount, 1);
    assert.equal(summary.insertedCount, 0);
    assert.equal(rows.length, 1);
    assert.equal(row?.reviewCount, 3);
    assert.equal(values.question?.value, "新题目内容");
    assert.deepEqual(row?.aiResults, {
        precheck: "{\"ok\":true}",
    });
});
