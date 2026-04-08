import assert from "node:assert/strict";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import { mergeImportedFileState } from "./importState.js";
import { parseJsonWorkbook } from "./jsonParser.js";

test("parseJsonWorkbook accepts canonical workspace JSON", async () => {
    const parsed = await parseJsonWorkbook(
        Buffer.from(
            JSON.stringify({
                columns: [
                    { key: "id", title: "id", required: true },
                    { key: "level1", title: "level1" },
                    { key: "question", title: "题目" },
                    { key: "image", title: "配图" },
                ],
                rows: [
                    {
                        rowId: "row-1",
                        values: {
                            id: "question-1",
                            level1: "数学",
                            question: "1 + 1 = ?",
                            image: {
                                type: "image",
                                src: "https://example.com/q1.png",
                            },
                        },
                    },
                    {
                        values: {
                            id: "question-2",
                            level1: "数学",
                            question: 2,
                            image: {
                                srcList: ["https://example.com/q2.png"],
                                value: "题图",
                            },
                        },
                    },
                ],
            }),
        ),
        "sample.json",
        "file-1",
    );

    assert.equal(parsed.fileId, "file-1");
    assert.equal(parsed.fileName, "sample.json");
    assert.equal(parsed.columns.length, 4);
    assert.equal(parsed.rows.length, 2);
    assert.equal(parsed.rows[0]?.values.id?.value, "question-1");
    assert.equal(parsed.rows[1]?.values.question?.value, "2");
    assert.deepEqual(parsed.rows[1]?.values.image, {
        type: "image",
        src: "https://example.com/q2.png",
        srcList: ["https://example.com/q2.png"],
        value: "题图",
    });
    assert.deepEqual(parsed.level1Options, ["数学"]);
});

test("parseJsonWorkbook accepts root arrays of plain records", async () => {
    const parsed = await parseJsonWorkbook(
        Buffer.from(
            JSON.stringify([
                {
                    level1: "有机化学",
                    题目文本: "示例题目",
                    题目图片: [
                        "image/first.png",
                        "image/second.png",
                    ],
                    uuid: "row-uuid-1",
                },
                {
                    level1: "无机化学",
                    题目文本: "第二题",
                    是否合格: "不合格",
                    uuid: "row-uuid-2",
                },
            ]),
        ),
        "records.json",
        "file-2",
        {
            sourceDir: "D:\\Data\\output",
        },
    );

    assert.equal(parsed.columns[0]?.key, "level1");
    assert.equal(parsed.columns[1]?.key, "题目文本");
    assert.equal(parsed.columns[2]?.key, "题目图片");
    assert.equal(parsed.columns[3]?.key, "uuid");
    assert.equal(parsed.rows[0]?.rowId, "row-uuid-1");
    assert.deepEqual(parsed.rows[0]?.values["题目图片"], {
        type: "image",
        src: "/api/images/local?path=%2Fmnt%2Fd%2FData%2Foutput%2Fimage%2Ffirst.png",
        srcList: [
            "/api/images/local?path=%2Fmnt%2Fd%2FData%2Foutput%2Fimage%2Ffirst.png",
            "/api/images/local?path=%2Fmnt%2Fd%2FData%2Foutput%2Fimage%2Fsecond.png",
        ],
    });
    assert.equal(parsed.rows[1]?.values["题目图片"]?.type, "text");
    assert.equal(parsed.rows[1]?.values["题目图片"]?.value, "");
    assert.deepEqual(parsed.level1Options, ["有机化学", "无机化学"]);
});

test("parseJsonWorkbook can auto-locate source dir for relative images", async () => {
    const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), "json-import-"));
    const sourceDir = path.join(tempRoot, "output");
    const imageDir = path.join(sourceDir, "image");
    fs.mkdirSync(imageDir, { recursive: true });
    fs.writeFileSync(path.join(sourceDir, "data.json"), "[]");
    fs.writeFileSync(path.join(imageDir, "sample.png"), "fake");

    const parsed = await parseJsonWorkbook(
        Buffer.from(
            JSON.stringify([
                {
                    level1: "化学",
                    题目图片: ["image/sample.png"],
                    uuid: "row-1",
                },
            ]),
        ),
        "data.json",
        "file-3",
        {
            searchRoots: [tempRoot],
        },
    );

    assert.deepEqual(parsed.rows[0]?.values["题目图片"], {
        type: "image",
        src: `/api/images/local?path=${encodeURIComponent(path.join(imageDir, "sample.png"))}`,
        srcList: [
            `/api/images/local?path=${encodeURIComponent(path.join(imageDir, "sample.png"))}`,
        ],
    });
});

test("parseJsonWorkbook rejects malformed JSON text", async () => {
    await assert.rejects(
        () => parseJsonWorkbook(Buffer.from("{invalid"), "broken.json", "file-1"),
        /JSON 解析失败/,
    );
});

test("parseJsonWorkbook rejects incompatible JSON structure", async () => {
    await assert.rejects(
        () =>
            parseJsonWorkbook(
                Buffer.from(
                    JSON.stringify({
                        columns: [{ key: "id", title: "id" }],
                        rows: [{ values: [] }],
                    }),
                ),
                "invalid-shape.json",
                "file-1",
            ),
        /JSON 导入格式无效/,
    );
});

test("JSON imports still fail when the merge pipeline cannot find id or uuid", async () => {
    const parsed = await parseJsonWorkbook(
        Buffer.from(
            JSON.stringify({
                columns: [{ key: "question", title: "题目" }],
                rows: [{ values: { question: "示例题目" } }],
            }),
        ),
        "missing-id.json",
        "file-1",
    );

    assert.throws(
        () =>
            mergeImportedFileState(parsed, {
                projectId: "project-1",
                projectName: "missing-id.json",
                sourceFileName: "missing-id.json",
            }),
        /缺少必需列: id\/uuid/,
    );
});
