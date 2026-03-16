import type { Express } from "express";
import type { Multer } from "multer";
import { randomUUID } from "node:crypto";
import { mergeImportedFileState } from "../importState.js";
import { parseWorkbook } from "../excelParser.js";
import {
    deleteFileState,
    getFileState,
    getColumnPrefs,
    listFileStates,
    saveColumnPrefs,
    saveFileState,
    updateFileStateAIResults,
} from "../db.js";
type AIDetectStageKey =
    | "precheck"
    | "context_audit"
    | "independent_solving"
    | "final_verdict";

const AI_STAGE_ORDER: AIDetectStageKey[] = [
    "precheck",
    "context_audit",
    "independent_solving",
    "final_verdict",
];

const isAIDetectStageKey = (value: unknown): value is AIDetectStageKey =>
    AI_STAGE_ORDER.includes(value as AIDetectStageKey);

const SHOULD_LOG_AI_RESULTS = process.env.DEBUG_AI_RESULTS === "1";

function normalizeUploadedFileName(fileName: string): string {
    const decoded = Buffer.from(fileName, "latin1").toString("utf8");
    if (decoded.includes("?")) {
        return fileName;
    }
    return decoded;
}

function normalizeUploadMode(value: unknown): "create" | "merge" {
    return value === "merge" ? "merge" : "create";
}

function summarizeFileStateAIResults(state: unknown): {
    fileId: string;
    fileName: string;
    rows: number;
    rowsWithAI: number;
    stageCounts: Record<AIDetectStageKey, number>;
} | null {
    if (!state || typeof state !== "object") {
        return null;
    }
    const record = state as {
        fileId?: unknown;
        fileName?: unknown;
        rows?: unknown;
    };
    if (!Array.isArray(record.rows)) {
        return null;
    }
    const fileId =
        typeof record.fileId === "string" ? record.fileId : "unknown";
    const fileName =
        typeof record.fileName === "string" ? record.fileName : "unknown";
    const stageCounts: Record<AIDetectStageKey, number> = {
        precheck: 0,
        context_audit: 0,
        independent_solving: 0,
        final_verdict: 0,
    };
    let rowsWithAI = 0;
    record.rows.forEach((row) => {
        if (!row || typeof row !== "object") {
            return;
        }
        const aiResults = (row as { aiResults?: unknown }).aiResults;
        if (!aiResults || typeof aiResults !== "object") {
            return;
        }
        let hasAny = false;
        AI_STAGE_ORDER.forEach((stageKey) => {
            const value = (aiResults as Record<string, unknown>)[stageKey];
            if (typeof value === "string" && value.trim().length > 0) {
                stageCounts[stageKey] += 1;
                hasAny = true;
            }
        });
        if (hasAny) {
            rowsWithAI += 1;
        }
    });
    return {
        fileId,
        fileName,
        rows: record.rows.length,
        rowsWithAI,
        stageCounts,
    };
}

export const registerFileRoutes = (app: Express, upload: Multer) => {
    app.post("/api/files/upload", upload.single("file"), async (req, res) => {
        try {
            const file = req.file;
            if (!file) {
                return res.status(400).json({
                    message: "请先选择 Excel 文件",
                });
            }

            const normalizedFileName = normalizeUploadedFileName(
                file.originalname,
            );
            const uploadMode = normalizeUploadMode(req.body?.mode);
            const requestedTargetFileId =
                typeof req.body?.targetFileId === "string"
                    ? req.body.targetFileId.trim()
                    : "";
            const persistedState =
                uploadMode === "merge" && requestedTargetFileId
                    ? getFileState(requestedTargetFileId)
                    : null;

            if (uploadMode === "merge") {
                if (!requestedTargetFileId) {
                    return res.status(400).json({
                        message: "缺少目标项目 ID",
                    });
                }
                if (!persistedState) {
                    return res.status(404).json({
                        message: "目标项目不存在",
                    });
                }
            }

            const projectId = persistedState?.fileId ?? randomUUID();
            const projectName = persistedState?.fileName ?? normalizedFileName;
            const parsed = await parseWorkbook(
                file.buffer,
                normalizedFileName,
                projectId,
            );
            const { state, summary } = mergeImportedFileState(parsed, {
                existingState: persistedState?.state,
                projectId,
                projectName,
                sourceFileName: normalizedFileName,
            });

            saveFileState(projectId, projectName, state);
            return res.json({
                file: state,
                summary,
            });
        } catch (error) {
            const message =
                error instanceof Error ? error.message : "解析 Excel 失败";
            return res.status(400).json({ message });
        }
    });

    app.get("/api/files", (_req, res) => {
        const files = listFileStates().map((item) => item.state);
        if (SHOULD_LOG_AI_RESULTS) {
            files.forEach((state) => {
                const summary = summarizeFileStateAIResults(state);
                if (!summary) {
                    return;
                }
                // eslint-disable-next-line no-console
                console.log(
                    `[FileStateList] fileId=${summary.fileId} fileName=${summary.fileName} rows=${summary.rows} rowsWithAI=${summary.rowsWithAI} stageCounts=${JSON.stringify(
                        summary.stageCounts,
                    )}`,
                );
            });
        }
        return res.json({ files });
    });

    app.put("/api/files/:fileId/state", (req, res) => {
        const { fileId } = req.params;
        const { state } = req.body as { state?: unknown };

        if (!state || typeof state !== "object") {
            return res.status(400).json({ message: "state must be an object" });
        }

        const nextState = state as { fileId?: unknown; fileName?: unknown };
        if (
            typeof nextState.fileId !== "string" ||
            nextState.fileId !== fileId
        ) {
            return res
                .status(400)
                .json({ message: "state.fileId must match route fileId" });
        }
        if (
            typeof nextState.fileName !== "string" ||
            nextState.fileName.trim().length === 0
        ) {
            return res
                .status(400)
                .json({ message: "state.fileName must be a non-empty string" });
        }

        if (SHOULD_LOG_AI_RESULTS) {
            const summary = summarizeFileStateAIResults(state);
            if (summary) {
                // eslint-disable-next-line no-console
                console.log(
                    `[FileStateSave] fileId=${summary.fileId} fileName=${summary.fileName} rows=${summary.rows} rowsWithAI=${summary.rowsWithAI} stageCounts=${JSON.stringify(
                        summary.stageCounts,
                    )}`,
                );
            }
        }

        saveFileState(fileId, nextState.fileName, state);
        return res.json({ ok: true });
    });

    app.put("/api/files/:fileId/ai-results", (req, res) => {
        const { fileId } = req.params;
        const { stageKey, results } = req.body as {
            stageKey?: unknown;
            results?: unknown;
        };

        if (!isAIDetectStageKey(stageKey)) {
            return res.status(400).json({ message: "stageKey is invalid" });
        }
        if (!results || typeof results !== "object") {
            return res
                .status(400)
                .json({ message: "results must be an object" });
        }

        const entries = Object.entries(results as Record<string, unknown>)
            .map(([rowId, value]) => ({
                rowId,
                resultText: typeof value === "string" ? value : null,
            }))
            .filter(
                (item): item is { rowId: string; resultText: string } =>
                    typeof item.rowId === "string" &&
                    item.rowId.length > 0 &&
                    item.resultText !== null,
            );

        if (entries.length === 0) {
            return res
                .status(400)
                .json({ message: "results must have string values" });
        }

        if (SHOULD_LOG_AI_RESULTS) {
            // eslint-disable-next-line no-console
            console.log(
                `[AIResultsPersist] fileId=${fileId} stageKey=${String(
                    stageKey,
                )} entries=${entries.length}`,
            );
        }

        const updatedCount = updateFileStateAIResults(
            fileId,
            stageKey,
            entries,
        );
        if (updatedCount === null) {
            return res.status(404).json({ message: "file state not found" });
        }
        if (SHOULD_LOG_AI_RESULTS) {
            // eslint-disable-next-line no-console
            console.log(
                `[AIResultsPersist] fileId=${fileId} stageKey=${String(
                    stageKey,
                )} updated=${updatedCount}`,
            );
        }
        return res.json({ ok: true, updatedCount });
    });

    app.delete("/api/files/:fileId", (req, res) => {
        const { fileId } = req.params;
        deleteFileState(fileId);
        return res.json({ ok: true });
    });

    app.get("/api/column-prefs/:fileName", (req, res) => {
        const { fileName } = req.params;
        const config = getColumnPrefs(decodeURIComponent(fileName));
        return res.json({ config });
    });

    app.put("/api/column-prefs/:fileName", (req, res) => {
        const { fileName } = req.params;
        const { fieldSignature, displayKeys, editableKeys, filterKeys } =
            req.body as {
                fieldSignature: unknown;
                displayKeys: unknown;
                editableKeys: unknown;
                filterKeys?: unknown;
            };

        if (
            typeof fieldSignature !== "string" ||
            fieldSignature.trim().length === 0
        ) {
            return res
                .status(400)
                .json({ message: "fieldSignature must be a non-empty string" });
        }
        if (
            !Array.isArray(displayKeys) ||
            !displayKeys.every((item) => typeof item === "string")
        ) {
            return res
                .status(400)
                .json({ message: "displayKeys must be a string array" });
        }
        if (
            !Array.isArray(editableKeys) ||
            !editableKeys.every((item) => typeof item === "string")
        ) {
            return res
                .status(400)
                .json({ message: "editableKeys must be a string array" });
        }
        if (
            filterKeys !== undefined &&
            (!Array.isArray(filterKeys) ||
                !filterKeys.every((item) => typeof item === "string"))
        ) {
            return res
                .status(400)
                .json({ message: "filterKeys must be a string array" });
        }

        saveColumnPrefs(decodeURIComponent(fileName), {
            fieldSignature,
            displayKeys,
            editableKeys,
            filterKeys: Array.isArray(filterKeys) ? filterKeys : [],
        });
        return res.json({ ok: true });
    });
};
