import type { Express } from "express";
import type { Multer } from "multer";
import { randomUUID } from "node:crypto";
import path from "node:path";
import { mergeImportedFileState } from "../importState.js";
import { parseWorkbook } from "../excelParser.js";
import { parseJsonWorkbook } from "../jsonParser.js";
import {
    deleteFileAICleaningResults,
    createDatabaseBackup,
    deleteFileState,
    listFileAICleaningResults,
    getFileState,
    getColumnPrefs,
    listFileStates,
    saveColumnPrefs,
    saveFileAICleaningToolResult,
    saveFileState,
    updateFileStateAIResults,
    type AICleaningToolKey,
    type FileAICleaningToolResult,
} from "../db.js";
type AIDetectStageKey =
    | "precheck"
    | "context_audit"
    | "independent_solving"
    | "final_verdict";
const AI_CLEANING_TOOL_ORDER: AICleaningToolKey[] = [
    "generate_level3_tags",
    "biochem_level1_refine",
];

const AI_STAGE_ORDER: AIDetectStageKey[] = [
    "precheck",
    "context_audit",
    "independent_solving",
    "final_verdict",
];

const isAIDetectStageKey = (value: unknown): value is AIDetectStageKey =>
    AI_STAGE_ORDER.includes(value as AIDetectStageKey);
const isAICleaningToolKey = (value: unknown): value is AICleaningToolKey =>
    AI_CLEANING_TOOL_ORDER.includes(value as AICleaningToolKey);

const SHOULD_LOG_AI_RESULTS = process.env.DEBUG_AI_RESULTS === "1";
const EXCEL_EXTENSIONS = new Set([".xls", ".xlsx"]);
const JSON_EXTENSIONS = new Set([".json"]);
const EXCEL_MIME_TYPES = new Set([
    "application/vnd.ms-excel",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "application/octet-stream",
]);
const JSON_MIME_TYPES = new Set([
    "application/json",
    "text/json",
    "application/ld+json",
]);

type UploadFileFormat = "excel" | "json";

function extractUploadedSourcePath(fileName: string): string | null {
    const decoded = Buffer.from(fileName, "latin1").toString("utf8");
    const source = decoded.includes("?") ? fileName : decoded;
    const matches = source.match(
        /([a-zA-Z]:[\\/][^"'<>|?*\r\n]+?\.(?:json|xlsx?|xls)|\/[^"'<>|?\r\n]+?\.(?:json|xlsx?|xls))/g,
    );
    if (!matches || matches.length === 0) {
        return null;
    }
    const candidate = matches[matches.length - 1]!.trim();
    if (/^[a-zA-Z]:[\\/]fakepath[\\/]/i.test(candidate)) {
        return null;
    }
    return candidate;
}

function normalizeUploadedFileName(fileName: string): string {
    const decoded = Buffer.from(fileName, "latin1").toString("utf8");
    const source = decoded.includes("?") ? fileName : decoded;
    const trimmed = source.trim().replace(/\0/g, "");
    const matchedNames = trimmed.match(/[^"'\\/]+\.(?:json|xlsx?|xls)/gi);
    if (matchedNames && matchedNames.length > 0) {
        return matchedNames[matchedNames.length - 1]!.trim();
    }

    const pathParts = trimmed
        .split(/[\\/]/)
        .map((item) => item.replace(/^"+|"+$/g, "").trim())
        .filter((item) => item.length > 0);
    if (pathParts.length > 0) {
        return pathParts[pathParts.length - 1]!;
    }

    return trimmed.replace(/^"+|"+$/g, "");
}

function normalizeUploadMode(value: unknown): "create" | "merge" {
    return value === "merge" ? "merge" : "create";
}

function detectUploadFileFormat(
    fileName: string,
    mimeType: string | undefined,
): UploadFileFormat | null {
    const extension = path.extname(fileName).toLowerCase();
    if (JSON_EXTENSIONS.has(extension)) {
        return "json";
    }
    if (EXCEL_EXTENSIONS.has(extension)) {
        return "excel";
    }

    const normalizedMimeType = mimeType?.trim().toLowerCase() ?? "";
    if (JSON_MIME_TYPES.has(normalizedMimeType)) {
        return "json";
    }
    if (EXCEL_MIME_TYPES.has(normalizedMimeType)) {
        return "excel";
    }

    return null;
}

async function parseUploadedFile(
    file: Express.Multer.File,
    normalizedFileName: string,
    projectId: string,
): Promise<Awaited<ReturnType<typeof parseWorkbook>>> {
    const format = detectUploadFileFormat(normalizedFileName, file.mimetype);
    const sourcePath = extractUploadedSourcePath(file.originalname);
    const sourceDir = sourcePath ? path.dirname(sourcePath) : null;
    if (format === "json") {
        return parseJsonWorkbook(file.buffer, normalizedFileName, projectId, {
            sourceDir,
        });
    }
    if (format === "excel") {
        return parseWorkbook(file.buffer, normalizedFileName, projectId);
    }
    throw new Error("仅支持导入 Excel(.xls/.xlsx) 或 JSON(.json) 文件");
}

function readClientStateVersion(value: unknown): number | null {
    if (!value || typeof value !== "object") {
        return null;
    }
    const rawVersion = (value as { clientStateVersion?: unknown })
        .clientStateVersion;
    if (typeof rawVersion === "number" && Number.isFinite(rawVersion)) {
        return Math.trunc(rawVersion);
    }
    if (typeof rawVersion === "string" && rawVersion.trim().length > 0) {
        const parsed = Number(rawVersion);
        if (Number.isFinite(parsed)) {
            return Math.trunc(parsed);
        }
    }
    return null;
}

function attachClientStateVersion<T extends Record<string, unknown>>(
    state: T,
): T {
    const currentVersion = readClientStateVersion(state) ?? 0;
    const nextVersion = Math.max(Date.now(), currentVersion + 1);
    return {
        ...state,
        clientStateVersion: nextVersion,
    };
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

function attachCleaningResultsToState(
    state: unknown,
    cleaningResultsByTool: Partial<
        Record<AICleaningToolKey, Record<string, FileAICleaningToolResult>>
    >,
): unknown {
    if (!state || typeof state !== "object") {
        return state;
    }
    const record = state as {
        rows?: unknown;
    };
    if (!Array.isArray(record.rows)) {
        return state;
    }

    const rows = record.rows.map((row) => {
        if (!row || typeof row !== "object") {
            return row;
        }
        const candidate = row as {
            rowId?: unknown;
            cleaningResults?: unknown;
        };
        if (typeof candidate.rowId !== "string") {
            return row;
        }
        const rowId = candidate.rowId;
        const mergedCleaningResults: Partial<
            Record<AICleaningToolKey, FileAICleaningToolResult>
        > = {};
        AI_CLEANING_TOOL_ORDER.forEach((toolKey) => {
            const toolResults = cleaningResultsByTool[toolKey];
            const result = toolResults?.[rowId];
            if (result) {
                mergedCleaningResults[toolKey] = result;
            }
        });
        if (Object.keys(mergedCleaningResults).length === 0) {
            return {
                ...(row as Record<string, unknown>),
                cleaningResults: undefined,
            };
        }
        return {
            ...(row as Record<string, unknown>),
            cleaningResults: mergedCleaningResults,
        };
    });

    return {
        ...(state as Record<string, unknown>),
        rows,
    };
}

export const registerFileRoutes = (app: Express, upload: Multer) => {
    app.post("/api/files/upload", upload.single("file"), async (req, res) => {
        try {
            const file = req.file;
            if (!file) {
                return res.status(400).json({
                    message: "请先选择 Excel 或 JSON 文件",
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
            const parsed = await parseUploadedFile(
                file,
                normalizedFileName,
                projectId,
            );
            const { state, summary } = mergeImportedFileState(parsed, {
                existingState: persistedState?.state,
                projectId,
                projectName,
                sourceFileName: normalizedFileName,
            });

            const backupLabel =
                uploadMode === "merge"
                    ? `merge-${projectName}`
                    : `upload-${normalizedFileName}`;
            const backupPath = await createDatabaseBackup(backupLabel);
            // eslint-disable-next-line no-console
            console.log(
                `[DatabaseBackup] mode=${uploadMode} file=${normalizedFileName} backup=${backupPath}`,
            );

            const versionedState = attachClientStateVersion(state);
            saveFileState(projectId, projectName, versionedState);
            const responseState = attachCleaningResultsToState(
                versionedState,
                listFileAICleaningResults(projectId),
            );
            return res.json({
                file: responseState,
                summary,
                backupPath,
            });
        } catch (error) {
            const message =
                error instanceof Error ? error.message : "解析文件失败";
            return res.status(400).json({ message });
        }
    });

    app.get("/api/files", (_req, res) => {
        const files = listFileStates().map((item) =>
            attachCleaningResultsToState(
                item.state,
                listFileAICleaningResults(item.fileId),
            ),
        );
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

        const currentPersistedState = getFileState(fileId)?.state;
        const currentVersion = readClientStateVersion(currentPersistedState);
        const nextVersion = readClientStateVersion(state);
        if (
            currentVersion !== null &&
            (nextVersion === null || nextVersion <= currentVersion)
        ) {
            return res.status(409).json({
                message: "stale state ignored",
            });
        }

        const versionedState =
            nextVersion === null
                ? attachClientStateVersion(state as Record<string, unknown>)
                : state;

        if (SHOULD_LOG_AI_RESULTS) {
            const summary = summarizeFileStateAIResults(versionedState);
            if (summary) {
                // eslint-disable-next-line no-console
                console.log(
                    `[FileStateSave] fileId=${summary.fileId} fileName=${summary.fileName} rows=${summary.rows} rowsWithAI=${summary.rowsWithAI} stageCounts=${JSON.stringify(
                        summary.stageCounts,
                    )}`,
                );
            }
        }

        saveFileState(fileId, nextState.fileName, versionedState);
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

    app.put("/api/files/:fileId/cleaning-results/:toolKey", (req, res) => {
        const { fileId, toolKey } = req.params;
        const {
            rowId,
            fileName,
            responseText,
            parsedJsonText,
        } = req.body as {
            rowId?: unknown;
            fileName?: unknown;
            responseText?: unknown;
            parsedJsonText?: unknown;
        };

        if (!isAICleaningToolKey(toolKey)) {
            return res.status(400).json({ message: "toolKey is invalid" });
        }
        if (typeof rowId !== "string" || rowId.trim().length === 0) {
            return res.status(400).json({ message: "rowId must be a non-empty string" });
        }
        if (typeof fileName !== "string" || fileName.trim().length === 0) {
            return res.status(400).json({ message: "fileName must be a non-empty string" });
        }
        if (
            typeof responseText !== "string" ||
            responseText.trim().length === 0
        ) {
            return res.status(400).json({ message: "responseText must be a non-empty string" });
        }
        if (
            parsedJsonText !== undefined &&
            parsedJsonText !== null &&
            typeof parsedJsonText !== "string"
        ) {
            return res.status(400).json({ message: "parsedJsonText must be a string" });
        }

        saveFileAICleaningToolResult(
            fileId,
            fileName.trim(),
            rowId.trim(),
            toolKey,
            responseText,
            typeof parsedJsonText === "string" ? parsedJsonText : undefined,
        );
        return res.json({ ok: true });
    });

    app.delete("/api/files/:fileId", (req, res) => {
        const { fileId } = req.params;
        deleteFileState(fileId);
        deleteFileAICleaningResults(fileId);
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
