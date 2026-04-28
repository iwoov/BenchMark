import type { Express } from "express";
import type { Multer } from "multer";
import { randomUUID } from "node:crypto";
import path from "node:path";
import { mergeImportedFileState } from "../importState.js";
import { parseWorkbook } from "../excelParser.js";
import { parseJsonWorkbook } from "../jsonParser.js";
import {
    MAX_ROW_REVIEW_COUNT,
    evaluateRowReviewSubmission,
    withRowReviewCount,
} from "../reviewLimit.js";
import {
    deleteFileAICleaningResults,
    deleteFileAIEvaluationResults,
    createDatabaseBackup,
    deleteFileState,
    listFileAICleaningResults,
    listFileAIEvaluationResults,
    getFileState,
    getColumnPrefs,
    isProjectNameInUse,
    listFileStateSummaries,
    patchDataSourceName,
    renameProject,
    saveColumnPrefs,
    saveFileAICleaningToolResult,
    saveFileAIEvaluationAttemptResult,
    saveFileState,
    updateFileStateAIResults,
    type AICleaningToolKey,
    type FileAICleaningToolResult,
    type FileAIEvaluationAttemptResult,
} from "../db.js";
type AIDetectStageKey =
    | "precheck"
    | "context_audit"
    | "independent_solving"
    | "final_verdict";
const AI_CLEANING_TOOL_ORDER: AICleaningToolKey[] = [
    "generate_level3_tags",
    "level1_tag_classification",
    "biochem_level1_refine",
    "knowledge_point_tag_classification",
    "question_formatting",
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

function normalizeUploadMode(
    value: unknown,
): "create" | "merge" | "add-datasource" {
    if (value === "merge") return "merge";
    if (value === "add-datasource") return "add-datasource";
    return "create";
}

function normalizeProjectName(value: unknown): string {
    return typeof value === "string" ? value.trim() : "";
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

function attachEvaluationResultsToState(
    state: unknown,
    evaluationResultsByTask: Record<
        string,
        Record<string, FileAIEvaluationAttemptResult[]>
    >,
): unknown {
    if (!state || typeof state !== "object") {
        return state;
    }
    const record = state as { rows?: unknown };
    if (!Array.isArray(record.rows)) {
        return state;
    }

    const rows = record.rows.map((row) => {
        if (!row || typeof row !== "object") {
            return row;
        }
        const candidate = row as { rowId?: unknown };
        if (typeof candidate.rowId !== "string") {
            return row;
        }
        const rowId = candidate.rowId;
        const mergedResults: Record<string, FileAIEvaluationAttemptResult[]> =
            {};
        Object.entries(evaluationResultsByTask).forEach(([taskId, rowMap]) => {
            const attempts = rowMap[rowId];
            if (attempts && attempts.length > 0) {
                mergedResults[taskId] = attempts;
            }
        });
        if (Object.keys(mergedResults).length === 0) {
            return {
                ...(row as Record<string, unknown>),
                evaluationResults: undefined,
            };
        }
        return {
            ...(row as Record<string, unknown>),
            evaluationResults: mergedResults,
        };
    });

    return {
        ...(state as Record<string, unknown>),
        rows,
    };
}

function attachPersistedMetadataToState(
    state: unknown,
    updatedAt: string,
): unknown {
    if (!state || typeof state !== "object") {
        return state;
    }
    return {
        ...(state as Record<string, unknown>),
        updatedAt,
    };
}

const EMPTY_FILTER_VALUE = "__EMPTY_FILTER__";
const NON_EMPTY_FILTER_VALUE = "__NON_EMPTY_FILTER__";
const DEFAULT_LIST_PAGE_SIZE = 50;
const MAX_LIST_PAGE_SIZE = 500;

type NormalizedFilterCondition = {
    columnKey: string;
    value: string;
};

type StatisticsDistributionItem = {
    label: string;
    count: number;
};

function parseJsonQueryValue(value: unknown): unknown {
    if (typeof value !== "string") {
        return null;
    }
    const trimmed = value.trim();
    if (trimmed.length === 0) {
        return null;
    }
    try {
        return JSON.parse(trimmed) as unknown;
    } catch {
        return null;
    }
}

function normalizeFilterConditionsFromUnknown(
    value: unknown,
): NormalizedFilterCondition[] {
    if (!Array.isArray(value)) {
        return [];
    }
    return value
        .map((item): NormalizedFilterCondition | null => {
            if (!item || typeof item !== "object") {
                return null;
            }
            const candidate = item as {
                columnKey?: unknown;
                value?: unknown;
            };
            if (
                typeof candidate.columnKey !== "string" ||
                typeof candidate.value !== "string"
            ) {
                return null;
            }
            const columnKey = candidate.columnKey.trim();
            const filterValue = candidate.value.trim();
            if (columnKey.length === 0 || filterValue.length === 0) {
                return null;
            }
            return {
                columnKey,
                value: filterValue,
            };
        })
        .filter(
            (condition): condition is NormalizedFilterCondition =>
                condition !== null,
        );
}

function parseFilterConditionsQuery(
    value: unknown,
): NormalizedFilterCondition[] {
    return normalizeFilterConditionsFromUnknown(parseJsonQueryValue(value));
}

function parseStringArrayQuery(value: unknown): string[] {
    const parsed = parseJsonQueryValue(value);
    if (!Array.isArray(parsed)) {
        return [];
    }
    return parsed
        .filter((item): item is string => typeof item === "string")
        .map((item) => item.trim())
        .filter((item) => item.length > 0);
}

function readPositiveInteger(value: unknown, fallback: number): number {
    if (typeof value === "number" && Number.isFinite(value) && value > 0) {
        return Math.trunc(value);
    }
    if (typeof value === "string" && value.trim().length > 0) {
        const parsed = Number(value);
        if (Number.isFinite(parsed) && parsed > 0) {
            return Math.trunc(parsed);
        }
    }
    return fallback;
}

function extractStateRows(state: unknown): Array<Record<string, unknown>> {
    if (!state || typeof state !== "object") {
        return [];
    }
    const rows = (state as { rows?: unknown }).rows;
    if (!Array.isArray(rows)) {
        return [];
    }
    return rows.filter(
        (row): row is Record<string, unknown> =>
            !!row && typeof row === "object",
    );
}

function getRowIdFromRecord(row: Record<string, unknown>): string | null {
    const rowId = row.rowId;
    return typeof rowId === "string" && rowId.trim().length > 0 ? rowId : null;
}

function getRowCellTextValue(
    row: Record<string, unknown>,
    columnKey: string,
): string {
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
        return value;
    }
    if (typeof value === "number" || typeof value === "boolean") {
        return String(value);
    }
    return "";
}

function rowMatchesFilters(
    row: Record<string, unknown>,
    filterConditions: NormalizedFilterCondition[],
): boolean {
    for (const condition of filterConditions) {
        const value = getRowCellTextValue(row, condition.columnKey).trim();
        if (condition.value === EMPTY_FILTER_VALUE) {
            if (value.length !== 0) {
                return false;
            }
            continue;
        }
        if (condition.value === NON_EMPTY_FILTER_VALUE) {
            if (value.length === 0) {
                return false;
            }
            continue;
        }
        if (value !== condition.value) {
            return false;
        }
    }
    return true;
}

function filterStateRows(
    rows: Array<Record<string, unknown>>,
    filterConditions: NormalizedFilterCondition[],
): Array<Record<string, unknown>> {
    if (filterConditions.length === 0) {
        return rows;
    }
    return rows.filter((row) => rowMatchesFilters(row, filterConditions));
}

function getDistinctOptionsForColumn(
    rows: Array<Record<string, unknown>>,
    columnKey: string,
): string[] {
    const unique = new Set<string>();
    let hasEmpty = false;
    let hasNonEmpty = false;
    rows.forEach((row) => {
        const value = getRowCellTextValue(row, columnKey).trim();
        if (value.length === 0) {
            hasEmpty = true;
            return;
        }
        hasNonEmpty = true;
        unique.add(value);
    });

    const options: string[] = [];
    if (hasNonEmpty) {
        options.push(NON_EMPTY_FILTER_VALUE);
    }
    if (hasEmpty) {
        options.push(EMPTY_FILTER_VALUE);
    }
    options.push(...Array.from(unique));
    return options;
}

function buildFilterOptionsMap(state: unknown): Record<string, string[]> {
    if (!state || typeof state !== "object") {
        return {};
    }
    const candidate = state as {
        columns?: unknown;
    };
    if (!Array.isArray(candidate.columns)) {
        return {};
    }
    const rows = extractStateRows(state);
    const result: Record<string, string[]> = {};
    candidate.columns.forEach((column) => {
        if (!column || typeof column !== "object") {
            return;
        }
        const key = (column as { key?: unknown }).key;
        if (typeof key !== "string" || key.trim().length === 0) {
            return;
        }
        result[key] = getDistinctOptionsForColumn(rows, key);
    });
    return result;
}

function buildFilterOptionsForColumn(
    state: unknown,
    columnKey: string,
): string[] {
    if (columnKey.trim().length === 0) {
        return [];
    }
    return getDistinctOptionsForColumn(extractStateRows(state), columnKey);
}

function attachResultsToRows(
    fileId: string,
    rows: Array<Record<string, unknown>>,
    options?: {
        includeCleaning?: boolean;
        includeEvaluation?: boolean;
    },
): Array<Record<string, unknown>> {
    const includeCleaning = options?.includeCleaning !== false;
    const includeEvaluation = options?.includeEvaluation === true;
    let state: unknown = { rows };
    if (includeCleaning) {
        state = attachCleaningResultsToState(
            state,
            listFileAICleaningResults(fileId),
        );
    }
    if (includeEvaluation) {
        state = attachEvaluationResultsToState(
            state,
            listFileAIEvaluationResults(fileId),
        );
    }
    return extractStateRows(state);
}

function splitStatisticsFieldValues(rawValue: string): string[] {
    const trimmed = rawValue.trim();
    if (!trimmed) {
        return ["空值"];
    }
    const parts = trimmed
        .split(/[\n,，;；、|]+/)
        .map((item) => item.trim())
        .filter((item) => item.length > 0);
    if (parts.length === 0) {
        return ["空值"];
    }
    return Array.from(new Set(parts));
}

function buildStatisticsDistribution(
    rows: Array<Record<string, unknown>>,
    fieldKey: string,
): {
    total: number;
    distinctCount: number;
    items: StatisticsDistributionItem[];
} {
    const counts = new Map<string, number>();
    rows.forEach((row) => {
        splitStatisticsFieldValues(getRowCellTextValue(row, fieldKey)).forEach(
            (label) => {
                counts.set(label, (counts.get(label) ?? 0) + 1);
            },
        );
    });
    const items = Array.from(counts.entries())
        .map(([label, count]) => ({ label, count }))
        .sort((left, right) => {
            if (right.count !== left.count) {
                return right.count - left.count;
            }
            return left.label.localeCompare(right.label, "zh-CN");
        });
    return {
        total: items.reduce((sum, item) => sum + item.count, 0),
        distinctCount: items.length,
        items,
    };
}

function replaceStateRow(
    state: unknown,
    rowId: string,
    nextRow: Record<string, unknown>,
): {
    nextState: Record<string, unknown> | null;
    updated: boolean;
} {
    if (!state || typeof state !== "object") {
        return { nextState: null, updated: false };
    }
    const current = state as {
        rows?: unknown;
        clientStateVersion?: unknown;
    };
    if (!Array.isArray(current.rows)) {
        return { nextState: null, updated: false };
    }

    let updated = false;
    const nextRows = current.rows.map((row) => {
        if (!row || typeof row !== "object") {
            return row;
        }
        const currentRowId = getRowIdFromRecord(row as Record<string, unknown>);
        if (currentRowId !== rowId) {
            return row;
        }
        updated = true;
        return nextRow;
    });

    if (!updated) {
        return {
            nextState: current as Record<string, unknown>,
            updated: false,
        };
    }

    const currentVersion =
        typeof current.clientStateVersion === "number" &&
        Number.isFinite(current.clientStateVersion)
            ? Math.trunc(current.clientStateVersion)
            : 0;

    return {
        nextState: {
            ...(current as Record<string, unknown>),
            rows: nextRows,
            clientStateVersion: Math.max(Date.now(), currentVersion + 1),
        },
        updated: true,
    };
}

function sanitizeRowForState(
    row: Record<string, unknown>,
): Record<string, unknown> {
    const {
        cleaningResults: _cleaningResults,
        evaluationResults: _evaluationResults,
        ...rest
    } = row;
    return rest;
}

function toListCell(cell: unknown): unknown {
    if (!cell || typeof cell !== "object") {
        return cell;
    }
    const candidate = cell as {
        type?: unknown;
        value?: unknown;
        src?: unknown;
        srcList?: unknown;
    };
    const type = candidate.type === "image" ? "image" : "text";
    const value =
        typeof candidate.value === "string"
            ? candidate.value
            : typeof candidate.value === "number" ||
                typeof candidate.value === "boolean"
              ? String(candidate.value)
              : undefined;
    if (type === "image") {
        return {
            type: "image",
            value,
        };
    }
    return {
        type: "text",
        value,
    };
}

function toListAIResults(value: unknown): Record<string, string> | undefined {
    if (!value || typeof value !== "object") {
        return undefined;
    }
    const entries = Object.entries(value as Record<string, unknown>)
        .filter(
            ([, item]) => typeof item === "string" && item.trim().length > 0,
        )
        .map(([key]) => [key, "1"] as const);
    if (entries.length === 0) {
        return undefined;
    }
    return Object.fromEntries(entries);
}

function toListCleaningResults(
    value: unknown,
): Record<string, { responseText: string }> | undefined {
    if (!value || typeof value !== "object") {
        return undefined;
    }
    const entries = Object.entries(value as Record<string, unknown>)
        .filter(([, item]) => {
            if (!item || typeof item !== "object") {
                return false;
            }
            const responseText = (item as { responseText?: unknown })
                .responseText;
            return (
                typeof responseText === "string" &&
                responseText.trim().length > 0
            );
        })
        .map(([key]) => [key, { responseText: "1" }] as const);
    if (entries.length === 0) {
        return undefined;
    }
    return Object.fromEntries(entries);
}

function toListRow(row: Record<string, unknown>): Record<string, unknown> {
    const valuesRecord =
        row.values && typeof row.values === "object"
            ? (row.values as Record<string, unknown>)
            : {};
    const values = Object.entries(valuesRecord).reduce<Record<string, unknown>>(
        (acc, [key, cell]) => {
            acc[key] = toListCell(cell);
            return acc;
        },
        {},
    );
    const aiResults = toListAIResults(row.aiResults);
    const cleaningResults = toListCleaningResults(row.cleaningResults);
    return {
        rowId: getRowIdFromRecord(row),
        enabled: row.enabled !== false,
        values,
        ...(aiResults ? { aiResults } : {}),
        ...(cleaningResults ? { cleaningResults } : {}),
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
            const requestedProjectName = normalizeProjectName(
                req.body?.projectName,
            );
            const requestedDataSourceName =
                typeof req.body?.dataSourceName === "string"
                    ? req.body.dataSourceName.trim()
                    : "";
            const requestedTargetFileId =
                typeof req.body?.targetFileId === "string"
                    ? req.body.targetFileId.trim()
                    : "";

            // For "merge" and "add-datasource" we need the existing target file.
            const persistedState =
                (uploadMode === "merge" || uploadMode === "add-datasource") &&
                requestedTargetFileId
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
            } else if (uploadMode === "add-datasource") {
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
                if (!requestedDataSourceName) {
                    return res.status(400).json({
                        message: "缺少数据源名称",
                    });
                }
            } else {
                // "create" mode
                if (!requestedProjectName) {
                    return res.status(400).json({
                        message: "缺少项目名称",
                    });
                } else if (isProjectNameInUse(requestedProjectName)) {
                    return res.status(409).json({
                        message: "项目名称已存在，请使用其他名称",
                    });
                }
            }

            let fileId: string;
            let projectName: string;
            let dataSourceGroupId: string | undefined;
            let dataSourceName: string | undefined;
            let existingStateForMerge: unknown;

            if (uploadMode === "merge") {
                // Continue import into the same data source.
                fileId = persistedState!.fileId;
                projectName = persistedState!.fileName;
                const existingStateRaw = persistedState!.state as Record<
                    string,
                    unknown
                >;
                dataSourceGroupId =
                    typeof existingStateRaw.projectId === "string"
                        ? existingStateRaw.projectId
                        : undefined;
                dataSourceName =
                    typeof existingStateRaw.dataSourceName === "string"
                        ? existingStateRaw.dataSourceName
                        : undefined;
                existingStateForMerge = persistedState!.state;
            } else if (uploadMode === "add-datasource") {
                // New data source in existing project.
                fileId = randomUUID();
                const existingStateRaw = persistedState!.state as Record<
                    string,
                    unknown
                >;
                projectName = persistedState!.fileName;
                // The group ID is either an explicit projectId stored in state,
                // or falls back to the target file's own fileId.
                dataSourceGroupId =
                    typeof existingStateRaw.projectId === "string"
                        ? existingStateRaw.projectId
                        : persistedState!.fileId;
                dataSourceName = requestedDataSourceName;
                existingStateForMerge = undefined; // fresh data source
            } else {
                // "create" — brand new project.
                fileId = randomUUID();
                projectName = requestedProjectName;
                // For first data source use fileId as the group anchor.
                dataSourceGroupId = fileId;
                dataSourceName = requestedDataSourceName || undefined;
                existingStateForMerge = undefined;
            }

            const parsed = await parseUploadedFile(
                file,
                normalizedFileName,
                fileId,
            );
            const { state, summary } = mergeImportedFileState(parsed, {
                existingState: existingStateForMerge,
                projectId: fileId,
                projectName,
                sourceFileName: normalizedFileName,
                dataSourceGroupId,
                dataSourceName,
            });

            const backupLabel =
                uploadMode === "merge"
                    ? `merge-${projectName}`
                    : uploadMode === "add-datasource"
                      ? `add-ds-${projectName}-${requestedDataSourceName}`
                      : `upload-${normalizedFileName}`;
            const backupPath = await createDatabaseBackup(backupLabel);
            // eslint-disable-next-line no-console
            console.log(
                `[DatabaseBackup] mode=${uploadMode} file=${normalizedFileName} backup=${backupPath}`,
            );

            const versionedState = attachClientStateVersion(state);
            saveFileState(fileId, projectName, versionedState);
            const responseState = attachPersistedMetadataToState(
                attachEvaluationResultsToState(
                    attachCleaningResultsToState(
                        versionedState,
                        listFileAICleaningResults(fileId),
                    ),
                    listFileAIEvaluationResults(fileId),
                ),
                new Date().toISOString(),
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
        const summaries = listFileStateSummaries().map((item) => ({
            ...(item.state as Record<string, unknown>),
            rowCount: item.rowCount,
            updatedAt: item.updatedAt,
        }));
        return res.json({ files: summaries });
    });

    app.get("/api/files/:fileId/filter-options", (req, res) => {
        const { fileId } = req.params;
        const item = getFileState(fileId);
        if (!item) {
            return res.status(404).json({ message: "file not found" });
        }
        const columnKey =
            typeof req.query.columnKey === "string"
                ? req.query.columnKey.trim()
                : "";
        if (columnKey.length > 0) {
            return res.json({
                columnKey,
                options: buildFilterOptionsForColumn(item.state, columnKey),
            });
        }
        return res.json({
            filterOptions: buildFilterOptionsMap(item.state),
        });
    });

    app.get("/api/files/:fileId/row-ids", (req, res) => {
        const { fileId } = req.params;
        const item = getFileState(fileId);
        if (!item) {
            return res.status(404).json({ message: "file not found" });
        }
        const filterConditions = parseFilterConditionsQuery(req.query.filters);
        const rowIds = filterStateRows(
            extractStateRows(item.state),
            filterConditions,
        )
            .map((row) => getRowIdFromRecord(row))
            .filter((rowId): rowId is string => rowId !== null);
        return res.json({
            rowIds,
            totalCount: rowIds.length,
        });
    });

    app.get("/api/files/:fileId/rows", (req, res) => {
        const { fileId } = req.params;
        const item = getFileState(fileId);
        if (!item) {
            return res.status(404).json({ message: "file not found" });
        }

        const page = readPositiveInteger(req.query.page, 1);
        const requestedPageSize = readPositiveInteger(
            req.query.pageSize,
            DEFAULT_LIST_PAGE_SIZE,
        );
        const pageSize = Math.min(MAX_LIST_PAGE_SIZE, requestedPageSize);
        const filterConditions = parseFilterConditionsQuery(req.query.filters);
        const filteredRows = filterStateRows(
            extractStateRows(item.state),
            filterConditions,
        );
        const totalCount = filteredRows.length;
        const totalPages = Math.max(1, Math.ceil(totalCount / pageSize) || 1);
        const normalizedPage = Math.min(page, totalPages);
        const start = (normalizedPage - 1) * pageSize;
        const pagedRows = filteredRows.slice(start, start + pageSize);
        const rows = attachResultsToRows(fileId, pagedRows, {
            includeCleaning: true,
            includeEvaluation: false,
        }).map((row) => toListRow(row));

        return res.json({
            rows,
            totalCount,
            page: normalizedPage,
            pageSize,
            totalPages,
        });
    });

    app.post("/api/files/:fileId/rows/batch", (req, res) => {
        const { fileId } = req.params;
        const item = getFileState(fileId);
        if (!item) {
            return res.status(404).json({ message: "file not found" });
        }
        const rowIds = Array.isArray(req.body?.rowIds)
            ? req.body.rowIds
                  .filter(
                      (value: unknown): value is string =>
                          typeof value === "string",
                  )
                  .map((value: string) => value.trim())
                  .filter((value: string) => value.length > 0)
            : [];
        if (rowIds.length === 0) {
            return res
                .status(400)
                .json({ message: "rowIds must be a string array" });
        }

        const rowMap = new Map<string, Record<string, unknown>>();
        extractStateRows(item.state).forEach((row) => {
            const rowId = getRowIdFromRecord(row);
            if (rowId) {
                rowMap.set(rowId, row);
            }
        });
        const rows = rowIds
            .map(
                (rowId: string): Record<string, unknown> | null =>
                    rowMap.get(rowId) ?? null,
            )
            .filter(
                (
                    row: Record<string, unknown> | null,
                ): row is Record<string, unknown> => row !== null,
            );
        const rowsWithResults = attachResultsToRows(fileId, rows, {
            includeCleaning: true,
            includeEvaluation: true,
        });
        return res.json({ rows: rowsWithResults });
    });

    app.get("/api/files/:fileId/rows/:rowId", (req, res) => {
        const { fileId, rowId } = req.params;
        const item = getFileState(fileId);
        if (!item) {
            return res.status(404).json({ message: "file not found" });
        }
        const filterConditions = parseFilterConditionsQuery(req.query.filters);
        const filteredRows = filterStateRows(
            extractStateRows(item.state),
            filterConditions,
        );
        const index = filteredRows.findIndex(
            (row) => getRowIdFromRecord(row) === rowId,
        );
        if (index < 0) {
            return res.status(404).json({ message: "row not found" });
        }
        const row = attachResultsToRows(fileId, [filteredRows[index]!], {
            includeCleaning: true,
            includeEvaluation: true,
        })[0];
        if (!row) {
            return res.status(404).json({ message: "row not found" });
        }
        const previousRowId =
            index > 0 ? getRowIdFromRecord(filteredRows[index - 1]!) : null;
        const nextRowId =
            index < filteredRows.length - 1
                ? getRowIdFromRecord(filteredRows[index + 1]!)
                : null;
        return res.json({
            row,
            previousRowId,
            nextRowId,
            totalCount: filteredRows.length,
        });
    });

    app.get("/api/files/:fileId", (req, res) => {
        const { fileId } = req.params;
        const item = getFileState(fileId);
        if (!item) {
            return res.status(404).json({ message: "file not found" });
        }
        const file = attachPersistedMetadataToState(
            attachEvaluationResultsToState(
                attachCleaningResultsToState(
                    item.state,
                    listFileAICleaningResults(item.fileId),
                ),
                listFileAIEvaluationResults(item.fileId),
            ),
            item.updatedAt,
        );
        if (SHOULD_LOG_AI_RESULTS) {
            const summary = summarizeFileStateAIResults(file);
            if (summary) {
                // eslint-disable-next-line no-console
                console.log(
                    `[FileStateDetail] fileId=${summary.fileId} fileName=${summary.fileName} rows=${summary.rows} rowsWithAI=${summary.rowsWithAI} stageCounts=${JSON.stringify(
                        summary.stageCounts,
                    )}`,
                );
            }
        }
        return res.json({ file });
    });

    app.put("/api/files/:fileId/name", (req, res) => {
        const { fileId } = req.params;
        const nextFileName = normalizeProjectName(req.body?.fileName);

        if (!nextFileName) {
            return res.status(400).json({
                message: "fileName must be a non-empty string",
            });
        }

        const current = getFileState(fileId);
        if (!current) {
            return res.status(404).json({ message: "file state not found" });
        }

        // Use projectId from state (falls back to fileId for legacy single-datasource projects)
        const currentState = current.state as Record<string, unknown>;
        const projectId =
            typeof currentState.projectId === "string"
                ? currentState.projectId
                : fileId;

        if (
            nextFileName !== current.fileName &&
            isProjectNameInUse(nextFileName, projectId)
        ) {
            return res.status(409).json({
                message: "项目名称已存在，请使用其他名称",
            });
        }

        const renamedFiles = renameProject(fileId, nextFileName);
        if (!renamedFiles) {
            return res.status(404).json({ message: "file state not found" });
        }

        const files = renamedFiles.map((renamed) => {
            const fid = (renamed.state as Record<string, unknown>)
                .fileId as string;
            return attachPersistedMetadataToState(
                attachEvaluationResultsToState(
                    attachCleaningResultsToState(
                        renamed.state,
                        listFileAICleaningResults(fid),
                    ),
                    listFileAIEvaluationResults(fid),
                ),
                renamed.updatedAt,
            );
        });

        // Return the requested file as primary, plus all sibling datasources
        const primaryFile = files.find(
            (f) => (f as Record<string, unknown>).fileId === fileId,
        );
        return res.json({ file: primaryFile ?? files[0], files });
    });

    app.put("/api/files/:fileId/datasource-name", (req, res) => {
        const { fileId } = req.params;
        const rawName = req.body?.dataSourceName;
        const dataSourceName =
            typeof rawName === "string" ? rawName.trim() : "";

        const updated = patchDataSourceName(fileId, dataSourceName);
        if (!updated) {
            return res.status(404).json({ message: "file state not found" });
        }

        const responseState = attachPersistedMetadataToState(
            attachEvaluationResultsToState(
                attachCleaningResultsToState(
                    updated.state,
                    listFileAICleaningResults(fileId),
                ),
                listFileAIEvaluationResults(fileId),
            ),
            updated.updatedAt,
        );
        return res.json({ file: responseState });
    });

    app.put("/api/files/:fileId/state", (req, res) => {
        const { fileId } = req.params;
        const { state, preserveRows } = req.body as {
            state?: unknown;
            preserveRows?: unknown;
        };

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

        let versionedState =
            nextVersion === null
                ? attachClientStateVersion(state as Record<string, unknown>)
                : state;

        if (preserveRows === true) {
            const existingState = getFileState(fileId)?.state;
            const existingRows = extractStateRows(existingState);
            versionedState = {
                ...(versionedState as Record<string, unknown>),
                rows: existingRows,
            };
        }

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

    app.put("/api/files/:fileId/rows/:rowId", (req, res) => {
        const { fileId, rowId } = req.params;
        const item = getFileState(fileId);
        if (!item) {
            return res.status(404).json({ message: "file not found" });
        }

        const row = req.body?.row;
        if (!row || typeof row !== "object") {
            return res.status(400).json({ message: "row must be an object" });
        }
        if ((row as { rowId?: unknown }).rowId !== rowId) {
            return res
                .status(400)
                .json({ message: "row.rowId must match route rowId" });
        }

        const sanitizedRow = sanitizeRowForState(
            row as Record<string, unknown>,
        );
        const existingRow = extractStateRows(item.state).find(
            (candidate) => getRowIdFromRecord(candidate) === rowId,
        );
        if (!existingRow) {
            return res.status(404).json({ message: "row not found" });
        }

        const reviewSubmission = evaluateRowReviewSubmission({
            columns: (item.state as { columns?: unknown }).columns,
            previousRow: existingRow,
            nextRow: sanitizedRow,
        });
        if (reviewSubmission.blocked) {
            return res.status(409).json({
                message: `该题目已经审核${MAX_ROW_REVIEW_COUNT}次，无法继续提交审核`,
            });
        }

        const rowWithReviewCount = withRowReviewCount(
            sanitizedRow,
            reviewSubmission.nextReviewCount,
        );
        const { nextState, updated } = replaceStateRow(
            item.state,
            rowId,
            rowWithReviewCount,
        );
        if (!nextState) {
            return res.status(404).json({ message: "file state not found" });
        }
        if (!updated) {
            return res.status(404).json({ message: "row not found" });
        }

        saveFileState(fileId, item.fileName, nextState);
        const savedRow = attachResultsToRows(fileId, [rowWithReviewCount], {
            includeCleaning: true,
            includeEvaluation: true,
        })[0];
        return res.json({
            row: savedRow ?? row,
            updatedAt: new Date().toISOString(),
        });
    });

    app.get("/api/files/:fileId/statistics", (req, res) => {
        const { fileId } = req.params;
        const item = getFileState(fileId);
        if (!item) {
            return res.status(404).json({ message: "file not found" });
        }
        const fieldKeys = parseStringArrayQuery(req.query.fieldKeys);
        const rows = extractStateRows(item.state);
        const distributions = fieldKeys.reduce<
            Record<
                string,
                {
                    total: number;
                    distinctCount: number;
                    items: StatisticsDistributionItem[];
                }
            >
        >((acc, fieldKey) => {
            acc[fieldKey] = buildStatisticsDistribution(rows, fieldKey);
            return acc;
        }, {});
        return res.json({
            rowCount: rows.length,
            distributions,
        });
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
        const { rowId, fileName, responseText, parsedJsonText } = req.body as {
            rowId?: unknown;
            fileName?: unknown;
            responseText?: unknown;
            parsedJsonText?: unknown;
        };

        if (!isAICleaningToolKey(toolKey)) {
            return res.status(400).json({ message: "toolKey is invalid" });
        }
        if (typeof rowId !== "string" || rowId.trim().length === 0) {
            return res
                .status(400)
                .json({ message: "rowId must be a non-empty string" });
        }
        if (typeof fileName !== "string" || fileName.trim().length === 0) {
            return res
                .status(400)
                .json({ message: "fileName must be a non-empty string" });
        }
        if (
            typeof responseText !== "string" ||
            responseText.trim().length === 0
        ) {
            return res
                .status(400)
                .json({ message: "responseText must be a non-empty string" });
        }
        if (
            parsedJsonText !== undefined &&
            parsedJsonText !== null &&
            typeof parsedJsonText !== "string"
        ) {
            return res
                .status(400)
                .json({ message: "parsedJsonText must be a string" });
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

    app.put("/api/files/:fileId/evaluation-results/:taskId", (req, res) => {
        const { fileId, taskId } = req.params;
        const {
            rowId,
            fileName,
            attemptIndex,
            generationResponseText,
            generationParsedJsonText,
            judgmentResponseText,
            judgmentParsedJsonText,
            finalVerdict,
            generationLatencyMs,
            judgmentLatencyMs,
            generationInputTokens,
            generationOutputTokens,
            judgmentInputTokens,
            judgmentOutputTokens,
            generationFinishReason,
            judgmentFinishReason,
        } = req.body as {
            rowId?: unknown;
            fileName?: unknown;
            attemptIndex?: unknown;
            generationResponseText?: unknown;
            generationParsedJsonText?: unknown;
            judgmentResponseText?: unknown;
            judgmentParsedJsonText?: unknown;
            finalVerdict?: unknown;
            generationLatencyMs?: unknown;
            judgmentLatencyMs?: unknown;
            generationInputTokens?: unknown;
            generationOutputTokens?: unknown;
            judgmentInputTokens?: unknown;
            judgmentOutputTokens?: unknown;
            generationFinishReason?: unknown;
            judgmentFinishReason?: unknown;
        };
        if (typeof rowId !== "string" || rowId.trim().length === 0) {
            return res
                .status(400)
                .json({ message: "rowId must be a non-empty string" });
        }
        if (typeof fileName !== "string" || fileName.trim().length === 0) {
            return res
                .status(400)
                .json({ message: "fileName must be a non-empty string" });
        }
        if (
            typeof attemptIndex !== "number" ||
            !Number.isInteger(attemptIndex) ||
            attemptIndex < 1
        ) {
            return res
                .status(400)
                .json({ message: "attemptIndex must be a positive integer" });
        }
        if (
            typeof generationResponseText !== "string" ||
            generationResponseText.trim().length === 0
        ) {
            return res.status(400).json({
                message: "generationResponseText must be a non-empty string",
            });
        }
        if (
            typeof judgmentResponseText !== "string" ||
            judgmentResponseText.trim().length === 0
        ) {
            return res.status(400).json({
                message: "judgmentResponseText must be a non-empty string",
            });
        }
        if (
            typeof finalVerdict !== "string" ||
            finalVerdict.trim().length === 0
        ) {
            return res
                .status(400)
                .json({ message: "finalVerdict must be a non-empty string" });
        }
        if (
            generationParsedJsonText !== undefined &&
            generationParsedJsonText !== null &&
            typeof generationParsedJsonText !== "string"
        ) {
            return res
                .status(400)
                .json({ message: "generationParsedJsonText must be a string" });
        }
        if (
            judgmentParsedJsonText !== undefined &&
            judgmentParsedJsonText !== null &&
            typeof judgmentParsedJsonText !== "string"
        ) {
            return res
                .status(400)
                .json({ message: "judgmentParsedJsonText must be a string" });
        }

        saveFileAIEvaluationAttemptResult(
            fileId,
            fileName.trim(),
            rowId.trim(),
            taskId,
            attemptIndex,
            generationResponseText,
            typeof generationParsedJsonText === "string"
                ? generationParsedJsonText
                : undefined,
            judgmentResponseText,
            typeof judgmentParsedJsonText === "string"
                ? judgmentParsedJsonText
                : undefined,
            finalVerdict.trim(),
            typeof generationLatencyMs === "number"
                ? generationLatencyMs
                : undefined,
            typeof judgmentLatencyMs === "number"
                ? judgmentLatencyMs
                : undefined,
            typeof generationInputTokens === "number"
                ? generationInputTokens
                : undefined,
            typeof generationOutputTokens === "number"
                ? generationOutputTokens
                : undefined,
            typeof judgmentInputTokens === "number"
                ? judgmentInputTokens
                : undefined,
            typeof judgmentOutputTokens === "number"
                ? judgmentOutputTokens
                : undefined,
            typeof generationFinishReason === "string"
                ? generationFinishReason
                : undefined,
            typeof judgmentFinishReason === "string"
                ? judgmentFinishReason
                : undefined,
        );
        return res.json({ ok: true });
    });

    app.delete("/api/files/:fileId", (req, res) => {
        const { fileId } = req.params;
        deleteFileState(fileId);
        deleteFileAICleaningResults(fileId);
        deleteFileAIEvaluationResults(fileId);
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
