import type {
    ParsedCell,
    ParsedColumn,
    ParsedRow,
    ParsedWorkbook,
} from "./types.js";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import {
    getImageExtFromPathLike,
    LOCAL_IMAGE_API_PATH,
    normalizeCrossPlatformAbsolutePath,
    resolveImagePathLike,
} from "./utils/images.js";

const LEVEL1_ALIASES = ["level1"];
const LEVEL2_ALIASES = ["level2"];
const FILE_SEARCH_SKIP_DIRS = new Set([
    ".git",
    "node_modules",
    "Library",
    ".Trash",
    "Windows",
    "$RECYCLE.BIN",
    "System Volume Information",
]);
const SOURCE_FILE_SEARCH_MAX_DEPTH = 6;

function isRecord(value: unknown): value is Record<string, unknown> {
    return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function normalizeHeaderTitle(value: string): string {
    return value.replace(/\s+/g, "").toLowerCase();
}

function matchesHeader(title: string, aliases: string[]): boolean {
    const normalizedTitle = normalizeHeaderTitle(title);
    return aliases.some(
        (alias) => normalizeHeaderTitle(alias) === normalizedTitle,
    );
}

function getDistinctOptions(rows: ParsedRow[], columnKey?: string): string[] {
    if (!columnKey) {
        return [];
    }

    const values = new Set<string>();
    rows.forEach((row) => {
        const value = row.values[columnKey]?.value?.trim();
        if (value) {
            values.add(value);
        }
    });
    return Array.from(values);
}

function normalizeOptionalText(value: unknown): string {
    if (value === null || value === undefined) {
        return "";
    }
    if (
        typeof value === "string" ||
        typeof value === "number" ||
        typeof value === "boolean"
    ) {
        return String(value);
    }
    throw new Error("文本单元格值必须是字符串、数字、布尔值或 null");
}

function normalizeStringArray(value: unknown, fieldName: string): string[] {
    if (value === null || value === undefined) {
        return [];
    }
    if (!Array.isArray(value)) {
        throw new Error(`${fieldName} 必须是字符串数组`);
    }
    return value
        .filter((item): item is string => typeof item === "string")
        .map((item) => item.trim())
        .filter((item) => item.length > 0);
}

function isPathExistingDir(candidatePath: string): boolean {
    try {
        return fs.existsSync(candidatePath) && fs.statSync(candidatePath).isDirectory();
    } catch {
        return false;
    }
}

function normalizeImageSource(value: string, sourceDir?: string | null): string {
    const trimmed = value.trim();
    if (!trimmed) {
        return trimmed;
    }
    if (
        /^https?:\/\//i.test(trimmed) ||
        /^data:image\//i.test(trimmed) ||
        trimmed.startsWith(LOCAL_IMAGE_API_PATH)
    ) {
        return trimmed;
    }

    const absolutePath = resolveImagePathLike(trimmed, sourceDir);
    if (absolutePath) {
        return `${LOCAL_IMAGE_API_PATH}?path=${encodeURIComponent(absolutePath)}`;
    }

    return trimmed;
}

function isImageLikeString(value: string): boolean {
    return getImageExtFromPathLike(value.trim()) !== null;
}

function isLikelyRelativeImagePath(value: string): boolean {
    const trimmed = value.trim();
    if (!trimmed || !isImageLikeString(trimmed)) {
        return false;
    }
    if (
        /^https?:\/\//i.test(trimmed) ||
        /^data:image\//i.test(trimmed) ||
        trimmed.startsWith(LOCAL_IMAGE_API_PATH) ||
        normalizeCrossPlatformAbsolutePath(trimmed)
    ) {
        return false;
    }
    return true;
}

function getDefaultSearchRoots(): string[] {
    const roots: string[] = [];
    const envRootsRaw = (
        process.env.BENCHMARK_IMAGE_ROOTS ?? process.env.BENCHMARK_IMAGE_ROOT
    )?.trim();
    if (envRootsRaw) {
        envRootsRaw
            .split(/[;,]/)
            .map((item) => item.trim())
            .filter((item) => item.length > 0)
            .forEach((item) => {
                const normalized = normalizeCrossPlatformAbsolutePath(item);
                roots.push(normalized ?? path.resolve(item));
            });
    }

    roots.push(process.cwd());

    const home = os.homedir();
    if (home) {
        roots.push(home);
        roots.push(path.join(home, "Downloads"));
        roots.push(path.join(home, "Desktop"));
        roots.push(path.join(home, "Documents"));
    }

    if (process.platform === "linux" && isPathExistingDir("/mnt")) {
        try {
            const mounts = fs
                .readdirSync("/mnt", { withFileTypes: true })
                .filter((entry) => entry.isDirectory())
                .map((entry) => path.join("/mnt", entry.name));
            roots.push(...mounts);
        } catch {
            // ignore
        }
    }

    return Array.from(new Set(roots)).filter((item) => isPathExistingDir(item));
}

function findNamedFiles(
    rootPath: string,
    targetName: string,
    depth: number,
    results: string[],
    maxResults: number,
): void {
    if (depth < 0 || results.length >= maxResults || !isPathExistingDir(rootPath)) {
        return;
    }

    let entries: fs.Dirent[] = [];
    try {
        entries = fs.readdirSync(rootPath, { withFileTypes: true });
    } catch {
        return;
    }

    for (const entry of entries) {
        if (!entry.isFile() || entry.name !== targetName) {
            continue;
        }
        results.push(path.join(rootPath, entry.name));
        if (results.length >= maxResults) {
            return;
        }
    }

    if (depth === 0) {
        return;
    }

    for (const entry of entries) {
        if (!entry.isDirectory()) {
            continue;
        }
        if (FILE_SEARCH_SKIP_DIRS.has(entry.name)) {
            continue;
        }
        findNamedFiles(
            path.join(rootPath, entry.name),
            targetName,
            depth - 1,
            results,
            maxResults,
        );
        if (results.length >= maxResults) {
            return;
        }
    }
}

function extractRelativeImageCandidateFromCell(value: unknown): string | null {
    if (typeof value === "string") {
        return isLikelyRelativeImagePath(value) ? value.trim() : null;
    }
    if (Array.isArray(value)) {
        for (const item of value) {
            if (typeof item === "string" && isLikelyRelativeImagePath(item)) {
                return item.trim();
            }
        }
        return null;
    }
    if (!isRecord(value)) {
        return null;
    }
    if (typeof value.src === "string" && isLikelyRelativeImagePath(value.src)) {
        return value.src.trim();
    }
    if (Array.isArray(value.srcList)) {
        for (const item of value.srcList) {
            if (typeof item === "string" && isLikelyRelativeImagePath(item)) {
                return item.trim();
            }
        }
    }
    return null;
}

function extractRelativeImageCandidate(payload: unknown): string | null {
    if (Array.isArray(payload)) {
        for (let index = 0; index < Math.min(payload.length, 10); index += 1) {
            const row = payload[index];
            if (!isRecord(row)) {
                continue;
            }
            for (const value of Object.values(row)) {
                const candidate = extractRelativeImageCandidateFromCell(value);
                if (candidate) {
                    return candidate;
                }
            }
        }
        return null;
    }
    if (!isRecord(payload)) {
        return null;
    }
    if (Array.isArray(payload.rows)) {
        for (let index = 0; index < Math.min(payload.rows.length, 10); index += 1) {
            const row = payload.rows[index];
            if (!isRecord(row) || !isRecord(row.values)) {
                continue;
            }
            for (const value of Object.values(row.values)) {
                const candidate = extractRelativeImageCandidateFromCell(value);
                if (candidate) {
                    return candidate;
                }
            }
        }
    }
    return null;
}

function locateJsonSourceDir(
    fileName: string,
    payload: unknown,
    explicitRoots?: string[],
): string | null {
    const imageCandidate = extractRelativeImageCandidate(payload);
    if (!imageCandidate) {
        return null;
    }
    const roots = explicitRoots?.length ? explicitRoots : getDefaultSearchRoots();
    const matches: string[] = [];

    for (const root of roots) {
        findNamedFiles(
            root,
            fileName,
            SOURCE_FILE_SEARCH_MAX_DEPTH,
            matches,
            20,
        );
        if (matches.length >= 20) {
            break;
        }
    }

    if (matches.length === 0) {
        return null;
    }

    if (imageCandidate) {
        for (const match of matches) {
            const candidateDir = path.dirname(match);
            const resolvedImage = resolveImagePathLike(imageCandidate, candidateDir);
            if (resolvedImage && fs.existsSync(resolvedImage)) {
                return candidateDir;
            }
        }
    }

    return matches.length === 1 ? path.dirname(matches[0]!) : null;
}

function normalizeCell(
    value: unknown,
    fieldName: string,
    sourceDir?: string | null,
): ParsedCell {
    if (
        value === null ||
        value === undefined ||
        typeof value === "string" ||
        typeof value === "number" ||
        typeof value === "boolean"
    ) {
        return {
            type: "text",
            value: normalizeOptionalText(value),
        };
    }

    if (!isRecord(value)) {
        throw new Error(`${fieldName} 必须是单元格对象或基础值`);
    }

    const explicitType =
        typeof value.type === "string" ? value.type.trim() : undefined;

    if (explicitType === "image" || "src" in value || "srcList" in value) {
        const srcList = normalizeStringArray(value.srcList, `${fieldName}.srcList`);
        const src =
            typeof value.src === "string" ? value.src.trim() : undefined;
        const mergedSrcList = Array.from(
            new Set(
                [...srcList, ...(src && src.length > 0 ? [src] : [])].filter(
                    (item) => item.length > 0,
                ),
            ),
        );
        if (mergedSrcList.length === 0) {
            throw new Error(`${fieldName} 的图片单元格缺少 src 或 srcList`);
        }
        const textValue = normalizeOptionalText(value.value);
        const normalizedSrcList = mergedSrcList.map((item) =>
            normalizeImageSource(item, sourceDir),
        );
        return textValue
            ? {
                  type: "image",
                  src: normalizedSrcList[0],
                  srcList: normalizedSrcList,
                  value: textValue,
              }
            : {
                  type: "image",
                  src: normalizedSrcList[0],
                  srcList: normalizedSrcList,
              };
    }

    if (
        explicitType === undefined ||
        explicitType === "" ||
        explicitType === "text"
    ) {
        return {
            type: "text",
            value: normalizeOptionalText(value.value),
        };
    }

    throw new Error(`${fieldName} 使用了不支持的单元格类型: ${explicitType}`);
}

function stringifyLooseValue(value: unknown, fieldName: string): string {
    try {
        return JSON.stringify(value);
    } catch {
        throw new Error(`${fieldName} 无法序列化为文本`);
    }
}

function normalizeInferredCell(
    value: unknown,
    fieldName: string,
    sourceDir?: string | null,
): ParsedCell {
    if (Array.isArray(value)) {
        const stringItems = value
            .filter((item): item is string => typeof item === "string")
            .map((item) => item.trim())
            .filter((item) => item.length > 0);
        if (
            stringItems.length > 0 &&
            stringItems.length === value.filter((item) => item != null).length &&
            stringItems.every((item) => isImageLikeString(item))
        ) {
            const srcList = stringItems.map((item) =>
                normalizeImageSource(item, sourceDir),
            );
            return {
                type: "image",
                src: srcList[0],
                srcList,
            };
        }
        return {
            type: "text",
            value: stringifyLooseValue(value, fieldName),
        };
    }

    if (
        isRecord(value) &&
        !("type" in value) &&
        !("src" in value) &&
        !("srcList" in value) &&
        !("value" in value)
    ) {
        return {
            type: "text",
            value: stringifyLooseValue(value, fieldName),
        };
    }

    if (typeof value === "string" && isImageLikeString(value)) {
        const src = normalizeImageSource(value, sourceDir);
        return {
            type: "image",
            src,
            srcList: [src],
        };
    }

    return normalizeCell(value, fieldName, sourceDir);
}

function normalizeColumns(value: unknown): ParsedColumn[] {
    if (!Array.isArray(value) || value.length === 0) {
        throw new Error("columns 必须是非空数组");
    }

    const seenKeys = new Set<string>();
    return value.map((item, index) => {
        if (!isRecord(item)) {
            throw new Error(`columns[${index}] 必须是对象`);
        }
        const key =
            typeof item.key === "string" ? item.key.trim() : "";
        const title =
            typeof item.title === "string" ? item.title.trim() : "";

        if (!key) {
            throw new Error(`columns[${index}].key 不能为空`);
        }
        if (!title) {
            throw new Error(`columns[${index}].title 不能为空`);
        }
        if (seenKeys.has(key)) {
            throw new Error(`columns 存在重复 key: ${key}`);
        }
        seenKeys.add(key);

        return {
            key,
            title,
            editable: item.editable === true,
            required: item.required === true,
        };
    });
}

function normalizeRows(
    value: unknown,
    columns: ParsedColumn[],
    sourceDir?: string | null,
): ParsedRow[] {
    if (!Array.isArray(value)) {
        throw new Error("rows 必须是数组");
    }

    return value.map((item, index) => {
        if (!isRecord(item)) {
            throw new Error(`rows[${index}] 必须是对象`);
        }
        const rawValues = item.values;
        if (!isRecord(rawValues)) {
            throw new Error(`rows[${index}].values 必须是对象`);
        }

        const values: Record<string, ParsedCell> = {};
        columns.forEach((column) => {
            values[column.key] = normalizeCell(
                rawValues[column.key],
                `rows[${index}].values.${column.key}`,
                sourceDir,
            );
        });

        const rowId =
            typeof item.rowId === "string" && item.rowId.trim().length > 0
                ? item.rowId.trim()
                : `json-row-${index + 1}`;

        return {
            rowId,
            enabled: item.enabled !== false,
            values,
        };
    });
}

function inferColumnsFromRecords(records: Record<string, unknown>[]): ParsedColumn[] {
    const keys: string[] = [];
    const seenKeys = new Set<string>();

    records.forEach((record) => {
        Object.keys(record).forEach((key) => {
            if (seenKeys.has(key)) {
                return;
            }
            seenKeys.add(key);
            keys.push(key);
        });
    });

    if (keys.length === 0) {
        throw new Error("JSON 数组中没有可用字段");
    }

    return keys.map((key) => ({
        key,
        title: key,
        editable: false,
        required:
            matchesHeader(key, LEVEL1_ALIASES) ||
            matchesHeader(key, LEVEL2_ALIASES),
    }));
}

function toLooseRowId(record: Record<string, unknown>, index: number): string {
    const candidates = [record.rowId, record.uuid, record.id];
    for (const candidate of candidates) {
        if (
            typeof candidate === "string" ||
            typeof candidate === "number" ||
            typeof candidate === "boolean"
        ) {
            const normalized = String(candidate).trim();
            if (normalized.length > 0) {
                return normalized;
            }
        }
    }
    return `json-row-${index + 1}`;
}

function normalizeLooseRows(
    records: Record<string, unknown>[],
    columns: ParsedColumn[],
    sourceDir?: string | null,
): ParsedRow[] {
    return records.map((record, index) => {
        const values: Record<string, ParsedCell> = {};
        columns.forEach((column) => {
            values[column.key] = normalizeInferredCell(
                record[column.key],
                `rows[${index}].${column.key}`,
                sourceDir,
            );
        });
        return {
            rowId: toLooseRowId(record, index),
            enabled: record.enabled !== false,
            values,
        };
    });
}

function buildParsedWorkbook(
    fileName: string,
    fileId: string,
    columns: ParsedColumn[],
    rows: ParsedRow[],
    level1Options?: unknown,
    level2Options?: unknown,
): ParsedWorkbook {
    const level1Key = columns.find((column) =>
        matchesHeader(column.title, LEVEL1_ALIASES),
    )?.key;
    const level2Key = columns.find((column) =>
        matchesHeader(column.title, LEVEL2_ALIASES),
    )?.key;

    return {
        fileId,
        fileName,
        sourceFileName: fileName,
        columns,
        rows,
        level1Options:
            level1Options === undefined
                ? getDistinctOptions(rows, level1Key)
                : normalizeStringArray(level1Options, "level1Options"),
        level2Options:
            level2Options === undefined
                ? getDistinctOptions(rows, level2Key)
                : normalizeStringArray(level2Options, "level2Options"),
    };
}

function toFormatError(message: string): Error {
    return new Error(`JSON 导入格式无效: ${message}`);
}

export async function parseJsonWorkbook(
    buffer: Buffer,
    fileName: string,
    fileId: string,
    options?: {
        sourceDir?: string | null;
        searchRoots?: string[];
    },
): Promise<ParsedWorkbook> {
    let payload: unknown;

    try {
        payload = JSON.parse(buffer.toString("utf8").replace(/^\uFEFF/, "")) as unknown;
    } catch {
        throw new Error("JSON 解析失败，请检查文件内容是否为合法 JSON");
    }

    const effectiveSourceDir =
        options?.sourceDir ??
        locateJsonSourceDir(fileName, payload, options?.searchRoots);

    if (Array.isArray(payload)) {
        try {
            const records = payload.map((item, index) => {
                if (!isRecord(item)) {
                    throw new Error(`第 ${index + 1} 条记录必须是对象`);
                }
                return item;
            });
            const columns = inferColumnsFromRecords(records);
            const rows = normalizeLooseRows(
                records,
                columns,
                effectiveSourceDir,
            );
            return buildParsedWorkbook(fileName, fileId, columns, rows);
        } catch (error) {
            const message =
                error instanceof Error ? error.message : "JSON 数组结构不符合要求";
            throw toFormatError(message);
        }
    }

    if (!isRecord(payload)) {
        throw toFormatError("根节点必须是对象或对象数组");
    }

    let columns: ParsedColumn[];
    let rows: ParsedRow[];

    try {
        columns = normalizeColumns(payload.columns);
        rows = normalizeRows(payload.rows, columns, effectiveSourceDir);
    } catch (error) {
        const message =
            error instanceof Error ? error.message : "JSON 结构不符合要求";
        throw toFormatError(message);
    }

    try {
        return buildParsedWorkbook(
            fileName,
            fileId,
            columns,
            rows,
            payload.level1Options,
            payload.level2Options,
        );
    } catch (error) {
        const message =
            error instanceof Error ? error.message : "JSON 结构不符合要求";
        throw toFormatError(message);
    }
}
