import type {
    ParsedCell,
    ParsedColumn,
    ParsedRow,
    ParsedWorkbook,
} from "./types.js";

type ImportableFileState = ParsedWorkbook & {
    sourceFileName?: string;
    selectedDisplayColumnKeys?: string[];
    selectedEditableColumnKeys?: string[];
    selectedFilterColumnKeys?: string[];
    columnFilterValues?: Record<string, string>;
};

type MergeImportOptions = {
    existingState?: unknown;
    projectId: string;
    projectName: string;
    sourceFileName: string;
};

export type MergeImportSummary = {
    insertedCount: number;
    updatedCount: number;
    totalRows: number;
};

function isRecord(value: unknown): value is Record<string, unknown> {
    return Boolean(value) && typeof value === "object";
}

function normalizeHeaderTitle(value: string): string {
    return value.replace(/\s+/g, "").toLowerCase();
}

const RECORD_ID_TITLE_ALIASES = ["id", "uuid"] as const;

function toSafeStringArray(value: unknown): string[] {
    if (!Array.isArray(value)) {
        return [];
    }
    return value.filter((item): item is string => typeof item === "string");
}

function toSafeStringRecord(value: unknown): Record<string, string> {
    if (!isRecord(value)) {
        return {};
    }
    const result: Record<string, string> = {};
    Object.entries(value).forEach(([key, item]) => {
        if (typeof item === "string") {
            result[key] = item;
        }
    });
    return result;
}

function normalizeRowAIResults(value: unknown): Record<string, string> {
    if (!value) {
        return {};
    }

    let raw: Record<string, unknown> | null = null;
    if (typeof value === "string") {
        const trimmed = value.trim();
        if (trimmed.startsWith("{") && trimmed.endsWith("}")) {
            try {
                const parsed = JSON.parse(trimmed) as unknown;
                if (isRecord(parsed)) {
                    raw = parsed;
                }
            } catch {
                raw = null;
            }
        }
    } else if (isRecord(value)) {
        raw = value;
    }

    if (!raw) {
        return {};
    }

    const result: Record<string, string> = {};
    Object.entries(raw).forEach(([key, item]) => {
        if (typeof item === "string") {
            result[key] = item;
            return;
        }
        if (typeof item === "number" || typeof item === "boolean") {
            result[key] = String(item);
        }
    });
    return result;
}

function normalizeLoadedCell(value: unknown): ParsedCell {
    if (!isRecord(value)) {
        return { type: "text", value: "" };
    }

    const cellType = value.type;
    const cellValue = typeof value.value === "string" ? value.value : "";
    if (cellType === "image") {
        const srcList = Array.isArray(value.srcList)
            ? value.srcList.filter(
                  (item): item is string => typeof item === "string",
              )
            : [];
        const fallbackSrc = typeof value.src === "string" ? value.src : "";
        const nextSrcList =
            srcList.length > 0
                ? Array.from(new Set(srcList))
                : fallbackSrc
                  ? [fallbackSrc]
                  : [];
        const nextSrc = nextSrcList[0];
        if (nextSrc) {
            return cellValue.length > 0
                ? {
                      type: "image",
                      src: nextSrc,
                      srcList: nextSrcList,
                      value: cellValue,
                  }
                : {
                      type: "image",
                      src: nextSrc,
                      srcList: nextSrcList,
                  };
        }
    }

    return { type: "text", value: cellValue };
}

function cloneCell(cell: ParsedCell | undefined): ParsedCell {
    if (!cell) {
        return { type: "text", value: "" };
    }
    if (cell.type === "image") {
        const srcList = Array.isArray(cell.srcList)
            ? cell.srcList.filter(
                  (item): item is string => typeof item === "string",
              )
            : [];
        const nextSrcList =
            srcList.length > 0
                ? Array.from(new Set(srcList))
                : typeof cell.src === "string" && cell.src
                  ? [cell.src]
                  : [];
        const nextSrc = nextSrcList[0];
        return cell.value
            ? {
                  type: "image",
                  src: nextSrc,
                  srcList: nextSrcList,
                  value: cell.value,
              }
            : {
                  type: "image",
                  src: nextSrc,
                  srcList: nextSrcList,
              };
    }
    return {
        type: "text",
        value: typeof cell.value === "string" ? cell.value : "",
    };
}

function normalizeFileState(value: unknown): ImportableFileState | null {
    if (!isRecord(value)) {
        return null;
    }
    if (
        typeof value.fileId !== "string" ||
        typeof value.fileName !== "string" ||
        !Array.isArray(value.columns) ||
        !Array.isArray(value.rows)
    ) {
        return null;
    }

    const columns: ParsedColumn[] = value.columns
        .map((column): ParsedColumn | null => {
            if (!isRecord(column)) {
                return null;
            }
            if (
                typeof column.key !== "string" ||
                typeof column.title !== "string"
            ) {
                return null;
            }
            return {
                key: column.key,
                title: column.title,
                editable: column.editable === true,
                required: column.required === true,
            };
        })
        .filter((column): column is ParsedColumn => column !== null);

    if (columns.length === 0) {
        return null;
    }

    const rows: ParsedRow[] = value.rows
        .map((row): ParsedRow | null => {
            if (!isRecord(row) || typeof row.rowId !== "string") {
                return null;
            }
            const rawValues = isRecord(row.values) ? row.values : {};
            const values: Record<string, ParsedCell> = {};
            columns.forEach((column) => {
                values[column.key] = normalizeLoadedCell(rawValues[column.key]);
            });

            const legacyAIResults =
                row.ai_results ?? row.aiResult ?? row.ai_result;
            const aiResults = normalizeRowAIResults(
                row.aiResults ?? legacyAIResults,
            );

            return Object.keys(aiResults).length > 0
                ? {
                      rowId: row.rowId,
                      values,
                      aiResults,
                  }
                : {
                      rowId: row.rowId,
                      values,
                  };
        })
        .filter((row): row is ParsedRow => row !== null);

    return {
        fileId: value.fileId,
        fileName: value.fileName,
        sourceFileName:
            typeof value.sourceFileName === "string"
                ? value.sourceFileName
                : undefined,
        columns,
        rows,
        level1Options: toSafeStringArray(value.level1Options),
        level2Options: toSafeStringArray(value.level2Options),
        selectedDisplayColumnKeys: toSafeStringArray(
            value.selectedDisplayColumnKeys,
        ),
        selectedEditableColumnKeys: toSafeStringArray(
            value.selectedEditableColumnKeys,
        ),
        selectedFilterColumnKeys: toSafeStringArray(
            value.selectedFilterColumnKeys,
        ),
        columnFilterValues: toSafeStringRecord(value.columnFilterValues),
    };
}

type ColumnDescriptor = {
    column: ParsedColumn;
    identity: string;
    normalizedTitle: string;
};

function buildColumnDescriptors(columns: ParsedColumn[]): ColumnDescriptor[] {
    const occurrenceMap = new Map<string, number>();
    return columns.map((column) => {
        const normalizedTitle = normalizeHeaderTitle(column.title);
        const nextOccurrence = (occurrenceMap.get(normalizedTitle) ?? 0) + 1;
        occurrenceMap.set(normalizedTitle, nextOccurrence);
        return {
            column,
            identity: `${normalizedTitle || "__empty__"}#${nextOccurrence}`,
            normalizedTitle,
        };
    });
}

function reserveColumnKey(
    preferredKey: string,
    usedKeys: Set<string>,
    fallbackSeed: string,
): string {
    const normalizedPreferred =
        preferredKey.trim().length > 0 ? preferredKey.trim() : fallbackSeed;
    if (!usedKeys.has(normalizedPreferred)) {
        usedKeys.add(normalizedPreferred);
        return normalizedPreferred;
    }

    let counter = 2;
    while (usedKeys.has(`${normalizedPreferred}_${counter}`)) {
        counter += 1;
    }
    const nextKey = `${normalizedPreferred}_${counter}`;
    usedKeys.add(nextKey);
    return nextKey;
}

function buildMergedColumns(
    existingColumns: ParsedColumn[],
    importedColumns: ParsedColumn[],
): {
    mergedColumns: ParsedColumn[];
    existingKeyMap: Map<string, string>;
    importedKeyMap: Map<string, string>;
} {
    const existingDescriptors = buildColumnDescriptors(existingColumns);
    const importedDescriptors = buildColumnDescriptors(importedColumns);
    const existingByIdentity = new Map(
        existingDescriptors.map((descriptor) => [
            descriptor.identity,
            descriptor,
        ]),
    );
    const matchedExisting = new Set<string>();
    const mergedColumns: ParsedColumn[] = [];
    const existingKeyMap = new Map<string, string>();
    const importedKeyMap = new Map<string, string>();
    const usedKeys = new Set<string>();

    importedDescriptors.forEach((descriptor) => {
        const existingMatch = existingByIdentity.get(descriptor.identity);
        const mergedKey = reserveColumnKey(
            existingMatch?.column.key ?? descriptor.column.key,
            usedKeys,
            descriptor.column.key || descriptor.normalizedTitle || "col",
        );
        mergedColumns.push({
            key: mergedKey,
            title: descriptor.column.title,
            editable:
                existingMatch?.column.editable ?? descriptor.column.editable,
            required:
                descriptor.column.required ||
                existingMatch?.column.required === true,
        });
        importedKeyMap.set(descriptor.column.key, mergedKey);
        if (existingMatch) {
            existingKeyMap.set(existingMatch.column.key, mergedKey);
            matchedExisting.add(existingMatch.identity);
        }
    });

    existingDescriptors.forEach((descriptor) => {
        if (matchedExisting.has(descriptor.identity)) {
            if (!existingKeyMap.has(descriptor.column.key)) {
                existingKeyMap.set(
                    descriptor.column.key,
                    descriptor.column.key,
                );
            }
            return;
        }

        const mergedKey = reserveColumnKey(
            descriptor.column.key,
            usedKeys,
            descriptor.column.key || descriptor.normalizedTitle || "col",
        );
        mergedColumns.push({
            ...descriptor.column,
            key: mergedKey,
        });
        existingKeyMap.set(descriptor.column.key, mergedKey);
    });

    return {
        mergedColumns,
        existingKeyMap,
        importedKeyMap,
    };
}

function transformRows(
    rows: ParsedRow[],
    sourceColumns: ParsedColumn[],
    keyMap: Map<string, string>,
): ParsedRow[] {
    return rows.map((row) => {
        const values: Record<string, ParsedCell> = {};
        sourceColumns.forEach((column) => {
            const mergedKey = keyMap.get(column.key);
            if (!mergedKey) {
                return;
            }
            values[mergedKey] = cloneCell(row.values[column.key]);
        });

        const nextRow: ParsedRow = {
            rowId: row.rowId,
            values,
        };

        if (row.aiResults && Object.keys(row.aiResults).length > 0) {
            nextRow.aiResults = { ...row.aiResults };
        }

        return nextRow;
    });
}

function findColumnKeyByNormalizedTitle(
    columns: ParsedColumn[],
    normalizedTitle: string,
): string | undefined {
    return buildColumnDescriptors(columns).find(
        (descriptor) => descriptor.normalizedTitle === normalizedTitle,
    )?.column.key;
}

function findRecordIdColumnKey(columns: ParsedColumn[]): string | undefined {
    for (const title of RECORD_ID_TITLE_ALIASES) {
        const key = findColumnKeyByNormalizedTitle(columns, title);
        if (key) {
            return key;
        }
    }
    return undefined;
}

function getDistinctOptions(rows: ParsedRow[], columnKey?: string): string[] {
    if (!columnKey) {
        return [];
    }

    const unique = new Set<string>();
    rows.forEach((row) => {
        const value = row.values[columnKey]?.value?.trim();
        if (value) {
            unique.add(value);
        }
    });
    return Array.from(unique);
}

function getRecordId(row: ParsedRow, idColumnKey: string): string {
    return row.values[idColumnKey]?.value?.trim() ?? "";
}

function normalizeImportedRows(
    rows: ParsedRow[],
    idColumnKey: string,
): ParsedRow[] {
    const seenIds = new Set<string>();
    return rows.map((row, index) => {
        const recordId = getRecordId(row, idColumnKey);
        if (!recordId) {
            throw new Error(`第 ${index + 1} 条导入数据缺少 id/uuid`);
        }
        if (seenIds.has(recordId)) {
            throw new Error(`导入数据存在重复 id/uuid: ${recordId}`);
        }
        seenIds.add(recordId);
        return {
            rowId: recordId,
            values: row.values,
        };
    });
}

function deduplicateExistingRows(
    rows: ParsedRow[],
    idColumnKey: string,
): {
    rows: ParsedRow[];
    indexByRecordId: Map<string, number>;
} {
    const normalizedRows: ParsedRow[] = [];
    const indexByRecordId = new Map<string, number>();

    rows.forEach((row) => {
        const recordId = getRecordId(row, idColumnKey);
        if (!recordId) {
            normalizedRows.push(row);
            return;
        }

        const existingIndex = indexByRecordId.get(recordId);
        if (existingIndex === undefined) {
            indexByRecordId.set(recordId, normalizedRows.length);
            normalizedRows.push(row);
            return;
        }

        normalizedRows[existingIndex] = row;
    });

    return {
        rows: normalizedRows,
        indexByRecordId,
    };
}

export function mergeImportedFileState(
    imported: ParsedWorkbook,
    options: MergeImportOptions,
): {
    state: Record<string, unknown>;
    summary: MergeImportSummary;
} {
    const existingStateRecord = isRecord(options.existingState)
        ? options.existingState
        : {};
    const existingState = normalizeFileState(options.existingState);
    const { mergedColumns, existingKeyMap, importedKeyMap } =
        buildMergedColumns(existingState?.columns ?? [], imported.columns);
    const idColumnKey = findRecordIdColumnKey(mergedColumns);

    if (!idColumnKey) {
        throw new Error("缺少必需列: id/uuid");
    }

    const transformedImportedRows = normalizeImportedRows(
        transformRows(imported.rows, imported.columns, importedKeyMap),
        idColumnKey,
    );
    const existingRows = existingState
        ? transformRows(
              existingState.rows,
              existingState.columns,
              existingKeyMap,
          )
        : [];
    const { rows: mergedRows, indexByRecordId } = deduplicateExistingRows(
        existingRows,
        idColumnKey,
    );

    let updatedCount = 0;
    let insertedCount = 0;

    transformedImportedRows.forEach((row) => {
        const recordId = getRecordId(row, idColumnKey);
        const existingIndex = indexByRecordId.get(recordId);
        if (existingIndex === undefined) {
            indexByRecordId.set(recordId, mergedRows.length);
            mergedRows.push(row);
            insertedCount += 1;
            return;
        }
        const existingRow = mergedRows[existingIndex];
        mergedRows[existingIndex] =
            existingRow?.aiResults &&
            Object.keys(existingRow.aiResults).length > 0
                ? {
                      ...row,
                      aiResults: { ...existingRow.aiResults },
                  }
                : row;
        updatedCount += 1;
    });

    const validColumnKeys = new Set(mergedColumns.map((column) => column.key));
    const selectedDisplayColumnKeys = toSafeStringArray(
        existingStateRecord.selectedDisplayColumnKeys,
    ).filter((key) => validColumnKeys.has(key));
    const selectedEditableColumnKeys = toSafeStringArray(
        existingStateRecord.selectedEditableColumnKeys,
    ).filter((key) => validColumnKeys.has(key));
    const selectedFilterColumnKeys = toSafeStringArray(
        existingStateRecord.selectedFilterColumnKeys,
    ).filter((key) => validColumnKeys.has(key));
    const columnFilterValues = Object.fromEntries(
        Object.entries(
            toSafeStringRecord(existingStateRecord.columnFilterValues),
        ).filter(([key]) => validColumnKeys.has(key)),
    );

    const level1Key = findColumnKeyByNormalizedTitle(mergedColumns, "level1");
    const level2Key = findColumnKeyByNormalizedTitle(mergedColumns, "level2");

    const nextState: Record<string, unknown> = {
        ...existingStateRecord,
        fileId: options.projectId,
        fileName: options.projectName,
        sourceFileName: options.sourceFileName,
        columns: mergedColumns,
        rows: mergedRows,
        level1Options: getDistinctOptions(mergedRows, level1Key),
        level2Options: getDistinctOptions(mergedRows, level2Key),
    };

    if (
        Array.isArray(existingStateRecord.selectedDisplayColumnKeys) ||
        selectedDisplayColumnKeys.length > 0
    ) {
        nextState.selectedDisplayColumnKeys = selectedDisplayColumnKeys;
    }
    if (
        Array.isArray(existingStateRecord.selectedEditableColumnKeys) ||
        selectedEditableColumnKeys.length > 0
    ) {
        nextState.selectedEditableColumnKeys = selectedEditableColumnKeys;
    }
    if (
        Array.isArray(existingStateRecord.selectedFilterColumnKeys) ||
        selectedFilterColumnKeys.length > 0
    ) {
        nextState.selectedFilterColumnKeys = selectedFilterColumnKeys;
    }
    if (
        isRecord(existingStateRecord.columnFilterValues) ||
        Object.keys(columnFilterValues).length > 0
    ) {
        nextState.columnFilterValues = columnFilterValues;
    }

    return {
        state: nextState,
        summary: {
            insertedCount,
            updatedCount,
            totalRows: mergedRows.length,
        },
    };
}
