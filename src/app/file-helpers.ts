import type {
    AICleaningToolKey,
    AICleaningToolResult,
    FileViewState,
    FilterCondition,
    ParsedCell,
    ParsedColumn,
    ParsedFile,
    ParsedRow,
    AIDetectStageKey,
    StatisticsChartType,
    StatisticsConfig,
} from "../types";
import {
    EMPTY_FILTER_VALUE,
    NON_EMPTY_FILTER_VALUE,
    CREATOR_TITLE_ALIASES,
    FEEDBACK_TITLE_ALIASES,
    INSPECTOR_TITLE_ALIASES,
    OPENSOURCE_TITLE_ALIASES,
    QUALIFIED_TITLE_ALIASES,
    TIME_TITLE_ALIASES,
    AI_CLEANING_TOOL_ORDER,
    AI_STAGE_ORDER,
} from "./constants";

const LEGACY_AI_RESULT_WITH_CONFIG_KEY = "__ai_result_with_config__";
const LEGACY_AI_RESULT_WITH_CONFIG_TITLE = "AI解析结果+AI配置名";
const DEFAULT_STATISTICS_CHART_TYPE: StatisticsChartType = "bar";
const DEFAULT_STATISTICS_FIELD_TITLES = [
    "level1",
    "level2",
    "level3",
    "是否合格",
    "创建人",
    "质检员",
] as const;

export function normalizeHeaderTitle(value: string): string {
    return value.replace(/\s+/g, "").toLowerCase();
}

function matchesHeaderAlias(
    title: string,
    aliases: readonly string[],
): boolean {
    const normalizedTitle = normalizeHeaderTitle(title);
    return aliases.some(
        (alias) => normalizeHeaderTitle(alias) === normalizedTitle,
    );
}

function isLegacyAIResultColumn(column: ParsedColumn): boolean {
    if (column.key === LEGACY_AI_RESULT_WITH_CONFIG_KEY) {
        return true;
    }
    return (
        normalizeHeaderTitle(column.title) ===
        normalizeHeaderTitle(LEGACY_AI_RESULT_WITH_CONFIG_TITLE)
    );
}

function stripLegacyAIResultColumn(parsed: ParsedFile): ParsedFile {
    const legacyColumns = parsed.columns.filter((column) =>
        isLegacyAIResultColumn(column),
    );
    if (legacyColumns.length === 0) {
        return parsed;
    }
    const removedKeys = new Set(legacyColumns.map((column) => column.key));
    const columns = parsed.columns.filter(
        (column) => !removedKeys.has(column.key),
    );
    const rows = parsed.rows.map((row) => {
        const nextValues = { ...row.values };
        let changed = false;
        removedKeys.forEach((key) => {
            if (key in nextValues) {
                delete nextValues[key];
                changed = true;
            }
        });
        return changed ? { ...row, values: nextValues } : row;
    });
    return {
        ...parsed,
        columns,
        rows,
    };
}

export function isQualifiedColumnTitle(columnTitle: string): boolean {
    return matchesHeaderAlias(columnTitle, QUALIFIED_TITLE_ALIASES);
}

export function isTimeColumnTitle(columnTitle: string): boolean {
    return matchesHeaderAlias(columnTitle, TIME_TITLE_ALIASES);
}

export function isCreatorColumnTitle(columnTitle: string): boolean {
    return matchesHeaderAlias(columnTitle, CREATOR_TITLE_ALIASES);
}

export function isInspectorColumnTitle(columnTitle: string): boolean {
    return matchesHeaderAlias(columnTitle, INSPECTOR_TITLE_ALIASES);
}

export function isFeedbackColumnTitle(columnTitle: string): boolean {
    return matchesHeaderAlias(columnTitle, FEEDBACK_TITLE_ALIASES);
}

export function isOpensourceColumnTitle(columnTitle: string): boolean {
    return matchesHeaderAlias(columnTitle, OPENSOURCE_TITLE_ALIASES);
}

export function getCellImageSources(cell: ParsedCell | undefined): string[] {
    if (!cell || cell.type !== "image") {
        return [];
    }

    const list = Array.isArray(cell.srcList)
        ? cell.srcList.filter(
              (item): item is string => typeof item === "string",
          )
        : [];

    if (list.length > 0) {
        return Array.from(new Set(list));
    }

    if (typeof cell.src === "string" && cell.src.length > 0) {
        return [cell.src];
    }

    return [];
}

export function logUIImageRenderError(
    rowId: string,
    columnTitle: string,
    src: string,
): void {
    console.log(
        `[UIImageRenderError] row=${rowId} column=${columnTitle} src=${src}`,
    );
}

export function getFileNameFromDisposition(
    disposition: string | null,
): string | null {
    if (!disposition) {
        return null;
    }

    const utf8Match = /filename\*=UTF-8''([^;]+)/i.exec(disposition);
    if (utf8Match?.[1]) {
        try {
            return decodeURIComponent(utf8Match[1]);
        } catch {
            return utf8Match[1];
        }
    }

    const plainMatch = /filename="?([^";]+)"?/i.exec(disposition);
    return plainMatch?.[1] ?? null;
}

export function downloadBlob(blob: Blob, fileName: string): void {
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = fileName;
    link.style.display = "none";
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
}

function deduplicateKeys(keys: string[]): string[] {
    const seen = new Set<string>();
    const result: string[] = [];
    for (const key of keys) {
        if (!seen.has(key)) {
            seen.add(key);
            result.push(key);
        }
    }
    return result;
}

export function getAllColumnKeys(columns: ParsedColumn[]): string[] {
    return columns.map((column) => column.key);
}

export function isFilterColumnTitle(columnTitle: string): boolean {
    const normalized = normalizeHeaderTitle(columnTitle);
    return normalized === "level1" || normalized === "level2";
}

export function getFieldSignature(columns: ParsedColumn[]): string {
    return columns
        .map((column) => normalizeHeaderTitle(column.title))
        .join("|");
}

export function normalizeColumnSelection(
    columns: ParsedColumn[],
    selectedDisplayColumnKeys?: string[],
    selectedEditableColumnKeys?: string[],
): {
    displayKeys: string[];
    editableKeys: string[];
} {
    const allColumnKeys = getAllColumnKeys(columns);
    const allowedKeys = new Set(allColumnKeys);

    let editableKeys = deduplicateKeys(selectedEditableColumnKeys ?? []).filter(
        (key) => allowedKeys.has(key),
    );

    const displaySourceKeys = selectedDisplayColumnKeys ?? allColumnKeys;
    const displaySet = new Set(
        deduplicateKeys(displaySourceKeys).filter((key) =>
            allowedKeys.has(key),
        ),
    );
    editableKeys.forEach((key) => displaySet.add(key));

    return {
        displayKeys: allColumnKeys.filter((key) => displaySet.has(key)),
        editableKeys,
    };
}

function isStatisticsChartType(value: unknown): value is StatisticsChartType {
    return (
        value === "bar" ||
        value === "pie" ||
        value === "line" ||
        value === "table"
    );
}

function getDefaultStatisticsFieldKeys(columns: ParsedColumn[]): string[] {
    const matchedKeys = DEFAULT_STATISTICS_FIELD_TITLES.map(
        (title) =>
            columns.find(
                (column) => normalizeHeaderTitle(column.title) === title,
            )?.key,
    ).filter((key): key is string => typeof key === "string");

    if (matchedKeys.length > 0) {
        return deduplicateKeys(matchedKeys);
    }

    return columns.slice(0, 2).map((column) => column.key);
}

export function normalizeStatisticsConfig(
    columns: ParsedColumn[],
    statisticsConfig?: Partial<StatisticsConfig> | null,
): StatisticsConfig {
    const allowedKeys = new Set(getAllColumnKeys(columns));
    const rawSelectedFieldKeys = Array.isArray(
        statisticsConfig?.selectedFieldKeys,
    )
        ? deduplicateKeys(
              statisticsConfig.selectedFieldKeys.filter(
                  (key): key is string => typeof key === "string",
              ),
          ).filter((key) => allowedKeys.has(key))
        : getDefaultStatisticsFieldKeys(columns).filter((key) =>
              allowedKeys.has(key),
          );

    const rawChartTypeByField =
        statisticsConfig?.chartTypeByField &&
        typeof statisticsConfig.chartTypeByField === "object"
            ? statisticsConfig.chartTypeByField
            : {};

    const chartTypeByField = columns.reduce<
        Record<string, StatisticsChartType>
    >((acc, column) => {
        const chartType = rawChartTypeByField[column.key];
        acc[column.key] = isStatisticsChartType(chartType)
            ? chartType
            : DEFAULT_STATISTICS_CHART_TYPE;
        return acc;
    }, {});

    return {
        selectedFieldKeys: rawSelectedFieldKeys,
        chartTypeByField,
    };
}

function normalizeFilterConditions(
    columns: ParsedColumn[],
    filterConditions?: FilterCondition[],
): FilterCondition[] {
    const allowedKeys = new Set(getAllColumnKeys(columns));
    if (!Array.isArray(filterConditions)) {
        return [];
    }

    return filterConditions
        .map((condition, index): FilterCondition | null => {
            if (!condition || typeof condition !== "object") {
                return null;
            }
            const columnKey =
                typeof condition.columnKey === "string"
                    ? condition.columnKey
                    : "";
            const value =
                typeof condition.value === "string" ? condition.value : "";
            if (!allowedKeys.has(columnKey) || value.trim().length === 0) {
                return null;
            }
            const id =
                typeof condition.id === "string" &&
                condition.id.trim().length > 0
                    ? condition.id
                    : `${columnKey}-${index + 1}`;
            return {
                id,
                columnKey,
                value,
            };
        })
        .filter(
            (condition): condition is FilterCondition => condition !== null,
        );
}

function applyEditableConfig(
    columns: ParsedColumn[],
    editableKeys: string[],
): ParsedColumn[] {
    const editableSet = new Set(editableKeys);
    return columns.map((column) => ({
        ...column,
        editable: editableSet.has(column.key),
        required:
            isFilterColumnTitle(column.title) || editableSet.has(column.key),
    }));
}

export function toViewState(
    parsed: ParsedFile,
    selectedDisplayColumnKeys?: string[],
    selectedEditableColumnKeys?: string[],
    filterConditions?: FilterCondition[],
): FileViewState {
    const cleanedParsed = stripLegacyAIResultColumn(parsed);
    const normalized = normalizeColumnSelection(
        cleanedParsed.columns,
        selectedDisplayColumnKeys,
        selectedEditableColumnKeys,
    );
    return {
        ...cleanedParsed,
        columns: applyEditableConfig(
            cleanedParsed.columns,
            normalized.editableKeys,
        ),
        selectedDisplayColumnKeys: normalized.displayKeys,
        selectedEditableColumnKeys: normalized.editableKeys,
        filterConditions: normalizeFilterConditions(
            cleanedParsed.columns,
            filterConditions,
        ),
        statisticsConfig: normalizeStatisticsConfig(cleanedParsed.columns),
    };
}

export function applyColumnConfigToFile(
    file: FileViewState,
    selectedDisplayColumnKeys: string[],
    selectedEditableColumnKeys: string[],
): FileViewState {
    const normalized = normalizeColumnSelection(
        file.columns,
        selectedDisplayColumnKeys,
        selectedEditableColumnKeys,
    );

    return {
        ...file,
        columns: applyEditableConfig(file.columns, normalized.editableKeys),
        selectedDisplayColumnKeys: normalized.displayKeys,
        selectedEditableColumnKeys: normalized.editableKeys,
        filterConditions: normalizeFilterConditions(
            file.columns,
            file.filterConditions,
        ),
        statisticsConfig: normalizeStatisticsConfig(
            file.columns,
            file.statisticsConfig,
        ),
    };
}

function toSafeStringArray(value: unknown): string[] {
    if (!Array.isArray(value)) {
        return [];
    }
    return value.filter((item): item is string => typeof item === "string");
}

function toSafeStringRecord(value: unknown): Record<string, string> {
    if (!value || typeof value !== "object") {
        return {};
    }
    const entries = Object.entries(value as Record<string, unknown>);
    const result: Record<string, string> = {};
    entries.forEach(([key, item]) => {
        if (typeof item === "string") {
            result[key] = item;
        }
    });
    return result;
}

function toSafeFilterConditions(value: unknown): FilterCondition[] {
    if (!Array.isArray(value)) {
        return [];
    }
    return value
        .map((item, index): FilterCondition | null => {
            if (!item || typeof item !== "object") {
                return null;
            }
            const candidate = item as Partial<FilterCondition>;
            if (
                typeof candidate.columnKey !== "string" ||
                typeof candidate.value !== "string"
            ) {
                return null;
            }
            return {
                id:
                    typeof candidate.id === "string" &&
                    candidate.id.trim().length > 0
                        ? candidate.id
                        : `${candidate.columnKey}-${index + 1}`,
                columnKey: candidate.columnKey,
                value: candidate.value,
            };
        })
        .filter((item): item is FilterCondition => item !== null);
}

function normalizeRowAIResults(
    value: unknown,
): Partial<Record<AIDetectStageKey, string>> {
    if (!value) {
        return {};
    }

    let raw: Record<string, unknown> | null = null;
    if (typeof value === "string") {
        const trimmed = value.trim();
        if (trimmed.startsWith("{") && trimmed.endsWith("}")) {
            try {
                const parsed = JSON.parse(trimmed) as unknown;
                if (parsed && typeof parsed === "object") {
                    raw = parsed as Record<string, unknown>;
                }
            } catch {
                raw = null;
            }
        }
    } else if (typeof value === "object") {
        raw = value as Record<string, unknown>;
    }

    if (!raw) {
        return {};
    }

    const result: Partial<Record<AIDetectStageKey, string>> = {};
    AI_STAGE_ORDER.forEach((stageKey) => {
        const item = raw[stageKey];
        if (typeof item === "string") {
            result[stageKey] = item;
            return;
        }
        if (item === null || item === undefined) {
            return;
        }
        if (typeof item === "number" || typeof item === "boolean") {
            result[stageKey] = String(item);
            return;
        }
        if (typeof item === "object") {
            const candidate = item as Record<string, unknown>;
            const textCandidate =
                typeof candidate.resultText === "string"
                    ? candidate.resultText
                    : typeof candidate.text === "string"
                      ? candidate.text
                      : typeof candidate.value === "string"
                        ? candidate.value
                        : null;
            if (textCandidate) {
                result[stageKey] = textCandidate;
                return;
            }
            try {
                result[stageKey] = JSON.stringify(candidate);
            } catch {
                // Ignore non-serializable payloads.
            }
        }
    });
    return result;
}

function normalizeRowCleaningResults(
    value: unknown,
): Partial<Record<AICleaningToolKey, AICleaningToolResult>> {
    if (!value || typeof value !== "object") {
        return {};
    }

    const raw = value as Record<string, unknown>;
    const result: Partial<Record<AICleaningToolKey, AICleaningToolResult>> = {};
    AI_CLEANING_TOOL_ORDER.forEach((toolKey) => {
        const item = raw[toolKey];
        if (!item || typeof item !== "object") {
            return;
        }
        const candidate = item as {
            responseText?: unknown;
            parsedJsonText?: unknown;
            updatedAt?: unknown;
        };
        if (typeof candidate.responseText !== "string") {
            return;
        }
        result[toolKey] = {
            responseText: candidate.responseText,
            parsedJsonText:
                typeof candidate.parsedJsonText === "string"
                    ? candidate.parsedJsonText
                    : undefined,
            updatedAt:
                typeof candidate.updatedAt === "string"
                    ? candidate.updatedAt
                    : undefined,
        };
    });
    return result;
}

function normalizeLoadedCell(value: unknown): ParsedCell {
    if (!value || typeof value !== "object") {
        return { type: "text", value: "" };
    }

    const cell = value as Partial<ParsedCell>;
    const cellValue = typeof cell.value === "string" ? cell.value : "";
    if (cell.type === "image") {
        const srcList = Array.isArray(cell.srcList)
            ? cell.srcList.filter(
                  (item): item is string => typeof item === "string",
              )
            : [];
        const fallbackSrc = typeof cell.src === "string" ? cell.src : "";
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
                : { type: "image", src: nextSrc, srcList: nextSrcList };
        }
    }

    return { type: "text", value: cellValue };
}

function normalizeRowEnabled(value: unknown): boolean {
    return value !== false;
}

export function getDistinctOptions(
    rows: ParsedRow[],
    columnKey?: string,
): string[] {
    if (!columnKey) {
        return [];
    }
    const unique = new Set<string>();
    let hasEmpty = false;
    let hasNonEmpty = false;
    rows.forEach((row) => {
        const rawValue = row.values[columnKey]?.value ?? "";
        const value = rawValue.trim();
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

export function getLevelColumnKey(
    columns: ParsedColumn[],
    title: string,
): string | undefined {
    return columns.find(
        (column) => normalizeHeaderTitle(column.title) === title,
    )?.key;
}

export function getCellText(row: ParsedRow, columnKey: string): string {
    return row.values[columnKey]?.value ?? "";
}

export function normalizeLoadedFileState(value: unknown): FileViewState | null {
    if (!value || typeof value !== "object") {
        return null;
    }

    const candidate = value as Partial<FileViewState> & {
        selectedOptionalColumnKeys?: unknown;
        selectedFilterColumnKeys?: unknown;
        columnFilterValues?: unknown;
        level1Filter?: unknown;
        level2Filter?: unknown;
        timeFilter?: unknown;
    };
    if (
        typeof candidate.fileId !== "string" ||
        typeof candidate.fileName !== "string"
    ) {
        return null;
    }
    if (!Array.isArray(candidate.columns) || !Array.isArray(candidate.rows)) {
        return null;
    }

    const columns: ParsedColumn[] = candidate.columns
        .map((column) => {
            if (!column || typeof column !== "object") {
                return null;
            }
            const item = column as Partial<ParsedColumn>;
            if (
                typeof item.key !== "string" ||
                typeof item.title !== "string"
            ) {
                return null;
            }
            return {
                key: item.key,
                title: item.title,
                editable: item.editable === true,
                required: item.required === true,
            };
        })
        .filter((column): column is ParsedColumn => column !== null);

    if (columns.length === 0) {
        return null;
    }

    const rows: ParsedRow[] = candidate.rows
        .map((row): ParsedRow | null => {
            if (!row || typeof row !== "object") {
                return null;
            }
            const item = row as Partial<ParsedRow>;
            if (typeof item.rowId !== "string") {
                return null;
            }
            const rawValues =
                item.values && typeof item.values === "object"
                    ? (item.values as Record<string, unknown>)
                    : {};

            const values: Record<string, ParsedCell> = {};
            columns.forEach((column) => {
                values[column.key] = normalizeLoadedCell(rawValues[column.key]);
            });

            const legacyAIResults =
                (item as Record<string, unknown>).ai_results ??
                (item as Record<string, unknown>).aiResult ??
                (item as Record<string, unknown>).ai_result;
            const aiResults = normalizeRowAIResults(
                item.aiResults ?? legacyAIResults,
            );
            const cleaningResults = normalizeRowCleaningResults(
                (item as Record<string, unknown>).cleaningResults,
            );
            const nextRow: ParsedRow = {
                rowId: item.rowId,
                enabled: normalizeRowEnabled(item.enabled),
                values,
            };
            if (Object.keys(aiResults).length > 0) {
                nextRow.aiResults = aiResults;
            }
            if (Object.keys(cleaningResults).length > 0) {
                nextRow.cleaningResults = cleaningResults;
            }
            return nextRow;
        })
        .filter((row): row is ParsedRow => row !== null);

    const parsed: ParsedFile = {
        fileId: candidate.fileId,
        fileName: candidate.fileName,
        sourceFileName:
            typeof candidate.sourceFileName === "string"
                ? candidate.sourceFileName
                : undefined,
        updatedAt:
            typeof candidate.updatedAt === "string"
                ? candidate.updatedAt
                : undefined,
        columns,
        rows,
        level1Options: toSafeStringArray(candidate.level1Options),
        level2Options: toSafeStringArray(candidate.level2Options),
    };
    const cleanedParsed = stripLegacyAIResultColumn(parsed);

    const level1Key = getLevelColumnKey(cleanedParsed.columns, "level1");
    const level2Key = getLevelColumnKey(cleanedParsed.columns, "level2");
    if (cleanedParsed.level1Options.length === 0) {
        cleanedParsed.level1Options = getDistinctOptions(rows, level1Key);
    }
    if (cleanedParsed.level2Options.length === 0) {
        cleanedParsed.level2Options = getDistinctOptions(rows, level2Key);
    }

    const hasDisplayKeys = Array.isArray(candidate.selectedDisplayColumnKeys);
    const hasEditableKeys = Array.isArray(candidate.selectedEditableColumnKeys);
    const hasLegacyOptionalKeys = Array.isArray(
        candidate.selectedOptionalColumnKeys,
    );

    const displayKeysFromState = hasDisplayKeys
        ? toSafeStringArray(candidate.selectedDisplayColumnKeys)
        : hasLegacyOptionalKeys
          ? toSafeStringArray(candidate.selectedOptionalColumnKeys)
          : undefined;
    const editableKeysFromState = hasEditableKeys
        ? toSafeStringArray(candidate.selectedEditableColumnKeys)
        : cleanedParsed.columns
              .filter((column) => column.editable)
              .map((column) => column.key);
    const filterConditionsFromState = toSafeFilterConditions(
        (candidate as { filterConditions?: unknown }).filterConditions,
    );

    const normalized = toViewState(
        cleanedParsed,
        displayKeysFromState,
        editableKeysFromState,
        filterConditionsFromState,
    );
    const legacyFilterValues: Record<string, string> = {};
    const legacyLevel1Key = getLevelColumnKey(cleanedParsed.columns, "level1");
    const legacyLevel2Key = getLevelColumnKey(cleanedParsed.columns, "level2");
    const legacyTimeKey = cleanedParsed.columns.find((column) =>
        isTimeColumnTitle(column.title),
    )?.key;
    if (typeof candidate.level1Filter === "string" && legacyLevel1Key) {
        legacyFilterValues[legacyLevel1Key] = candidate.level1Filter;
    }
    if (typeof candidate.level2Filter === "string" && legacyLevel2Key) {
        legacyFilterValues[legacyLevel2Key] = candidate.level2Filter;
    }
    if (typeof candidate.timeFilter === "string" && legacyTimeKey) {
        legacyFilterValues[legacyTimeKey] = candidate.timeFilter;
    }
    const legacyFilterConditions = Object.entries(legacyFilterValues).map(
        ([columnKey, value], index): FilterCondition => ({
            id: `legacy-${columnKey}-${index + 1}`,
            columnKey,
            value,
        }),
    );
    const selectedFilterColumnKeys = Array.isArray(
        candidate.selectedFilterColumnKeys,
    )
        ? toSafeStringArray(candidate.selectedFilterColumnKeys)
        : [];
    const rawFilterValues = toSafeStringRecord(candidate.columnFilterValues);
    const legacySelectedFilterConditions = selectedFilterColumnKeys
        .map((key, index): FilterCondition | null => {
            const value = rawFilterValues[key];
            if (typeof value !== "string" || value.trim().length === 0) {
                return null;
            }
            return {
                id: `legacy-selected-${key}-${index + 1}`,
                columnKey: key,
                value,
            };
        })
        .filter((item): item is FilterCondition => item !== null);
    const mergedFilterConditions =
        normalized.filterConditions.length > 0
            ? normalized.filterConditions
            : legacyFilterConditions.length > 0
              ? legacyFilterConditions
              : legacySelectedFilterConditions;
    return {
        ...normalized,
        filterConditions: normalizeFilterConditions(
            cleanedParsed.columns,
            mergedFilterConditions,
        ),
        statisticsConfig: normalizeStatisticsConfig(
            cleanedParsed.columns,
            candidate.statisticsConfig,
        ),
    };
}
