import type {
  FileViewState,
  ParsedCell,
  ParsedColumn,
  ParsedFile,
  ParsedRow,
} from "../types";
import {
  AI_RESULT_WITH_CONFIG_COLUMN_KEY,
  AI_RESULT_WITH_CONFIG_COLUMN_TITLE,
  ALL_FILTER_VALUE,
  CREATOR_TITLE_ALIASES,
  FEEDBACK_TITLE_ALIASES,
  INSPECTOR_TITLE_ALIASES,
  OPENSOURCE_TITLE_ALIASES,
  QUALIFIED_TITLE_ALIASES,
  TIME_TITLE_ALIASES,
} from "./constants";

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
    ? cell.srcList.filter((item): item is string => typeof item === "string")
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

export function isAIResultWithConfigColumn(column: ParsedColumn): boolean {
  if (column.key === AI_RESULT_WITH_CONFIG_COLUMN_KEY) {
    return true;
  }
  return (
    normalizeHeaderTitle(column.title) ===
    normalizeHeaderTitle(AI_RESULT_WITH_CONFIG_COLUMN_TITLE)
  );
}

export function getAIResultWithConfigColumn(
  columns: ParsedColumn[],
): ParsedColumn | null {
  return columns.find((column) => isAIResultWithConfigColumn(column)) ?? null;
}

function ensureAIResultWithConfigColumn(parsed: ParsedFile): ParsedFile {
  const existingColumn = getAIResultWithConfigColumn(parsed.columns);
  const targetColumn =
    existingColumn ??
    ({
      key: AI_RESULT_WITH_CONFIG_COLUMN_KEY,
      title: AI_RESULT_WITH_CONFIG_COLUMN_TITLE,
      editable: true,
      required: false,
    } as ParsedColumn);
  const columns = existingColumn
    ? parsed.columns
    : [...parsed.columns, targetColumn];
  const targetKey = targetColumn.key;
  const rows = parsed.rows.map((row) => {
    if (row.values[targetKey]) {
      return row;
    }
    return {
      ...row,
      values: {
        ...row.values,
        [targetKey]: {
          type: "text",
          value: "",
        } as ParsedCell,
      },
    };
  });

  if (
    columns === parsed.columns &&
    rows.every((row, index) => row === parsed.rows[index])
  ) {
    return parsed;
  }
  return {
    ...parsed,
    columns,
    rows,
  };
}

export function isFilterColumnTitle(columnTitle: string): boolean {
  const normalized = normalizeHeaderTitle(columnTitle);
  return normalized === "level1" || normalized === "level2";
}

export function getFieldSignature(columns: ParsedColumn[]): string {
  return columns
    .filter((column) => !isAIResultWithConfigColumn(column))
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
  const aiResultWithConfigColumn = getAIResultWithConfigColumn(columns);
  if (
    aiResultWithConfigColumn &&
    !editableKeys.includes(aiResultWithConfigColumn.key)
  ) {
    editableKeys = [...editableKeys, aiResultWithConfigColumn.key];
  }

  const displaySourceKeys = selectedDisplayColumnKeys ?? allColumnKeys;
  const displaySet = new Set(
    deduplicateKeys(displaySourceKeys).filter((key) => allowedKeys.has(key)),
  );
  editableKeys.forEach((key) => displaySet.add(key));

  return {
    displayKeys: allColumnKeys.filter((key) => displaySet.has(key)),
    editableKeys,
  };
}

export function normalizeFilterSelection(
  columns: ParsedColumn[],
  selectedFilterColumnKeys?: string[],
): string[] {
  const allColumnKeys = getAllColumnKeys(columns);
  const allowedKeys = new Set(allColumnKeys);
  const sourceKeys = selectedFilterColumnKeys ?? [];
  return deduplicateKeys(sourceKeys).filter((key) => allowedKeys.has(key));
}

function applyEditableConfig(
  columns: ParsedColumn[],
  editableKeys: string[],
): ParsedColumn[] {
  const editableSet = new Set(editableKeys);
  return columns.map((column) => ({
    ...column,
    editable: editableSet.has(column.key),
    required: isFilterColumnTitle(column.title) || editableSet.has(column.key),
  }));
}

export function toViewState(
  parsed: ParsedFile,
  selectedDisplayColumnKeys?: string[],
  selectedEditableColumnKeys?: string[],
  selectedFilterColumnKeys?: string[],
  filterValues?: Record<string, string>,
): FileViewState {
  const nextParsed = ensureAIResultWithConfigColumn(parsed);
  const normalized = normalizeColumnSelection(
    nextParsed.columns,
    selectedDisplayColumnKeys,
    selectedEditableColumnKeys,
  );
  const normalizedFilterKeys = normalizeFilterSelection(
    nextParsed.columns,
    selectedFilterColumnKeys,
  );
  const columnFilterValues: Record<string, string> = {};
  normalizedFilterKeys.forEach((key) => {
    const value = filterValues?.[key];
    columnFilterValues[key] =
      typeof value === "string" ? value : ALL_FILTER_VALUE;
  });
  return {
    ...nextParsed,
    columns: applyEditableConfig(nextParsed.columns, normalized.editableKeys),
    selectedDisplayColumnKeys: normalized.displayKeys,
    selectedEditableColumnKeys: normalized.editableKeys,
    selectedFilterColumnKeys: normalizedFilterKeys,
    columnFilterValues,
  };
}

export function applyColumnConfigToFile(
  file: FileViewState,
  selectedDisplayColumnKeys: string[],
  selectedEditableColumnKeys: string[],
  selectedFilterColumnKeys: string[],
): FileViewState {
  const normalized = normalizeColumnSelection(
    file.columns,
    selectedDisplayColumnKeys,
    selectedEditableColumnKeys,
  );
  const normalizedFilterKeys = normalizeFilterSelection(
    file.columns,
    selectedFilterColumnKeys,
  );
  const nextFilterValues: Record<string, string> = {};
  normalizedFilterKeys.forEach((key) => {
    const value = file.columnFilterValues?.[key];
    nextFilterValues[key] =
      typeof value === "string" ? value : ALL_FILTER_VALUE;
  });

  return {
    ...file,
    columns: applyEditableConfig(file.columns, normalized.editableKeys),
    selectedDisplayColumnKeys: normalized.displayKeys,
    selectedEditableColumnKeys: normalized.editableKeys,
    selectedFilterColumnKeys: normalizedFilterKeys,
    columnFilterValues: nextFilterValues,
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

function normalizeLoadedCell(value: unknown): ParsedCell {
  if (!value || typeof value !== "object") {
    return { type: "text", value: "" };
  }

  const cell = value as Partial<ParsedCell>;
  const cellValue = typeof cell.value === "string" ? cell.value : "";
  if (cell.type === "image") {
    const srcList = Array.isArray(cell.srcList)
      ? cell.srcList.filter((item): item is string => typeof item === "string")
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

export function getDistinctOptions(
  rows: ParsedRow[],
  columnKey?: string,
): string[] {
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

export function getLevelColumnKey(
  columns: ParsedColumn[],
  title: string,
): string | undefined {
  return columns.find((column) => normalizeHeaderTitle(column.title) === title)
    ?.key;
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
      if (typeof item.key !== "string" || typeof item.title !== "string") {
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
    .map((row) => {
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

      return {
        rowId: item.rowId,
        values,
      };
    })
    .filter((row): row is ParsedRow => row !== null);

  const parsed: ParsedFile = {
    fileId: candidate.fileId,
    fileName: candidate.fileName,
    columns,
    rows,
    level1Options: toSafeStringArray(candidate.level1Options),
    level2Options: toSafeStringArray(candidate.level2Options),
  };

  const level1Key = getLevelColumnKey(columns, "level1");
  const level2Key = getLevelColumnKey(columns, "level2");
  if (parsed.level1Options.length === 0) {
    parsed.level1Options = getDistinctOptions(rows, level1Key);
  }
  if (parsed.level2Options.length === 0) {
    parsed.level2Options = getDistinctOptions(rows, level2Key);
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
    : columns.filter((column) => column.editable).map((column) => column.key);
  const filterKeysFromState = Array.isArray(candidate.selectedFilterColumnKeys)
    ? toSafeStringArray(candidate.selectedFilterColumnKeys)
    : undefined;
  const rawFilterValues = toSafeStringRecord(candidate.columnFilterValues);

  const normalized = toViewState(
    parsed,
    displayKeysFromState,
    editableKeysFromState,
    filterKeysFromState,
    rawFilterValues,
  );
  const legacyFilterValues: Record<string, string> = {};
  const legacyLevel1Key = getLevelColumnKey(columns, "level1");
  const legacyLevel2Key = getLevelColumnKey(columns, "level2");
  const legacyTimeKey = columns.find((column) =>
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
  const mergedFilterValues = {
    ...normalized.columnFilterValues,
    ...legacyFilterValues,
  };
  const mergedFilterKeys =
    normalized.selectedFilterColumnKeys.length > 0
      ? normalized.selectedFilterColumnKeys
      : normalizeFilterSelection(columns, Object.keys(legacyFilterValues));
  return {
    ...normalized,
    selectedFilterColumnKeys: mergedFilterKeys,
    columnFilterValues: mergedFilterKeys.reduce<Record<string, string>>(
      (acc, key) => {
        const value = mergedFilterValues[key];
        acc[key] = typeof value === "string" ? value : ALL_FILTER_VALUE;
        return acc;
      },
      {},
    ),
  };
}
