import { useEffect, useMemo, useRef, useState } from "react";
import type {
  AIDetectConfig,
  FileViewState,
  NamedAIDetectConfig,
  ParsedCell,
  ParsedColumn,
  ParsedFile,
  ParsedRow,
} from "./types";

const ALL_FILTER_VALUE = "全部";
const QUALIFIED_TITLE_ALIASES = ["是否合格"] as const;
const TIME_TITLE_ALIASES = ["时间"] as const;
const CREATOR_TITLE_ALIASES = ["创建人"] as const;
const INSPECTOR_TITLE_ALIASES = ["质检员"] as const;
const FEEDBACK_TITLE_ALIASES = ["业务反馈意见", "质检员业务反馈意见"] as const;
const OPENSOURCE_TITLE_ALIASES = ["是否开源"] as const;
const AI_RESULT_WITH_CONFIG_COLUMN_KEY = "__ai_result_with_config__";
const AI_RESULT_WITH_CONFIG_COLUMN_TITLE = "AI解析结果+AI配置名";

function normalizeHeaderTitle(value: string): string {
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

function isQualifiedColumnTitle(columnTitle: string): boolean {
  return matchesHeaderAlias(columnTitle, QUALIFIED_TITLE_ALIASES);
}

function isTimeColumnTitle(columnTitle: string): boolean {
  return matchesHeaderAlias(columnTitle, TIME_TITLE_ALIASES);
}

function isCreatorColumnTitle(columnTitle: string): boolean {
  return matchesHeaderAlias(columnTitle, CREATOR_TITLE_ALIASES);
}

function isInspectorColumnTitle(columnTitle: string): boolean {
  return matchesHeaderAlias(columnTitle, INSPECTOR_TITLE_ALIASES);
}

function isFeedbackColumnTitle(columnTitle: string): boolean {
  return matchesHeaderAlias(columnTitle, FEEDBACK_TITLE_ALIASES);
}

function isOpensourceColumnTitle(columnTitle: string): boolean {
  return matchesHeaderAlias(columnTitle, OPENSOURCE_TITLE_ALIASES);
}

interface ColumnPrefsConfig {
  fieldSignature: string;
  displayKeys: string[];
  editableKeys: string[];
}

const DEFAULT_AI_CONFIG: AIDetectConfig = {
  provider: "openai",
  url: "https://api.openai.com/v1",
  model: "gpt-4.1-mini",
  apiKey: "",
  vertexProject: "",
  vertexLocation: "us-central1",
  submitFieldKeys: [],
  prompt:
    "你是一个质检助手。请根据输入字段给出回答结论、问题点和建议。\n输出要求：\n1. 先给出结论（合格/不合格）\n2. 再列出具体问题\n3. 最后给出修改建议\n\n字段内容如下：\n{{fields_json}}",
  resultFieldKey: "",
  reasoningEffort: "high",
  retryCount: 2,
};
const DEFAULT_AI_CONFIG_NAME = "默认配置";
const AI_REASONING_EFFORT_OPTIONS = ["low", "medium", "high"] as const;
const AI_PROVIDER_OPTIONS = [
  { value: "openai", label: "OpenAI兼容接口" },
  { value: "vertex", label: "Google Vertex 原生" },
] as const;
const DEFAULT_AI_RETRY_COUNT = 2;
const MIN_AI_RETRY_COUNT = 0;
const MAX_AI_RETRY_COUNT = 10;
const DEFAULT_AI_BATCH_CONCURRENCY = 4;
const MIN_AI_BATCH_CONCURRENCY = 1;
const MAX_AI_BATCH_CONCURRENCY = 32;

interface AIDetectFieldPayload {
  title: string;
  type: "text" | "image";
  value: string;
  imageUrl?: string;
  imageUrls?: string[];
}

interface AIDetectStreamResult {
  answerText: string;
  thinkingText: string;
}

type AIBatchTaskStatus = "idle" | "running" | "completed";

interface AIBatchTaskState {
  status: AIBatchTaskStatus;
  fileId: string | null;
  fileName: string;
  total: number;
  completed: number;
  success: number;
  failed: number;
  message: string;
}

const INITIAL_AI_BATCH_TASK: AIBatchTaskState = {
  status: "idle",
  fileId: null,
  fileName: "",
  total: 0,
  completed: 0,
  success: 0,
  failed: 0,
  message: "",
};

function getAIBatchTaskStatusText(task: AIBatchTaskState): string {
  if (task.status === "running") {
    return "运行中";
  }
  if (task.status === "completed") {
    return task.failed > 0 ? "已完成（含失败）" : "已完成";
  }
  return "未启动";
}

function formatDuration(ms: number): string {
  const totalSeconds = Math.max(0, Math.floor(ms / 1000));
  const minutes = Math.floor(totalSeconds / 60)
    .toString()
    .padStart(2, "0");
  const seconds = (totalSeconds % 60).toString().padStart(2, "0");
  return `${minutes}:${seconds}`;
}

function normalizeAIBatchConcurrency(value: unknown): number {
  if (typeof value !== "number" || !Number.isFinite(value)) {
    return DEFAULT_AI_BATCH_CONCURRENCY;
  }
  const rounded = Math.floor(value);
  if (rounded < MIN_AI_BATCH_CONCURRENCY) {
    return MIN_AI_BATCH_CONCURRENCY;
  }
  if (rounded > MAX_AI_BATCH_CONCURRENCY) {
    return MAX_AI_BATCH_CONCURRENCY;
  }
  return rounded;
}

function normalizeAIRetryCount(value: unknown): number {
  if (typeof value !== "number" || !Number.isInteger(value)) {
    return DEFAULT_AI_RETRY_COUNT;
  }
  if (value < MIN_AI_RETRY_COUNT) {
    return MIN_AI_RETRY_COUNT;
  }
  if (value > MAX_AI_RETRY_COUNT) {
    return MAX_AI_RETRY_COUNT;
  }
  return value;
}

function composeAISaveText(answerText: string, thinkingText: string): string {
  const answer = answerText.trim();
  const thinking = thinkingText.trim();

  if (answer.length === 0 && thinking.length === 0) {
    return "";
  }
  if (thinking.length === 0) {
    return answerText;
  }
  if (answer.length === 0) {
    return `【思考过程】\n${thinkingText}`;
  }
  return `【思考过程】\n${thinkingText}\n\n【AI结果】\n${answerText}`;
}

function composeAISaveTextWithConfigName(
  answerText: string,
  thinkingText: string,
  configName: string,
): string {
  const content = composeAISaveText(answerText, thinkingText).trim();
  if (content.length === 0) {
    return "";
  }
  const normalizedConfigName =
    configName.trim().length > 0 ? configName.trim() : DEFAULT_AI_CONFIG_NAME;
  return `【AI配置】${normalizedConfigName}\n${content}`;
}

function getCellImageSources(cell: ParsedCell | undefined): string[] {
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

function logUIImageRenderError(
  rowId: string,
  columnTitle: string,
  src: string,
): void {
  // eslint-disable-next-line no-console
  console.log(
    `[UIImageRenderError] row=${rowId} column=${columnTitle} src=${src}`,
  );
}

function getFileNameFromDisposition(disposition: string | null): string | null {
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

function downloadBlob(blob: Blob, fileName: string): void {
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

const AUTO_BLOCK_FORMULA_LENGTH = 24;

function getPureInlineFormulaBody(value: string): string | null {
  const trimmed = value.trim();

  if (trimmed.startsWith("$$") && trimmed.endsWith("$$")) {
    return null;
  }
  if (trimmed.startsWith("\\[") && trimmed.endsWith("\\]")) {
    return null;
  }

  if (trimmed.startsWith("$") && trimmed.endsWith("$")) {
    const body = trimmed.slice(1, -1);
    if (!body.includes("$")) {
      return body.trim();
    }
  }

  if (trimmed.startsWith("\\(") && trimmed.endsWith("\\)")) {
    return trimmed.slice(2, -2).trim();
  }

  return null;
}

function shouldAutoDisplayLatex(value: string): boolean {
  const formulaBody = getPureInlineFormulaBody(value);
  return (
    formulaBody !== null && formulaBody.length >= AUTO_BLOCK_FORMULA_LENGTH
  );
}

function toDisplayMathIfNeeded(value: string, forceDisplay: boolean): string {
  if (!forceDisplay) {
    return value;
  }
  const formulaBody = getPureInlineFormulaBody(value);
  if (formulaBody === null) {
    return value;
  }
  return `\\[${formulaBody}\\]`;
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

function getAllColumnKeys(columns: ParsedColumn[]): string[] {
  return columns.map((column) => column.key);
}

function isAIResultWithConfigColumn(column: ParsedColumn): boolean {
  if (column.key === AI_RESULT_WITH_CONFIG_COLUMN_KEY) {
    return true;
  }
  return (
    normalizeHeaderTitle(column.title) ===
    normalizeHeaderTitle(AI_RESULT_WITH_CONFIG_COLUMN_TITLE)
  );
}

function getAIResultWithConfigColumn(
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

function isFilterColumnTitle(columnTitle: string): boolean {
  const normalized = normalizeHeaderTitle(columnTitle);
  return normalized === "level1" || normalized === "level2";
}

function getFieldSignature(columns: ParsedColumn[]): string {
  return columns
    .filter((column) => !isAIResultWithConfigColumn(column))
    .map((column) => normalizeHeaderTitle(column.title))
    .join("|");
}

function normalizeColumnSelection(
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

function toViewState(
  parsed: ParsedFile,
  selectedDisplayColumnKeys?: string[],
  selectedEditableColumnKeys?: string[],
): FileViewState {
  const nextParsed = ensureAIResultWithConfigColumn(parsed);
  const normalized = normalizeColumnSelection(
    nextParsed.columns,
    selectedDisplayColumnKeys,
    selectedEditableColumnKeys,
  );
  return {
    ...nextParsed,
    columns: applyEditableConfig(nextParsed.columns, normalized.editableKeys),
    selectedDisplayColumnKeys: normalized.displayKeys,
    selectedEditableColumnKeys: normalized.editableKeys,
    level1Filter: ALL_FILTER_VALUE,
    level2Filter: ALL_FILTER_VALUE,
    timeFilter: ALL_FILTER_VALUE,
  };
}

function applyColumnConfigToFile(
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
  };
}

function toSafeStringArray(value: unknown): string[] {
  if (!Array.isArray(value)) {
    return [];
  }
  return value.filter((item): item is string => typeof item === "string");
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

function normalizeLoadedFileState(value: unknown): FileViewState | null {
  if (!value || typeof value !== "object") {
    return null;
  }

  const candidate = value as Partial<FileViewState> & {
    selectedOptionalColumnKeys?: unknown;
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

  const normalized = toViewState(
    parsed,
    displayKeysFromState,
    editableKeysFromState,
  );
  return {
    ...normalized,
    level1Filter:
      typeof candidate.level1Filter === "string"
        ? candidate.level1Filter
        : ALL_FILTER_VALUE,
    level2Filter:
      typeof candidate.level2Filter === "string"
        ? candidate.level2Filter
        : ALL_FILTER_VALUE,
    timeFilter:
      typeof candidate.timeFilter === "string"
        ? candidate.timeFilter
        : ALL_FILTER_VALUE,
  };
}

function normalizeLoadedAIDetectConfig(value: unknown): AIDetectConfig {
  if (!value || typeof value !== "object") {
    return { ...DEFAULT_AI_CONFIG };
  }

  const candidate = value as Partial<AIDetectConfig>;
  const provider =
    candidate.provider === "openai" || candidate.provider === "vertex"
      ? candidate.provider
      : DEFAULT_AI_CONFIG.provider;
  const submitFieldKeys = Array.isArray(candidate.submitFieldKeys)
    ? candidate.submitFieldKeys.filter(
        (item): item is string => typeof item === "string",
      )
    : [];
  const reasoningEffort =
    candidate.reasoningEffort === "low" ||
    candidate.reasoningEffort === "medium" ||
    candidate.reasoningEffort === "high"
      ? candidate.reasoningEffort
      : DEFAULT_AI_CONFIG.reasoningEffort;
  const retryCount = normalizeAIRetryCount(candidate.retryCount);

  return {
    provider,
    url:
      typeof candidate.url === "string" && candidate.url.trim().length > 0
        ? candidate.url
        : DEFAULT_AI_CONFIG.url,
    model:
      typeof candidate.model === "string" && candidate.model.trim().length > 0
        ? candidate.model
        : DEFAULT_AI_CONFIG.model,
    apiKey: typeof candidate.apiKey === "string" ? candidate.apiKey : "",
    vertexProject:
      typeof candidate.vertexProject === "string"
        ? candidate.vertexProject
        : "",
    vertexLocation:
      typeof candidate.vertexLocation === "string" &&
      candidate.vertexLocation.trim().length > 0
        ? candidate.vertexLocation
        : DEFAULT_AI_CONFIG.vertexLocation,
    submitFieldKeys,
    prompt:
      typeof candidate.prompt === "string" && candidate.prompt.trim().length > 0
        ? candidate.prompt
        : DEFAULT_AI_CONFIG.prompt,
    resultFieldKey:
      typeof candidate.resultFieldKey === "string"
        ? candidate.resultFieldKey
        : "",
    reasoningEffort,
    retryCount,
  };
}

function normalizeAIConfigName(value: unknown): string {
  if (typeof value !== "string") {
    return DEFAULT_AI_CONFIG_NAME;
  }
  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : DEFAULT_AI_CONFIG_NAME;
}

function normalizeLoadedNamedAIDetectConfigs(
  value: unknown,
): NamedAIDetectConfig[] {
  if (!Array.isArray(value)) {
    return [];
  }

  const usedNames = new Set<string>();
  const result: NamedAIDetectConfig[] = [];

  value.forEach((item) => {
    if (!item || typeof item !== "object") {
      return;
    }
    const candidate = item as {
      name?: unknown;
      config?: unknown;
    };
    const name =
      typeof candidate.name === "string" ? candidate.name.trim() : "";
    if (name.length === 0 || usedNames.has(name)) {
      return;
    }

    const config = normalizeLoadedAIDetectConfig(
      candidate.config && typeof candidate.config === "object"
        ? candidate.config
        : item,
    );
    usedNames.add(name);
    result.push({
      name,
      config,
    });
  });

  return result;
}

function normalizeNamedAIDetectConfigsForColumns(
  configs: NamedAIDetectConfig[],
  columns: ParsedColumn[],
): NamedAIDetectConfig[] {
  return configs.map((item) => ({
    name: item.name,
    config: normalizeAIDetectConfigForColumns(item.config, columns),
  }));
}

function pickAIConfigName(
  configs: NamedAIDetectConfig[],
  preferredName: unknown,
): string {
  if (configs.length === 0) {
    return DEFAULT_AI_CONFIG_NAME;
  }
  if (typeof preferredName === "string") {
    const trimmed = preferredName.trim();
    if (trimmed.length > 0 && configs.some((item) => item.name === trimmed)) {
      return trimmed;
    }
  }
  return configs[0].name;
}

function getDefaultResultFieldKey(columns: ParsedColumn[]): string {
  const aiResultWithConfigColumn = getAIResultWithConfigColumn(columns);
  if (aiResultWithConfigColumn) {
    return aiResultWithConfigColumn.key;
  }
  const feedbackEditable = columns.find(
    (column) => column.editable && isFeedbackColumnTitle(column.title),
  );
  if (feedbackEditable) {
    return feedbackEditable.key;
  }
  const firstEditable = columns.find((column) => column.editable);
  return firstEditable?.key ?? "";
}

function isLegacyAIDetectResultColumnTitle(columnTitle: string): boolean {
  return (
    isOpensourceColumnTitle(columnTitle) ||
    isQualifiedColumnTitle(columnTitle) ||
    isInspectorColumnTitle(columnTitle) ||
    isFeedbackColumnTitle(columnTitle)
  );
}

function normalizeAIDetectConfigForColumns(
  config: AIDetectConfig,
  columns: ParsedColumn[],
): AIDetectConfig {
  const keySet = new Set(columns.map((column) => column.key));
  const submitFieldKeys = config.submitFieldKeys.filter((key) =>
    keySet.has(key),
  );
  const editableKeySet = new Set(
    columns.filter((column) => column.editable).map((column) => column.key),
  );
  const aiResultWithConfigColumn = getAIResultWithConfigColumn(columns);
  const currentResultColumn = columns.find(
    (column) => column.key === config.resultFieldKey,
  );
  const shouldMigrateLegacyResultField =
    aiResultWithConfigColumn !== null &&
    currentResultColumn !== undefined &&
    isLegacyAIDetectResultColumnTitle(currentResultColumn.title);
  const nextResultFieldKey = shouldMigrateLegacyResultField
    ? aiResultWithConfigColumn.key
    : editableKeySet.has(config.resultFieldKey)
      ? config.resultFieldKey
      : getDefaultResultFieldKey(columns);

  return {
    ...config,
    submitFieldKeys,
    resultFieldKey: nextResultFieldKey,
    retryCount: normalizeAIRetryCount(config.retryCount),
  };
}

function buildAIDetectFieldsForRow(
  columns: ParsedColumn[],
  row: ParsedRow,
  submitFieldKeys: string[],
): AIDetectFieldPayload[] {
  const fieldMap = new Map(columns.map((column) => [column.key, column]));
  const fields: AIDetectFieldPayload[] = [];

  submitFieldKeys.forEach((key) => {
    const column = fieldMap.get(key);
    if (!column) {
      return;
    }

    const cell = row.values[key];
    const imageSources = getCellImageSources(cell);
    if (cell?.type === "image" && imageSources.length > 0) {
      fields.push({
        title: column.title,
        type: "image",
        value: cell.value ?? "",
        imageUrl: imageSources[0],
        imageUrls: imageSources,
      });
      return;
    }

    fields.push({
      title: column.title,
      type: "text",
      value: cell?.value ?? "",
    });
  });

  return fields;
}

async function requestAIDetectResult(
  payload: {
    provider: AIDetectConfig["provider"];
    url: string;
    model: string;
    apiKey: string;
    vertexProject: string;
    vertexLocation: string;
    prompt: string;
    fields: AIDetectFieldPayload[];
    reasoningEffort: AIDetectConfig["reasoningEffort"];
    retryCount: number;
  },
  options?: {
    signal?: AbortSignal;
    onAnswerChunk?: (chunk: string) => void;
    onThinkingChunk?: (chunk: string) => void;
    onChunk?: (chunk: string) => void;
  },
): Promise<AIDetectStreamResult> {
  const response = await fetch("/api/ai-detect/stream", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    signal: options?.signal,
    body: JSON.stringify(payload),
  });

  if (!response.ok) {
    const payload = (await response.json().catch(() => ({}))) as {
      message?: string;
    };
    throw new Error(payload.message ?? "AI 回答失败");
  }
  if (!response.body) {
    throw new Error("AI 响应流为空");
  }

  const reader = response.body.getReader();
  const decoder = new TextDecoder("utf-8");
  const contentType = response.headers.get("content-type")?.toLowerCase() ?? "";
  let answerText = "";
  let thinkingText = "";

  if (contentType.includes("application/x-ndjson")) {
    let buffer = "";
    while (true) {
      const { value, done } = await reader.read();
      if (done) {
        buffer += decoder.decode();
        break;
      }
      if (!value) {
        continue;
      }
      buffer += decoder.decode(value, { stream: true });

      const lines = buffer.split(/\r?\n/);
      buffer = lines.pop() ?? "";

      for (const rawLine of lines) {
        const line = rawLine.trim();
        if (!line) {
          continue;
        }

        try {
          const event = JSON.parse(line) as {
            type?: string;
            text?: string;
          };
          if (event.type === "answer" && typeof event.text === "string") {
            answerText += event.text;
            options?.onAnswerChunk?.(event.text);
            options?.onChunk?.(event.text);
            continue;
          }
          if (event.type === "thinking" && typeof event.text === "string") {
            thinkingText += event.text;
            options?.onThinkingChunk?.(event.text);
            continue;
          }
          if (event.type === "done") {
            continue;
          }
        } catch {
          // Compatibility fallback: treat unknown lines as answer text.
          answerText += rawLine;
          options?.onAnswerChunk?.(rawLine);
          options?.onChunk?.(rawLine);
        }
      }
    }

    const rest = buffer.trim();
    if (rest.length > 0) {
      try {
        const event = JSON.parse(rest) as {
          type?: string;
          text?: string;
        };
        if (event.type === "answer" && typeof event.text === "string") {
          answerText += event.text;
          options?.onAnswerChunk?.(event.text);
          options?.onChunk?.(event.text);
        } else if (
          event.type === "thinking" &&
          typeof event.text === "string"
        ) {
          thinkingText += event.text;
          options?.onThinkingChunk?.(event.text);
        }
      } catch {
        answerText += rest;
        options?.onAnswerChunk?.(rest);
        options?.onChunk?.(rest);
      }
    }
  } else {
    while (true) {
      const { value, done } = await reader.read();
      if (done) {
        const flushText = decoder.decode();
        if (flushText.length > 0) {
          answerText += flushText;
          options?.onAnswerChunk?.(flushText);
          options?.onChunk?.(flushText);
        }
        break;
      }
      if (!value) {
        continue;
      }
      const chunkText = decoder.decode(value, { stream: true });
      if (chunkText.length > 0) {
        answerText += chunkText;
        options?.onAnswerChunk?.(chunkText);
        options?.onChunk?.(chunkText);
      }
    }
  }

  return {
    answerText,
    thinkingText,
  };
}

function getLevelColumnKey(
  columns: ParsedColumn[],
  title: string,
): string | undefined {
  return columns.find((column) => normalizeHeaderTitle(column.title) === title)
    ?.key;
}

function getCellText(row: ParsedRow, columnKey: string): string {
  return row.values[columnKey]?.value ?? "";
}

const LATEX_PATTERN =
  /(\$\$[\s\S]+?\$\$)|((^|[^\\])\$[^$\n]+?\$)|(\\\([\s\S]+?\\\))|(\\\[[\s\S]+?\\\])|(\\(?:frac|sqrt|sum|int|left|right|begin|end|alpha|beta|gamma|delta|theta|lambda|pi|times|cdot|pm|leq|geq|neq|ce|pu)\b)/;

function hasLatexSyntax(value: string): boolean {
  return LATEX_PATTERN.test(value.trim());
}

function hasMathDelimiter(value: string): boolean {
  const trimmed = value.trim();
  return (
    (trimmed.startsWith("$$") && trimmed.endsWith("$$")) ||
    (trimmed.startsWith("\\[") && trimmed.endsWith("\\]")) ||
    (trimmed.startsWith("$") && trimmed.endsWith("$")) ||
    (trimmed.startsWith("\\(") && trimmed.endsWith("\\)"))
  );
}

function toMathJaxSource(value: string, forceDisplay: boolean): string {
  const normalized = toDisplayMathIfNeeded(value, forceDisplay);
  if (hasMathDelimiter(normalized)) {
    return normalized;
  }

  const trimmed = normalized.trim();
  if (trimmed.length > 0 && hasLatexSyntax(trimmed)) {
    return `\\(${trimmed}\\)`;
  }

  return normalized;
}

type MathJaxConfig = {
  loader?: {
    load?: string[];
  };
  tex?: {
    inlineMath?: Array<[string, string]>;
    displayMath?: Array<[string, string]>;
    packages?: Record<string, string[]>;
  };
  options?: {
    skipHtmlTags?: string[];
  };
  svg?: {
    fontCache?: string;
  };
  startup?: {
    promise?: Promise<unknown>;
  };
  typesetPromise?: (elements?: Element[]) => Promise<unknown>;
};

declare global {
  interface Window {
    MathJax?: MathJaxConfig;
  }
}

let mathJaxLoadPromise: Promise<void> | null = null;

async function ensureMathJaxLoaded(): Promise<MathJaxConfig | null> {
  if (typeof window === "undefined") {
    return null;
  }

  if (window.MathJax?.typesetPromise) {
    return window.MathJax;
  }

  if (!mathJaxLoadPromise) {
    mathJaxLoadPromise = new Promise<void>((resolve, reject) => {
      const existingScript = document.querySelector<HTMLScriptElement>(
        "script[data-mathjax-loader='true']",
      );

      if (existingScript) {
        if (window.MathJax?.typesetPromise) {
          resolve();
          return;
        }
        existingScript.addEventListener("load", () => resolve(), {
          once: true,
        });
        existingScript.addEventListener(
          "error",
          () => reject(new Error("MathJax 脚本加载失败")),
          { once: true },
        );
        return;
      }

      window.MathJax = window.MathJax ?? {};
      const currentLoader = window.MathJax.loader ?? {};
      const currentLoaderLoads = currentLoader.load ?? [];
      window.MathJax.loader = {
        ...currentLoader,
        load: Array.from(new Set([...currentLoaderLoads, "[tex]/mhchem"])),
      };

      const currentTex = window.MathJax.tex ?? {};
      const currentPackages = currentTex.packages ?? {};
      const extraPackages = currentPackages["[+]"] ?? [];
      window.MathJax.tex = {
        ...currentTex,
        inlineMath: currentTex.inlineMath ?? [
          ["$", "$"],
          ["\\(", "\\)"],
        ],
        displayMath: currentTex.displayMath ?? [
          ["$$", "$$"],
          ["\\[", "\\]"],
        ],
        packages: {
          ...currentPackages,
          "[+]": Array.from(new Set([...extraPackages, "mhchem"])),
        },
      };
      window.MathJax.options = window.MathJax.options ?? {
        skipHtmlTags: [
          "script",
          "noscript",
          "style",
          "textarea",
          "pre",
          "code",
        ],
      };
      window.MathJax.svg = window.MathJax.svg ?? {
        fontCache: "global",
      };

      const script = document.createElement("script");
      script.src = "https://cdn.jsdelivr.net/npm/mathjax@3/es5/tex-svg.js";
      script.async = true;
      script.defer = true;
      script.setAttribute("data-mathjax-loader", "true");
      script.onload = () => resolve();
      script.onerror = () => reject(new Error("MathJax 脚本加载失败"));
      document.head.appendChild(script);
    }).catch((error) => {
      mathJaxLoadPromise = null;
      throw error;
    });
  }

  await mathJaxLoadPromise;
  return window.MathJax ?? null;
}

function LatexRenderer({
  value,
  forceDisplay = false,
}: {
  value: string;
  forceDisplay?: boolean;
}) {
  const containerRef = useRef<HTMLDivElement>(null);
  const [renderFailed, setRenderFailed] = useState(false);

  useEffect(() => {
    let disposed = false;

    const render = async () => {
      const container = containerRef.current;
      if (!container) {
        return;
      }

      container.textContent = toMathJaxSource(value, forceDisplay);
      setRenderFailed(false);

      try {
        const mathJax = await ensureMathJaxLoaded();
        if (disposed || !mathJax?.typesetPromise || !containerRef.current) {
          return;
        }

        if (mathJax.startup?.promise) {
          await mathJax.startup.promise;
        }
        await mathJax.typesetPromise([containerRef.current]);
      } catch {
        if (!disposed) {
          setRenderFailed(true);
        }
      }
    };

    void render();
    return () => {
      disposed = true;
    };
  }, [value, forceDisplay]);

  if (renderFailed) {
    return <div className="latex-plain">{value}</div>;
  }

  return (
    <div
      className={`latex-rendered ${forceDisplay ? "latex-rendered-display" : "latex-rendered-inline"}`}
      ref={containerRef}
    />
  );
}

/* ─── SVG Icons ─── */
const IconUpload = () => (
  <svg
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
  >
    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
    <polyline points="17 8 12 3 7 8" />
    <line x1="12" y1="3" x2="12" y2="15" />
  </svg>
);

const IconDownload = () => (
  <svg
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
  >
    <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4" />
    <polyline points="7 10 12 15 17 10" />
    <line x1="12" y1="15" x2="12" y2="3" />
  </svg>
);

const IconFile = () => (
  <svg
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
  >
    <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />
    <polyline points="14 2 14 8 20 8" />
  </svg>
);

const IconChevron = () => (
  <svg
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
  >
    <polyline points="9 18 15 12 9 6" />
  </svg>
);

const IconSun = () => (
  <svg
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
  >
    <circle cx="12" cy="12" r="5" />
    <line x1="12" y1="1" x2="12" y2="3" />
    <line x1="12" y1="21" x2="12" y2="23" />
    <line x1="4.22" y1="4.22" x2="5.64" y2="5.64" />
    <line x1="18.36" y1="18.36" x2="19.78" y2="19.78" />
    <line x1="1" y1="12" x2="3" y2="12" />
    <line x1="21" y1="12" x2="23" y2="12" />
    <line x1="4.22" y1="19.78" x2="5.64" y2="18.36" />
    <line x1="18.36" y1="5.64" x2="19.78" y2="4.22" />
  </svg>
);

const IconMoon = () => (
  <svg
    viewBox="0 0 24 24"
    fill="none"
    stroke="currentColor"
    strokeWidth="2"
    strokeLinecap="round"
    strokeLinejoin="round"
  >
    <path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z" />
  </svg>
);

function App() {
  type PendingConfigMode = "import" | "edit";
  const [files, setFiles] = useState<FileViewState[]>([]);
  const [activeFileId, setActiveFileId] = useState<string | null>(null);
  const [isUploading, setIsUploading] = useState(false);
  const [isExporting, setIsExporting] = useState(false);
  const [errorMessage, setErrorMessage] = useState<string>("");
  const [isAIConfigModalOpen, setIsAIConfigModalOpen] = useState(false);
  const [aiConfigLoading, setAIConfigLoading] = useState(false);
  const [aiConfigSaving, setAIConfigSaving] = useState(false);
  const [aiConfigList, setAIConfigList] = useState<NamedAIDetectConfig[]>([
    {
      name: DEFAULT_AI_CONFIG_NAME,
      config: { ...DEFAULT_AI_CONFIG },
    },
  ]);
  const [selectedAIConfigName, setSelectedAIConfigName] = useState<string>(
    DEFAULT_AI_CONFIG_NAME,
  );
  const [draftAIConfigName, setDraftAIConfigName] = useState<string>(
    DEFAULT_AI_CONFIG_NAME,
  );
  const [aiConfig, setAIConfig] = useState<AIDetectConfig>({
    ...DEFAULT_AI_CONFIG,
  });
  const [draftAIConfig, setDraftAIConfig] = useState<AIDetectConfig>({
    ...DEFAULT_AI_CONFIG,
  });
  const [aiConfigFormMessage, setAIConfigFormMessage] = useState("");
  const [isAIDetecting, setIsAIDetecting] = useState(false);
  const [aiThinkingText, setAIThinkingText] = useState("");
  const [aiResultText, setAIResultText] = useState("");
  const [aiResultConfigName, setAIResultConfigName] = useState<string>(
    DEFAULT_AI_CONFIG_NAME,
  );
  const [aiResultMessage, setAIResultMessage] = useState("");
  const [aiDetectElapsedMs, setAIDetectElapsedMs] = useState(0);
  const [isSavingAIResult, setIsSavingAIResult] = useState(false);
  const [aiBatchTask, setAIBatchTask] = useState<AIBatchTaskState>(
    INITIAL_AI_BATCH_TASK,
  );
  const [aiBatchConcurrency, setAIBatchConcurrency] = useState<number>(
    DEFAULT_AI_BATCH_CONCURRENCY,
  );
  const [selectedRowId, setSelectedRowId] = useState<string | null>(null);
  const [batchSelectedRowIds, setBatchSelectedRowIds] = useState<string[]>([]);
  const [pendingFile, setPendingFile] = useState<ParsedFile | null>(null);
  const [pendingSelectedDisplayKeys, setPendingSelectedDisplayKeys] = useState<
    string[]
  >([]);
  const [pendingEditableColumnKeys, setPendingEditableColumnKeys] = useState<
    string[]
  >([]);
  const [pendingConfigNotice, setPendingConfigNotice] = useState<string>("");
  const [pendingConfigMode, setPendingConfigMode] =
    useState<PendingConfigMode>("import");
  const [showHiddenFields, setShowHiddenFields] = useState(false);
  const [latexRenderOverrides, setLatexRenderOverrides] = useState<
    Record<string, boolean>
  >({});
  const [previewImageSrc, setPreviewImageSrc] = useState<string | null>(null);
  const [theme, setTheme] = useState<"dark" | "light">(() => {
    if (typeof window !== "undefined") {
      return (localStorage.getItem("theme") as "dark" | "light") || "dark";
    }
    return "dark";
  });
  const uploadInputRef = useRef<HTMLInputElement>(null);
  const persistTimersRef = useRef<Record<string, number>>({});
  const pendingPersistRef = useRef<Record<string, FileViewState>>({});
  const aiStreamAbortRef = useRef<AbortController | null>(null);
  const aiBatchAbortRef = useRef<AbortController | null>(null);
  const aiDetectStartedAtRef = useRef<number | null>(null);

  useEffect(() => {
    document.documentElement.setAttribute("data-theme", theme);
    localStorage.setItem("theme", theme);
  }, [theme]);

  useEffect(() => {
    let disposed = false;
    const controller = new AbortController();

    const loadPersistedFiles = async () => {
      try {
        const response = await fetch("/api/files", {
          signal: controller.signal,
        });
        if (!response.ok) {
          return;
        }

        const payload = (await response.json()) as { files?: unknown };
        const rawFiles = Array.isArray(payload.files) ? payload.files : [];
        const restoredFiles = rawFiles
          .map((item) => normalizeLoadedFileState(item))
          .filter((item): item is FileViewState => item !== null);

        if (disposed || restoredFiles.length === 0) {
          return;
        }

        setFiles((previous) => {
          if (previous.length === 0) {
            return restoredFiles;
          }

          const merged = new Map<string, FileViewState>();
          restoredFiles.forEach((file) => merged.set(file.fileId, file));
          previous.forEach((file) => merged.set(file.fileId, file));
          return Array.from(merged.values());
        });
        setActiveFileId((previous) => previous ?? restoredFiles[0].fileId);
      } catch {
        // Ignore load errors and keep empty startup state.
      }
    };

    void loadPersistedFiles();

    return () => {
      disposed = true;
      controller.abort();
    };
  }, []);

  useEffect(() => {
    return () => {
      aiStreamAbortRef.current?.abort();
      aiStreamAbortRef.current = null;
      aiBatchAbortRef.current?.abort();
      aiBatchAbortRef.current = null;
      Object.values(persistTimersRef.current).forEach((timerId) => {
        window.clearTimeout(timerId);
      });
      persistTimersRef.current = {};
      pendingPersistRef.current = {};
    };
  }, []);

  const toggleTheme = () => {
    setTheme((prev) => (prev === "dark" ? "light" : "dark"));
  };

  const activeFile = useMemo(
    () => files.find((item) => item.fileId === activeFileId) ?? null,
    [files, activeFileId],
  );

  useEffect(() => {
    if (!activeFile) {
      const nextConfig = { ...DEFAULT_AI_CONFIG };
      setAIConfigList([
        {
          name: DEFAULT_AI_CONFIG_NAME,
          config: nextConfig,
        },
      ]);
      setSelectedAIConfigName(DEFAULT_AI_CONFIG_NAME);
      setDraftAIConfigName(DEFAULT_AI_CONFIG_NAME);
      setAIConfig(nextConfig);
      setDraftAIConfig(nextConfig);
      setAIConfigFormMessage("");
      setAIThinkingText("");
      setAIResultText("");
      setAIResultConfigName(DEFAULT_AI_CONFIG_NAME);
      setAIResultMessage("");
      setBatchSelectedRowIds([]);
      aiDetectStartedAtRef.current = null;
      setAIDetectElapsedMs(0);
      setAIConfigLoading(false);
      setIsAIConfigModalOpen(false);
      return;
    }

    let disposed = false;
    const controller = new AbortController();
    setAIConfigLoading(true);

    const loadAIDetectConfig = async () => {
      try {
        const response = await fetch(
          `/api/ai-config/${encodeURIComponent(activeFile.fileName)}`,
          { signal: controller.signal },
        );
        if (!response.ok) {
          throw new Error("加载 AI 配置失败");
        }

        const payload = (await response.json()) as {
          configs?: unknown;
          activeConfigName?: unknown;
          config?: unknown;
        };
        let loadedConfigs = normalizeLoadedNamedAIDetectConfigs(
          payload.configs,
        );
        if (loadedConfigs.length === 0) {
          loadedConfigs = [
            {
              name: DEFAULT_AI_CONFIG_NAME,
              config: normalizeLoadedAIDetectConfig(payload.config),
            },
          ];
        }

        const normalizedConfigs = normalizeNamedAIDetectConfigsForColumns(
          loadedConfigs,
          activeFile.columns,
        );
        const activeConfigName = pickAIConfigName(
          normalizedConfigs,
          payload.activeConfigName,
        );
        const activeConfig =
          normalizedConfigs.find((item) => item.name === activeConfigName)
            ?.config ?? normalizedConfigs[0].config;

        if (disposed) {
          return;
        }
        setAIConfigList(normalizedConfigs);
        setSelectedAIConfigName(activeConfigName);
        setDraftAIConfigName(activeConfigName);
        setAIConfig(activeConfig);
        setDraftAIConfig(activeConfig);
        setAIResultConfigName(activeConfigName);
        setAIConfigFormMessage("");
      } catch {
        if (disposed) {
          return;
        }
        const fallbackConfig = normalizeAIDetectConfigForColumns(
          { ...DEFAULT_AI_CONFIG },
          activeFile.columns,
        );
        setAIConfigList([
          {
            name: DEFAULT_AI_CONFIG_NAME,
            config: fallbackConfig,
          },
        ]);
        setSelectedAIConfigName(DEFAULT_AI_CONFIG_NAME);
        setDraftAIConfigName(DEFAULT_AI_CONFIG_NAME);
        setAIConfig(fallbackConfig);
        setDraftAIConfig(fallbackConfig);
        setAIResultConfigName(DEFAULT_AI_CONFIG_NAME);
        setAIConfigFormMessage("");
      } finally {
        if (!disposed) {
          setAIConfigLoading(false);
        }
      }
    };

    void loadAIDetectConfig();

    return () => {
      disposed = true;
      controller.abort();
    };
  }, [activeFile?.fileId, activeFile?.fileName]);

  useEffect(() => {
    if (!activeFile) {
      return;
    }
    const normalizedConfigs = normalizeNamedAIDetectConfigsForColumns(
      aiConfigList,
      activeFile.columns,
    );
    const nextConfigList =
      normalizedConfigs.length > 0
        ? normalizedConfigs
        : [
            {
              name: DEFAULT_AI_CONFIG_NAME,
              config: normalizeAIDetectConfigForColumns(
                { ...DEFAULT_AI_CONFIG },
                activeFile.columns,
              ),
            },
          ];
    const nextSelectedName = pickAIConfigName(
      nextConfigList,
      selectedAIConfigName,
    );
    const nextSelectedConfig =
      nextConfigList.find((item) => item.name === nextSelectedName)?.config ??
      nextConfigList[0].config;

    setAIConfigList(nextConfigList);
    setSelectedAIConfigName(nextSelectedName);
    setAIConfig(nextSelectedConfig);
    setDraftAIConfigName((previous) =>
      nextConfigList.some((item) => item.name === previous)
        ? previous
        : nextSelectedName,
    );
    setDraftAIConfig((previous) =>
      normalizeAIDetectConfigForColumns(previous, activeFile.columns),
    );
  }, [activeFile?.fileId, activeFile?.columns]);

  useEffect(() => {
    if (!isAIDetecting) {
      return;
    }
    const startedAt = aiDetectStartedAtRef.current ?? Date.now();
    aiDetectStartedAtRef.current = startedAt;
    const timerId = window.setInterval(() => {
      setAIDetectElapsedMs(Date.now() - startedAt);
    }, 250);
    return () => {
      window.clearInterval(timerId);
    };
  }, [isAIDetecting]);

  useEffect(() => {
    setAIThinkingText("");
    setAIResultText("");
    setAIResultConfigName(selectedAIConfigName);
    setAIResultMessage("");
    aiStreamAbortRef.current?.abort();
    aiStreamAbortRef.current = null;
    aiDetectStartedAtRef.current = null;
    setAIDetectElapsedMs(0);
    setIsAIDetecting(false);
    setIsSavingAIResult(false);
  }, [activeFileId, selectedRowId]);

  const level1ColumnKey = activeFile
    ? getLevelColumnKey(activeFile.columns, "level1")
    : undefined;
  const level2ColumnKey = activeFile
    ? getLevelColumnKey(activeFile.columns, "level2")
    : undefined;
  const timeColumnKey = activeFile
    ? activeFile.columns.find((column) => isTimeColumnTitle(column.title))?.key
    : undefined;
  const timeOptions = useMemo(
    () => getDistinctOptions(activeFile?.rows ?? [], timeColumnKey),
    [activeFile?.rows, timeColumnKey],
  );

  const displayColumns = useMemo(() => {
    if (!activeFile) {
      return [];
    }
    return activeFile.columns.filter((column) => {
      return activeFile.selectedDisplayColumnKeys.includes(column.key);
    });
  }, [activeFile]);

  const hiddenColumns = useMemo(() => {
    if (!activeFile) {
      return [];
    }
    return activeFile.columns.filter((column) => {
      return !activeFile.selectedDisplayColumnKeys.includes(column.key);
    });
  }, [activeFile]);

  const aiSubmitFieldColumns = useMemo(
    () =>
      activeFile
        ? activeFile.columns.filter(
            (column) => !isAIResultWithConfigColumn(column),
          )
        : [],
    [activeFile],
  );

  const aiResultFieldColumns = useMemo(
    () =>
      activeFile
        ? activeFile.columns.filter((column) => column.editable)
        : ([] as ParsedColumn[]),
    [activeFile],
  );

  const aiResultFieldTitle = useMemo(() => {
    const matched = aiResultFieldColumns.find(
      (column) => column.key === aiConfig.resultFieldKey,
    );
    return matched?.title ?? "";
  }, [aiResultFieldColumns, aiConfig.resultFieldKey]);

  const isAIBatchRunning = aiBatchTask.status === "running";
  const aiBatchProgressPercent =
    aiBatchTask.total > 0
      ? Math.round((aiBatchTask.completed / aiBatchTask.total) * 100)
      : 0;
  const aiDetectElapsedText = formatDuration(aiDetectElapsedMs);
  const aiMergedStreamText = useMemo(
    () => composeAISaveText(aiResultText, aiThinkingText),
    [aiResultText, aiThinkingText],
  );
  const hasAISaveContent = aiMergedStreamText.trim().length > 0;
  const batchSelectedRowIdSet = useMemo(
    () => new Set(batchSelectedRowIds),
    [batchSelectedRowIds],
  );

  const visibleRows = useMemo(() => {
    if (!activeFile) {
      return [];
    }

    return activeFile.rows.filter((row) => {
      if (level1ColumnKey && activeFile.level1Filter !== ALL_FILTER_VALUE) {
        const value = getCellText(row, level1ColumnKey).trim();
        if (value !== activeFile.level1Filter) {
          return false;
        }
      }

      if (level2ColumnKey && activeFile.level2Filter !== ALL_FILTER_VALUE) {
        const value = getCellText(row, level2ColumnKey).trim();
        if (value !== activeFile.level2Filter) {
          return false;
        }
      }

      if (timeColumnKey && activeFile.timeFilter !== ALL_FILTER_VALUE) {
        const value = getCellText(row, timeColumnKey).trim();
        if (value !== activeFile.timeFilter) {
          return false;
        }
      }

      return true;
    });
  }, [activeFile, level1ColumnKey, level2ColumnKey, timeColumnKey]);

  useEffect(() => {
    if (!activeFile || visibleRows.length === 0) {
      setSelectedRowId(null);
      return;
    }

    if (
      selectedRowId !== null &&
      !visibleRows.some((row) => row.rowId === selectedRowId)
    ) {
      setSelectedRowId(null);
    }
  }, [activeFile, visibleRows, selectedRowId]);

  const selectedRow = useMemo(
    () => visibleRows.find((row) => row.rowId === selectedRowId) ?? null,
    [visibleRows, selectedRowId],
  );

  useEffect(() => {
    if (!activeFile) {
      setBatchSelectedRowIds([]);
      return;
    }

    const visibleIdSet = new Set(visibleRows.map((row) => row.rowId));
    setBatchSelectedRowIds((previous) =>
      previous.filter((rowId) => visibleIdSet.has(rowId)),
    );
  }, [activeFile?.fileId, visibleRows]);

  const rowPreviewColumns = useMemo(() => {
    if (!activeFile) {
      return [];
    }

    const preferred = activeFile.columns.filter(
      (column) =>
        isFilterColumnTitle(column.title) ||
        isQualifiedColumnTitle(column.title) ||
        isTimeColumnTitle(column.title) ||
        isCreatorColumnTitle(column.title),
    );
    const merged = [...preferred, ...displayColumns];
    const uniqueMap = new Map<string, ParsedColumn>();
    merged.forEach((column) => {
      if (!isFeedbackColumnTitle(column.title)) {
        uniqueMap.set(column.key, column);
      }
    });
    return Array.from(uniqueMap.values()).slice(0, 5);
  }, [activeFile, displayColumns]);

  const getRowPreviewText = (row: ParsedRow): string =>
    rowPreviewColumns
      .map((column) => {
        const value = row.values[column.key]?.value?.trim();
        return `${column.title}: ${value || "-"}`;
      })
      .join(" ｜ ");

  const persistFileState = (file: FileViewState) => {
    fetch(`/api/files/${encodeURIComponent(file.fileId)}/state`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ state: file }),
    }).catch(() => {});
  };

  const cancelScheduledPersist = (fileId: string) => {
    const timerId = persistTimersRef.current[fileId];
    if (timerId !== undefined) {
      window.clearTimeout(timerId);
      delete persistTimersRef.current[fileId];
    }
    delete pendingPersistRef.current[fileId];
  };

  const schedulePersistFileState = (
    file: FileViewState,
    delayMs: number = 400,
  ) => {
    cancelScheduledPersist(file.fileId);
    pendingPersistRef.current[file.fileId] = file;
    const timerId = window.setTimeout(() => {
      const latest = pendingPersistRef.current[file.fileId];
      if (latest) {
        persistFileState(latest);
      }
      delete pendingPersistRef.current[file.fileId];
      delete persistTimersRef.current[file.fileId];
    }, delayMs);
    persistTimersRef.current[file.fileId] = timerId;
  };

  const patchActiveFile = (updater: (file: FileViewState) => FileViewState) => {
    if (!activeFile) {
      return;
    }

    const nextFile = updater(activeFile);
    if (nextFile === activeFile) {
      return;
    }

    setFiles((previous) =>
      previous.map((file) =>
        file.fileId === nextFile.fileId ? nextFile : file,
      ),
    );
    schedulePersistFileState(nextFile);
  };

  const persistColumnPrefs = (file: FileViewState) => {
    fetch(`/api/column-prefs/${encodeURIComponent(file.fileName)}`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        fieldSignature: getFieldSignature(file.columns),
        displayKeys: file.selectedDisplayColumnKeys,
        editableKeys: file.selectedEditableColumnKeys,
      }),
    }).catch(() => {});
  };

  const resetPendingConfigState = () => {
    setPendingFile(null);
    setPendingSelectedDisplayKeys([]);
    setPendingEditableColumnKeys([]);
    setPendingConfigNotice("");
    setPendingConfigMode("import");
  };

  const onOpenActiveFileConfig = () => {
    if (!activeFile) {
      return;
    }
    setPendingFile(activeFile);
    setPendingSelectedDisplayKeys(activeFile.selectedDisplayColumnKeys);
    setPendingEditableColumnKeys(activeFile.selectedEditableColumnKeys);
    setPendingConfigNotice("");
    setPendingConfigMode("edit");
  };

  const syncActiveAIConfigState = (nextConfig: AIDetectConfig) => {
    setAIConfig(nextConfig);
    setAIConfigList((previous) =>
      previous.map((item) =>
        item.name === selectedAIConfigName
          ? {
              ...item,
              config: nextConfig,
            }
          : item,
      ),
    );
  };

  const onSelectAIConfigForRun = (configName: string) => {
    if (!activeFile) {
      return;
    }
    const matched = aiConfigList.find((item) => item.name === configName);
    if (!matched) {
      return;
    }

    const normalized = normalizeAIDetectConfigForColumns(
      matched.config,
      activeFile.columns,
    );
    setSelectedAIConfigName(configName);
    setAIConfig(normalized);
    setAIConfigList((previous) =>
      previous.map((item) =>
        item.name === configName
          ? {
              ...item,
              config: normalized,
            }
          : item,
      ),
    );
    if (!isAIConfigModalOpen) {
      setDraftAIConfigName(configName);
      setDraftAIConfig(normalized);
    }
    setAIConfigFormMessage("");

    fetch(`/api/ai-config/${encodeURIComponent(activeFile.fileName)}/active`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        name: configName,
      }),
    }).catch(() => {});
  };

  const onOpenAIConfigModal = () => {
    if (!activeFile) {
      return;
    }
    setDraftAIConfig(
      normalizeAIDetectConfigForColumns(aiConfig, activeFile.columns),
    );
    setDraftAIConfigName(selectedAIConfigName);
    setAIConfigFormMessage("");
    setIsAIConfigModalOpen(true);
  };

  const onCancelAIConfigModal = () => {
    setDraftAIConfig(aiConfig);
    setDraftAIConfigName(selectedAIConfigName);
    setAIConfigFormMessage("");
    setIsAIConfigModalOpen(false);
  };

  const onToggleDraftAISubmitField = (columnKey: string) => {
    setDraftAIConfig((previous) => {
      const exists = previous.submitFieldKeys.includes(columnKey);
      const submitFieldKeys = exists
        ? previous.submitFieldKeys.filter((key) => key !== columnKey)
        : [...previous.submitFieldKeys, columnKey];
      return {
        ...previous,
        submitFieldKeys,
      };
    });
  };

  const onChangeDraftResultField = (columnKey: string) => {
    setDraftAIConfig((previous) => ({
      ...previous,
      resultFieldKey: columnKey,
    }));
  };

  const onSaveAIConfig = async () => {
    if (!activeFile) {
      return;
    }

    const nextConfigName = normalizeAIConfigName(draftAIConfigName);
    if (draftAIConfigName.trim().length === 0) {
      setAIConfigFormMessage("配置名称不能为空");
      return;
    }

    const nextConfig = normalizeAIDetectConfigForColumns(
      draftAIConfig,
      activeFile.columns,
    );

    if (nextConfig.model.trim().length === 0) {
      setAIConfigFormMessage("模型不能为空");
      return;
    }
    if (nextConfig.provider === "openai") {
      if (nextConfig.url.trim().length === 0) {
        setAIConfigFormMessage("OpenAI 兼容接口 URL 不能为空");
        return;
      }
      if (nextConfig.apiKey.trim().length === 0) {
        setAIConfigFormMessage("OpenAI API Key 不能为空");
        return;
      }
    }
    if (nextConfig.provider === "vertex") {
      if (nextConfig.vertexProject.trim().length === 0) {
        setAIConfigFormMessage("Vertex Project 不能为空");
        return;
      }
      if (nextConfig.vertexLocation.trim().length === 0) {
        setAIConfigFormMessage("Vertex Location 不能为空");
        return;
      }
    }
    if (nextConfig.submitFieldKeys.length === 0) {
      setAIConfigFormMessage("请至少选择一个提交回答字段");
      return;
    }
    if (nextConfig.prompt.trim().length === 0) {
      setAIConfigFormMessage("Prompt 不能为空");
      return;
    }
    if (
      aiResultFieldColumns.length > 0 &&
      nextConfig.resultFieldKey.trim().length === 0
    ) {
      setAIConfigFormMessage("请选择 AI 结果保存字段");
      return;
    }

    setAIConfigSaving(true);
    setAIConfigFormMessage("");
    setErrorMessage("");

    try {
      const response = await fetch(
        `/api/ai-config/${encodeURIComponent(activeFile.fileName)}`,
        {
          method: "PUT",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            name: nextConfigName,
            ...nextConfig,
            setActive: true,
          }),
        },
      );

      if (!response.ok) {
        const payload = (await response.json().catch(() => ({}))) as {
          message?: string;
        };
        throw new Error(payload.message ?? "保存 AI 配置失败");
      }

      setAIConfigList((previous) => [
        {
          name: nextConfigName,
          config: nextConfig,
        },
        ...previous.filter((item) => item.name !== nextConfigName),
      ]);
      setSelectedAIConfigName(nextConfigName);
      setAIConfig(nextConfig);
      setDraftAIConfigName(nextConfigName);
      setDraftAIConfig(nextConfig);
      setAIConfigFormMessage("");
      setIsAIConfigModalOpen(false);
    } catch (error) {
      const message =
        error instanceof Error ? error.message : "保存 AI 配置失败";
      setAIConfigFormMessage(message);
    } finally {
      setAIConfigSaving(false);
    }
  };

  const onRunAIDetect = async () => {
    if (!activeFile || !selectedRow) {
      return;
    }
    if (isAIBatchRunning) {
      setAIResultMessage("批量 AI 任务运行中，暂不可发起单条回答");
      return;
    }

    const normalizedConfig = normalizeAIDetectConfigForColumns(
      aiConfig,
      activeFile.columns,
    );
    syncActiveAIConfigState(normalizedConfig);
    const runningConfigName = selectedAIConfigName;

    if (normalizedConfig.model.trim().length === 0) {
      setAIResultMessage("请先配置模型");
      return;
    }
    if (normalizedConfig.provider === "openai") {
      if (normalizedConfig.url.trim().length === 0) {
        setAIResultMessage("请先配置 OpenAI 兼容接口 URL");
        return;
      }
      if (normalizedConfig.apiKey.trim().length === 0) {
        setAIResultMessage("请先配置 OpenAI API Key");
        return;
      }
    }
    if (normalizedConfig.provider === "vertex") {
      if (normalizedConfig.vertexProject.trim().length === 0) {
        setAIResultMessage("请先配置 Vertex Project");
        return;
      }
      if (normalizedConfig.vertexLocation.trim().length === 0) {
        setAIResultMessage("请先配置 Vertex Location");
        return;
      }
    }
    if (normalizedConfig.submitFieldKeys.length === 0) {
      setAIResultMessage("请先在 AI 配置中选择提交回答字段");
      return;
    }
    if (normalizedConfig.prompt.trim().length === 0) {
      setAIResultMessage("请先配置 Prompt");
      return;
    }

    const fields = buildAIDetectFieldsForRow(
      activeFile.columns,
      selectedRow,
      normalizedConfig.submitFieldKeys,
    );

    if (fields.length === 0) {
      setAIResultMessage("当前记录没有可提交的回答字段");
      return;
    }

    aiStreamAbortRef.current?.abort();
    const controller = new AbortController();
    aiStreamAbortRef.current = controller;
    aiDetectStartedAtRef.current = Date.now();
    setAIDetectElapsedMs(0);
    setIsAIDetecting(true);
    setAIThinkingText("");
    setAIResultText("");
    setAIResultConfigName(runningConfigName);
    setAIResultMessage("");
    setErrorMessage("");

    try {
      const streamResult = await requestAIDetectResult(
        {
          provider: normalizedConfig.provider,
          url: normalizedConfig.url,
          model: normalizedConfig.model,
          apiKey: normalizedConfig.apiKey,
          vertexProject: normalizedConfig.vertexProject,
          vertexLocation: normalizedConfig.vertexLocation,
          prompt: normalizedConfig.prompt,
          fields,
          reasoningEffort: normalizedConfig.reasoningEffort,
          retryCount: normalizedConfig.retryCount,
        },
        {
          signal: controller.signal,
          onAnswerChunk: (chunk) => {
            setAIResultText((previous) => previous + chunk);
          },
          onThinkingChunk: (chunk) => {
            setAIThinkingText((previous) => previous + chunk);
          },
        },
      );
      setAIResultText(streamResult.answerText);
      setAIThinkingText(streamResult.thinkingText);
      if (
        streamResult.answerText.trim().length === 0 &&
        streamResult.thinkingText.trim().length === 0
      ) {
        setAIResultMessage("AI 返回为空");
      } else {
        setAIResultMessage(
          `AI 回答完成（配置：${runningConfigName}），可直接保存到目标字段`,
        );
      }
    } catch (error) {
      if (controller.signal.aborted) {
        setAIResultMessage("AI 回答已取消");
      } else {
        const message = error instanceof Error ? error.message : "AI 回答失败";
        setAIResultMessage(message);
      }
    } finally {
      if (aiStreamAbortRef.current === controller) {
        aiStreamAbortRef.current = null;
      }
      if (aiDetectStartedAtRef.current) {
        setAIDetectElapsedMs(Date.now() - aiDetectStartedAtRef.current);
        aiDetectStartedAtRef.current = null;
      }
      setIsAIDetecting(false);
    }
  };

  const applyBatchAIResultsToFile = (
    fileId: string,
    resultFieldKey: string,
    resultMap: Map<string, string>,
  ) => {
    if (resultMap.size === 0) {
      return;
    }

    let nextFileToPersist: FileViewState | null = null;
    setFiles((previous) =>
      previous.map((file) => {
        if (file.fileId !== fileId) {
          return file;
        }

        const nextRows = file.rows.map((row) => {
          const result = resultMap.get(row.rowId);
          if (result === undefined) {
            return row;
          }

          const currentCell = row.values[resultFieldKey];
          const nextCell: ParsedCell =
            currentCell?.type === "image" && currentCell.src
              ? {
                  type: "image",
                  src: currentCell.src,
                  value: result,
                }
              : {
                  type: "text",
                  value: result,
                };

          return {
            ...row,
            values: {
              ...row.values,
              [resultFieldKey]: nextCell,
            },
          };
        });

        const nextFile: FileViewState = {
          ...file,
          rows: nextRows,
        };
        nextFileToPersist = nextFile;
        return nextFile;
      }),
    );

    if (nextFileToPersist) {
      schedulePersistFileState(nextFileToPersist);
    }
  };

  const onRunBatchAIAnswer = async (rowIds?: string[]) => {
    if (!activeFile) {
      return;
    }
    if (isAIDetecting || isAIBatchRunning) {
      return;
    }

    const normalizedConfig = normalizeAIDetectConfigForColumns(
      aiConfig,
      activeFile.columns,
    );
    syncActiveAIConfigState(normalizedConfig);
    const runningConfigName = selectedAIConfigName;

    if (normalizedConfig.model.trim().length === 0) {
      setErrorMessage("请先配置模型");
      return;
    }
    if (normalizedConfig.provider === "openai") {
      if (normalizedConfig.url.trim().length === 0) {
        setErrorMessage("请先配置 OpenAI 兼容接口 URL");
        return;
      }
      if (normalizedConfig.apiKey.trim().length === 0) {
        setErrorMessage("请先配置 OpenAI API Key");
        return;
      }
    }
    if (normalizedConfig.provider === "vertex") {
      if (normalizedConfig.vertexProject.trim().length === 0) {
        setErrorMessage("请先配置 Vertex Project");
        return;
      }
      if (normalizedConfig.vertexLocation.trim().length === 0) {
        setErrorMessage("请先配置 Vertex Location");
        return;
      }
    }
    if (normalizedConfig.submitFieldKeys.length === 0) {
      setErrorMessage("请先在 AI 配置中选择提交回答字段");
      return;
    }
    if (normalizedConfig.prompt.trim().length === 0) {
      setErrorMessage("请先配置 Prompt");
      return;
    }
    if (normalizedConfig.resultFieldKey.trim().length === 0) {
      setErrorMessage("请先在 AI 配置中设置结果保存字段");
      return;
    }

    const targetResultColumn = activeFile.columns.find(
      (column) =>
        column.key === normalizedConfig.resultFieldKey && column.editable,
    );
    if (!targetResultColumn) {
      setErrorMessage("结果保存字段无效，请重新配置");
      return;
    }

    const rowIdSet = new Set(activeFile.rows.map((row) => row.rowId));
    const normalizedTargetRowIds = rowIds
      ? Array.from(new Set(rowIds.filter((rowId) => rowIdSet.has(rowId))))
      : null;
    const selectedRowIdSet = normalizedTargetRowIds
      ? new Set(normalizedTargetRowIds)
      : null;
    const targetRows =
      normalizedTargetRowIds && normalizedTargetRowIds.length > 0
        ? activeFile.rows.filter(
            (row) => selectedRowIdSet?.has(row.rowId) === true,
          )
        : normalizedTargetRowIds
          ? []
          : activeFile.rows;
    if (targetRows.length === 0) {
      setErrorMessage(
        normalizedTargetRowIds
          ? "请先至少选择一条数据再执行批量回答"
          : "当前文件没有可执行的行数据",
      );
      return;
    }

    const targetFileId = activeFile.fileId;
    const targetFileName = activeFile.fileName;
    const targetColumns = activeFile.columns;
    const resultMap = new Map<string, string>();
    let nextCursor = 0;
    const requestedConcurrency =
      normalizeAIBatchConcurrency(aiBatchConcurrency);
    const workerCount = Math.min(requestedConcurrency, targetRows.length);

    aiBatchAbortRef.current?.abort();
    const controller = new AbortController();
    aiBatchAbortRef.current = controller;
    setAIBatchTask({
      status: "running",
      fileId: targetFileId,
      fileName: targetFileName,
      total: targetRows.length,
      completed: 0,
      success: 0,
      failed: 0,
      message:
        normalizedTargetRowIds && normalizedTargetRowIds.length > 0
          ? `已选择 ${targetRows.length} 条，并发 ${workerCount} 线程`
          : `并发 ${workerCount} 线程`,
    });
    setErrorMessage("");
    setAIResultMessage("");

    const runWorker = async () => {
      while (!controller.signal.aborted) {
        const currentIndex = nextCursor;
        nextCursor += 1;
        if (currentIndex >= targetRows.length) {
          return;
        }

        const row = targetRows[currentIndex];
        try {
          const fields = buildAIDetectFieldsForRow(
            targetColumns,
            row,
            normalizedConfig.submitFieldKeys,
          );
          if (fields.length === 0) {
            throw new Error("没有可提交的回答字段");
          }

          const streamResult = await requestAIDetectResult(
            {
              provider: normalizedConfig.provider,
              url: normalizedConfig.url,
              model: normalizedConfig.model,
              apiKey: normalizedConfig.apiKey,
              vertexProject: normalizedConfig.vertexProject,
              vertexLocation: normalizedConfig.vertexLocation,
              prompt: normalizedConfig.prompt,
              fields,
              reasoningEffort: normalizedConfig.reasoningEffort,
              retryCount: normalizedConfig.retryCount,
            },
            { signal: controller.signal },
          );
          const text = composeAISaveTextWithConfigName(
            streamResult.answerText,
            streamResult.thinkingText,
            runningConfigName,
          );

          if (text.trim().length === 0) {
            throw new Error("AI 返回为空");
          }
          resultMap.set(row.rowId, text);
          setAIBatchTask((previous) => ({
            ...previous,
            completed: previous.completed + 1,
            success: previous.success + 1,
          }));
        } catch {
          if (controller.signal.aborted) {
            return;
          }
          setAIBatchTask((previous) => ({
            ...previous,
            completed: previous.completed + 1,
            failed: previous.failed + 1,
          }));
        }
      }
    };

    try {
      await Promise.all(Array.from({ length: workerCount }, () => runWorker()));

      if (controller.signal.aborted) {
        return;
      }

      applyBatchAIResultsToFile(
        targetFileId,
        normalizedConfig.resultFieldKey,
        resultMap,
      );

      setAIBatchTask((previous) => ({
        ...previous,
        status: "completed",
        message: `结果已写入字段：${targetResultColumn.title}（配置：${runningConfigName}）`,
      }));
      setErrorMessage("");
    } catch (error) {
      if (controller.signal.aborted) {
        return;
      }

      const message =
        error instanceof Error ? error.message : "批量 AI 回答任务执行失败";
      setAIBatchTask((previous) => ({
        ...previous,
        status: "completed",
        message,
      }));
      setErrorMessage(message);
    } finally {
      if (aiBatchAbortRef.current === controller) {
        aiBatchAbortRef.current = null;
      }
    }
  };

  const onRunSelectedBatchAIAnswer = async () => {
    await onRunBatchAIAnswer(batchSelectedRowIds);
  };

  const onSaveAIResult = () => {
    if (!activeFile || !selectedRow) {
      return;
    }

    const composedText = composeAISaveTextWithConfigName(
      aiResultText,
      aiThinkingText,
      aiResultConfigName,
    );
    if (composedText.trim().length === 0) {
      setAIResultMessage("暂无可保存的 AI 返回结果");
      return;
    }

    const resultFieldKey = aiConfig.resultFieldKey;
    if (resultFieldKey.trim().length === 0) {
      setAIResultMessage("请先在 AI 配置中设置结果保存字段");
      return;
    }

    const targetColumn = activeFile.columns.find(
      (column) => column.key === resultFieldKey,
    );
    if (!targetColumn || !targetColumn.editable) {
      setAIResultMessage("结果字段无效，请重新配置");
      return;
    }

    setIsSavingAIResult(true);
    try {
      onEditCell(selectedRow.rowId, resultFieldKey, composedText);
      if (aiThinkingText.trim().length > 0) {
        setAIResultMessage(
          `已保存到字段：${targetColumn.title}（配置：${aiResultConfigName}，含思考过程）`,
        );
      } else {
        setAIResultMessage(
          `已保存到字段：${targetColumn.title}（配置：${aiResultConfigName}）`,
        );
      }
    } finally {
      setIsSavingAIResult(false);
    }
  };

  const onExportFile = async () => {
    if (!activeFile) {
      return;
    }

    setIsExporting(true);
    setErrorMessage("");

    try {
      const exportColumns = activeFile.columns;
      const headers = exportColumns.map((column) => column.title);
      const rows = activeFile.rows.map((row) =>
        exportColumns.map((column) => row.values[column.key]?.value ?? ""),
      );

      const response = await fetch("/api/files/export", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          fileName: activeFile.fileName,
          headers,
          rows,
        }),
      });

      if (!response.ok) {
        const payload = (await response.json().catch(() => ({}))) as {
          message?: string;
        };
        throw new Error(payload.message ?? "导出失败");
      }

      const blob = await response.blob();
      const headerFileName = getFileNameFromDisposition(
        response.headers.get("Content-Disposition"),
      );
      const fallbackFileName = `${activeFile.fileName.replace(/\.[^.]+$/, "")}-导出.xlsx`;
      downloadBlob(blob, headerFileName ?? fallbackFileName);
    } catch (error) {
      const message = error instanceof Error ? error.message : "导出失败";
      setErrorMessage(message);
    } finally {
      setIsExporting(false);
    }
  };

  const onUploadClick = () => {
    uploadInputRef.current?.click();
  };

  const onUploadFile = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const selected = event.target.files?.[0];
    if (!selected) {
      return;
    }

    setIsUploading(true);
    setErrorMessage("");

    try {
      const formData = new FormData();
      formData.append("file", selected);
      const response = await fetch("/api/files/upload", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const payload = (await response.json().catch(() => ({}))) as {
          message?: string;
        };
        throw new Error(payload.message ?? "文件解析失败");
      }

      const parsed = (await response.json()) as ParsedFile;
      let parsedImageCellCount = 0;
      let parsedTextLikeImageCellCount = 0;
      const textLikeSamples: string[] = [];
      parsed.rows.forEach((row) => {
        parsed.columns.forEach((column) => {
          const cell = row.values[column.key];
          if (!cell) {
            return;
          }
          if (
            cell.type === "image" &&
            typeof cell.src === "string" &&
            cell.src
          ) {
            parsedImageCellCount += 1;
            return;
          }
          if (
            cell.type === "text" &&
            typeof cell.value === "string" &&
            /\.(png|jpe?g|webp|gif|bmp|tiff?)([?#].*)?$/i.test(
              cell.value.trim(),
            )
          ) {
            parsedTextLikeImageCellCount += 1;
            if (textLikeSamples.length < 8) {
              textLikeSamples.push(
                `row=${row.rowId} column=${column.title} value=${cell.value.trim()}`,
              );
            }
          }
        });
      });
      // eslint-disable-next-line no-console
      console.log(
        `[UIParsedImage] file=${parsed.fileName} imageCells=${parsedImageCellCount} textLikeImageCells=${parsedTextLikeImageCellCount}`,
      );
      if (textLikeSamples.length > 0) {
        // eslint-disable-next-line no-console
        console.log(
          `[UIParsedImageTextLike] ${JSON.stringify(textLikeSamples)}`,
        );
      }
      const defaultDisplayKeys = getAllColumnKeys(parsed.columns);
      let initialDisplayKeys = defaultDisplayKeys;
      let initialEditableKeys: string[] = [];
      let shouldShowColumnModal = true;
      let nextPendingNotice = "";

      try {
        const prefsRes = await fetch(
          `/api/column-prefs/${encodeURIComponent(parsed.fileName)}`,
        );
        if (prefsRes.ok) {
          const prefsData = (await prefsRes.json()) as {
            config: ColumnPrefsConfig | null;
          };
          if (prefsData.config) {
            const normalizedSaved = normalizeColumnSelection(
              parsed.columns,
              prefsData.config.displayKeys,
              prefsData.config.editableKeys,
            );
            const currentSignature = getFieldSignature(parsed.columns);
            if (prefsData.config.fieldSignature === currentSignature) {
              const nextFile = toViewState(
                parsed,
                normalizedSaved.displayKeys,
                normalizedSaved.editableKeys,
              );
              setFiles((previous) => [...previous, nextFile]);
              setActiveFileId(nextFile.fileId);
              persistFileState(nextFile);
              shouldShowColumnModal = false;
            } else {
              nextPendingNotice =
                "检测到该 Excel 字段与已保存配置不一致，请重新选择并保存新配置。";
              initialDisplayKeys = normalizedSaved.displayKeys;
              initialEditableKeys = normalizedSaved.editableKeys;
            }
          }
        }
      } catch {
        // Ignore and fall back to default selection
      }

      if (shouldShowColumnModal) {
        setPendingFile(parsed);
        setPendingSelectedDisplayKeys(initialDisplayKeys);
        setPendingEditableColumnKeys(initialEditableKeys);
        setPendingConfigNotice(nextPendingNotice);
        setPendingConfigMode("import");
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : "上传失败";
      setErrorMessage(message);
    } finally {
      setIsUploading(false);
      event.target.value = "";
    }
  };

  const onRemoveFile = (fileId: string) => {
    cancelScheduledPersist(fileId);
    fetch(`/api/files/${encodeURIComponent(fileId)}`, {
      method: "DELETE",
    }).catch(() => {});
    setFiles((previous) => {
      const next = previous.filter((file) => file.fileId !== fileId);
      if (activeFileId === fileId) {
        setActiveFileId(next[0]?.fileId ?? null);
      }
      return next;
    });
  };

  const onTogglePendingDisplayColumn = (columnKey: string) => {
    if (!pendingFile) {
      return;
    }
    setPendingSelectedDisplayKeys((previous) => {
      const shouldHide = previous.includes(columnKey);
      const next = shouldHide
        ? previous.filter((key) => key !== columnKey)
        : [...previous, columnKey];

      if (shouldHide) {
        setPendingEditableColumnKeys((editableKeys) =>
          editableKeys.filter((key) => key !== columnKey),
        );
      }
      return next;
    });
  };

  const onTogglePendingEditableColumn = (columnKey: string) => {
    if (!pendingFile) {
      return;
    }
    setPendingEditableColumnKeys((previous) => {
      const exists = previous.includes(columnKey);
      const next = exists
        ? previous.filter((key) => key !== columnKey)
        : [...previous, columnKey];

      if (!exists) {
        setPendingSelectedDisplayKeys((displayKeys) =>
          displayKeys.includes(columnKey)
            ? displayKeys
            : [...displayKeys, columnKey],
        );
      }

      return next;
    });
  };

  const onPendingSelectAllDisplayColumns = () => {
    if (!pendingFile) {
      return;
    }
    setPendingSelectedDisplayKeys(getAllColumnKeys(pendingFile.columns));
  };

  const onPendingClearDisplayColumns = () => {
    setPendingSelectedDisplayKeys([]);
    setPendingEditableColumnKeys([]);
  };

  const onPendingClearEditableColumns = () => {
    setPendingEditableColumnKeys([]);
  };

  const onCancelPendingFile = () => {
    resetPendingConfigState();
  };

  const onConfirmPendingFile = () => {
    if (!pendingFile) {
      return;
    }

    if (pendingConfigMode === "edit") {
      patchActiveFile((file) => {
        const nextFile = applyColumnConfigToFile(
          file,
          pendingSelectedDisplayKeys,
          pendingEditableColumnKeys,
        );
        persistColumnPrefs(nextFile);
        return nextFile;
      });
      resetPendingConfigState();
      return;
    }

    const nextFile = toViewState(
      pendingFile,
      pendingSelectedDisplayKeys,
      pendingEditableColumnKeys,
    );
    setFiles((previous) => [...previous, nextFile]);
    setActiveFileId(nextFile.fileId);
    persistColumnPrefs(nextFile);
    persistFileState(nextFile);
    resetPendingConfigState();
  };

  const onToggleDisplayColumn = (columnKey: string) => {
    patchActiveFile((file) => {
      if (file.selectedEditableColumnKeys.includes(columnKey)) {
        return file;
      }

      const exists = file.selectedDisplayColumnKeys.includes(columnKey);
      const selectedDisplayColumnKeys = exists
        ? file.selectedDisplayColumnKeys.filter((key) => key !== columnKey)
        : [...file.selectedDisplayColumnKeys, columnKey];

      const normalized = normalizeColumnSelection(
        file.columns,
        selectedDisplayColumnKeys,
        file.selectedEditableColumnKeys,
      );
      const nextFile: FileViewState = {
        ...file,
        selectedDisplayColumnKeys: normalized.displayKeys,
        selectedEditableColumnKeys: normalized.editableKeys,
      };
      persistColumnPrefs(nextFile);
      return nextFile;
    });
  };

  const onFilterChange = (
    type: "level1" | "level2" | "time",
    value: string,
  ) => {
    patchActiveFile((file) => ({
      ...file,
      level1Filter: type === "level1" ? value : file.level1Filter,
      level2Filter: type === "level2" ? value : file.level2Filter,
      timeFilter: type === "time" ? value : file.timeFilter,
    }));
  };

  const onToggleBatchRowSelection = (rowId: string) => {
    setBatchSelectedRowIds((previous) => {
      if (previous.includes(rowId)) {
        return previous.filter((item) => item !== rowId);
      }
      return [...previous, rowId];
    });
  };

  const onSelectAllBatchRows = () => {
    setBatchSelectedRowIds(visibleRows.map((row) => row.rowId));
  };

  const onClearBatchRows = () => {
    setBatchSelectedRowIds([]);
  };

  const onEditCell = (rowId: string, columnKey: string, value: string) => {
    patchActiveFile((file) => ({
      ...file,
      rows: file.rows.map((row) => {
        if (row.rowId !== rowId) {
          return row;
        }

        const currentCell = row.values[columnKey];

        return {
          ...row,
          values: {
            ...row.values,
            [columnKey]:
              currentCell?.type === "image" && currentCell.src
                ? {
                    type: "image",
                    src: currentCell.src,
                    value,
                  }
                : {
                    type: "text",
                    value,
                  },
          },
        };
      }),
    }));
  };

  const getLatexToggleKey = (columnKey: string) =>
    activeFileId ? `${activeFileId}::${columnKey}` : columnKey;

  const onToggleLatexRender = (columnKey: string) => {
    const key = getLatexToggleKey(columnKey);
    setLatexRenderOverrides((previous) => ({
      ...previous,
      [key]: !(previous[key] ?? true),
    }));
  };

  const renderReadonlyCell = (
    row: ParsedRow,
    column: ParsedColumn,
    cell: ParsedCell | undefined,
    shouldRenderLatex: boolean,
  ) => {
    if (!cell) {
      return <span className="empty-text">-</span>;
    }

    if (cell.type === "image" && cell.src) {
      return (
        <div className="image-cell">
          <img
            src={cell.src}
            alt={cell.value || "Excel图片"}
            onClick={() => setPreviewImageSrc(cell.src!)}
            onError={() => {
              logUIImageRenderError(row.rowId, column.title, cell.src ?? "");
            }}
          />
          {cell.value ? <span>{cell.value}</span> : null}
        </div>
      );
    }

    const textValue = cell.value ?? "";
    if (cell.type === "text" && textValue.length > 0) {
      const hasLatex = hasLatexSyntax(textValue);
      const autoDisplayLatex = shouldAutoDisplayLatex(textValue);
      if (hasLatex && shouldRenderLatex) {
        return (
          <LatexRenderer value={textValue} forceDisplay={autoDisplayLatex} />
        );
      }
      return hasLatex ? (
        <div className="latex-plain">{textValue}</div>
      ) : (
        <div className="plain-text-value">{textValue}</div>
      );
    }

    return cell.value ? (
      <div className="plain-text-value">{cell.value}</div>
    ) : (
      <span className="empty-text">-</span>
    );
  };

  const renderCellContent = (
    row: ParsedRow,
    column: ParsedColumn,
    shouldRenderLatex = true,
  ) => {
    const cell = row.values[column.key];
    if (!column.editable) {
      return renderReadonlyCell(row, column, cell, shouldRenderLatex);
    }

    const currentValue = cell?.value ?? "";

    if (isQualifiedColumnTitle(column.title)) {
      const stableOptions = ["", "合格", "不合格"];
      const shouldAppendCurrent =
        currentValue.length > 0 && !stableOptions.includes(currentValue);
      return (
        <select
          className="qualified-select"
          value={currentValue}
          onChange={(event) =>
            onEditCell(row.rowId, column.key, event.target.value)
          }
        >
          <option value="">未填写</option>
          <option value="合格">合格</option>
          <option value="不合格">不合格</option>
          {shouldAppendCurrent ? (
            <option value={currentValue}>{currentValue}</option>
          ) : null}
        </select>
      );
    }

    if (isOpensourceColumnTitle(column.title)) {
      const stableOptions = ["", "是", "否"];
      const shouldAppendCurrent =
        currentValue.length > 0 && !stableOptions.includes(currentValue);
      return (
        <select
          className="qualified-select"
          value={currentValue}
          onChange={(event) =>
            onEditCell(row.rowId, column.key, event.target.value)
          }
        >
          <option value="">未填写</option>
          <option value="是">是</option>
          <option value="否">否</option>
          {shouldAppendCurrent ? (
            <option value={currentValue}>{currentValue}</option>
          ) : null}
        </select>
      );
    }

    if (isInspectorColumnTitle(column.title)) {
      return (
        <input
          className="inspector-input"
          value={currentValue}
          onChange={(event) =>
            onEditCell(row.rowId, column.key, event.target.value)
          }
          placeholder="请输入质检员"
        />
      );
    }

    if (isFeedbackColumnTitle(column.title)) {
      return (
        <textarea
          className="feedback-input"
          value={currentValue}
          onChange={(event) =>
            onEditCell(row.rowId, column.key, event.target.value)
          }
          placeholder="请输入质检反馈意见"
        />
      );
    }

    if (cell?.type === "image" && cell.src) {
      return (
        <div className="image-cell">
          <img
            src={cell.src}
            alt={cell.value || "Excel图片"}
            onClick={() => setPreviewImageSrc(cell.src!)}
            onError={() => {
              logUIImageRenderError(row.rowId, column.title, cell.src ?? "");
            }}
          />
          <input
            className="editable-text-input"
            value={currentValue}
            onChange={(event) =>
              onEditCell(row.rowId, column.key, event.target.value)
            }
            placeholder={`请输入${column.title}`}
          />
        </div>
      );
    }

    return (
      <input
        className="editable-text-input"
        value={currentValue}
        onChange={(event) =>
          onEditCell(row.rowId, column.key, event.target.value)
        }
        placeholder={`请输入${column.title}`}
      />
    );
  };

  const renderDetailField = (column: ParsedColumn, isHidden = false) => {
    if (!selectedRow) return null;
    const isRequired = column.editable;
    const isChecked = !isHidden;
    const cell = selectedRow.values[column.key];
    const hasLatex =
      !column.editable &&
      !isHidden &&
      cell?.type === "text" &&
      typeof cell.value === "string" &&
      hasLatexSyntax(cell.value);
    const latexToggleKey = getLatexToggleKey(column.key);
    const isLatexRenderingEnabled =
      latexRenderOverrides[latexToggleKey] ?? true;

    return (
      <div
        key={`${selectedRow.rowId}_${column.key}`}
        className={`detail-field ${isHidden ? "hidden-field" : ""}`}
      >
        <div className="detail-label">
          <button
            type="button"
            className={`field-toggle ${isRequired ? "locked" : ""} ${isChecked ? "checked" : ""}`}
            onClick={() => {
              if (!isRequired) {
                onToggleDisplayColumn(column.key);
              }
            }}
            title={
              isRequired
                ? "可编辑字段必须展示"
                : isHidden
                  ? "点击显示此字段"
                  : "点击隐藏此字段"
            }
          />
          <div className="field-name-wrap">
            <span className="field-name">{column.title}</span>
            {hasLatex ? (
              <label
                className="latex-toggle"
                title="控制该字段是否按 LaTeX 公式渲染"
              >
                <input
                  type="checkbox"
                  checked={isLatexRenderingEnabled}
                  onChange={() => onToggleLatexRender(column.key)}
                  aria-label={`${column.title} 的 LaTeX 渲染开关`}
                />
                <span>LaTeX渲染</span>
              </label>
            ) : null}
          </div>
          {column.editable ? (
            <span className="field-badge badge-editable">可编辑</span>
          ) : null}
          {isRequired ? (
            <span className="field-badge badge-locked">必显</span>
          ) : null}
        </div>
        {!isHidden ? (
          <div className="detail-value">
            {renderCellContent(selectedRow, column, isLatexRenderingEnabled)}
          </div>
        ) : null}
      </div>
    );
  };

  return (
    <div className="app-shell">
      {/* ─── Header Bar ─── */}
      <header className="header-bar">
        <div className="header-inner">
          <div className="header-brand">
            <div className="brand-icon">
              <svg viewBox="0 0 24 24">
                <path
                  d="M9 2L4 7v13a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V4a2 2 0 0 0-2-2H9zm0 0v5H4m4 4h8m-8 4h8m-8 4h4"
                  fill="none"
                  stroke="white"
                  strokeWidth="1.5"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
              </svg>
            </div>
            <h1>质检工作台</h1>
          </div>

          {/* ─── File Tabs ─── */}
          <div className="file-tabs">
            {files.map((file) => (
              <div
                key={file.fileId}
                className={`file-tab ${file.fileId === activeFileId ? "active" : ""}`}
              >
                <button
                  type="button"
                  style={{
                    all: "unset",
                    cursor: "pointer",
                    display: "contents",
                  }}
                  onClick={() => setActiveFileId(file.fileId)}
                >
                  {file.fileName}
                </button>
                <span className="tab-badge">{file.rows.length}</span>
                <button
                  type="button"
                  className="tab-close"
                  onClick={(e) => {
                    e.stopPropagation();
                    onRemoveFile(file.fileId);
                  }}
                  title="关闭"
                >
                  ×
                </button>
              </div>
            ))}
          </div>

          {/* ─── Header Actions ─── */}
          <div className="header-actions">
            {errorMessage ? (
              <span className="error-text">{errorMessage}</span>
            ) : null}
            {aiBatchTask.total > 0 ? (
              <div
                className={`ai-batch-status ${isAIBatchRunning ? "running" : "completed"}`}
              >
                <div className="ai-batch-status-head">
                  <span>{getAIBatchTaskStatusText(aiBatchTask)}</span>
                  <strong>
                    {aiBatchTask.completed}/{aiBatchTask.total}
                  </strong>
                </div>
                <div className="ai-batch-progress">
                  <div
                    className="ai-batch-progress-bar"
                    style={{ width: `${aiBatchProgressPercent}%` }}
                  />
                </div>
                <div className="ai-batch-counts">
                  <span className="ai-batch-success">
                    成功 {aiBatchTask.success}
                  </span>
                  <span className="ai-batch-failed">
                    失败 {aiBatchTask.failed}
                  </span>
                  <span>{aiBatchProgressPercent}%</span>
                </div>
                {aiBatchTask.fileName ? (
                  <div
                    className="ai-batch-file"
                    title={aiBatchTask.fileName}
                  >{`任务文件：${aiBatchTask.fileName}`}</div>
                ) : null}
                {aiBatchTask.message ? (
                  <div className="ai-batch-message">{aiBatchTask.message}</div>
                ) : null}
              </div>
            ) : null}
            <button
              type="button"
              className="theme-toggle"
              onClick={toggleTheme}
              title={theme === "dark" ? "切换浅色主题" : "切换深色主题"}
            >
              {theme === "dark" ? <IconSun /> : <IconMoon />}
            </button>
            <button
              type="button"
              className="btn"
              onClick={onOpenActiveFileConfig}
              disabled={!activeFile}
            >
              字段配置
            </button>
            <button
              type="button"
              className="btn"
              onClick={onOpenAIConfigModal}
              disabled={!activeFile || aiConfigLoading}
            >
              AI回答配置
            </button>
            <button
              type="button"
              className="btn"
              onClick={onExportFile}
              disabled={isExporting || !activeFile}
            >
              <IconDownload />
              {isExporting ? "导出中..." : "导出 Excel"}
            </button>
            <button
              type="button"
              className="btn btn-primary"
              onClick={onUploadClick}
              disabled={isUploading}
            >
              <IconUpload />
              {isUploading ? "解析中..." : "导入 Excel"}
            </button>
            <input
              ref={uploadInputRef}
              type="file"
              accept=".xls,.xlsx"
              className="hidden-input"
              onChange={onUploadFile}
            />
          </div>
        </div>
      </header>

      {/* ─── Main Content ─── */}
      <main className="main-content">
        {!activeFile ? (
          <section className="placeholder">
            <div className="placeholder-icon">
              <IconFile />
            </div>
            <h2>等待文件导入</h2>
            <p>
              点击右上角「导入
              Excel」按钮，支持展示/可编辑字段配置、level1/level2/时间筛选、图片展示与导出。
            </p>
          </section>
        ) : (
          <>
            {/* ─── Toolbar ─── */}
            <section className="toolbar">
              <div className="filter-group">
                <label htmlFor="level1-filter">level1</label>
                <select
                  id="level1-filter"
                  value={activeFile.level1Filter}
                  onChange={(event) =>
                    onFilterChange("level1", event.target.value)
                  }
                >
                  <option value={ALL_FILTER_VALUE}>{ALL_FILTER_VALUE}</option>
                  {activeFile.level1Options.map((item) => (
                    <option key={item} value={item}>
                      {item}
                    </option>
                  ))}
                </select>
              </div>
              <div className="filter-group">
                <label htmlFor="level2-filter">level2</label>
                <select
                  id="level2-filter"
                  value={activeFile.level2Filter}
                  onChange={(event) =>
                    onFilterChange("level2", event.target.value)
                  }
                >
                  <option value={ALL_FILTER_VALUE}>{ALL_FILTER_VALUE}</option>
                  {activeFile.level2Options.map((item) => (
                    <option key={item} value={item}>
                      {item}
                    </option>
                  ))}
                </select>
              </div>
              <div className="filter-group">
                <label htmlFor="time-filter">时间</label>
                <select
                  id="time-filter"
                  value={activeFile.timeFilter}
                  onChange={(event) =>
                    onFilterChange("time", event.target.value)
                  }
                >
                  <option value={ALL_FILTER_VALUE}>{ALL_FILTER_VALUE}</option>
                  {timeOptions.map((item) => (
                    <option key={item} value={item}>
                      {item}
                    </option>
                  ))}
                </select>
              </div>
              <div className="stats">
                <strong>{visibleRows.length}</strong>
                <span>可选条目</span>
              </div>
              <div className="stats">
                <strong>{batchSelectedRowIds.length}</strong>
                <span>已勾选条目</span>
              </div>
            </section>

            <section className="batch-answer-layout batch-answer-inline">
              <div className="batch-answer-header">
                <h3>AI批量回答</h3>
                <p>
                  在列表中勾选后执行批量回答，点击记录可在右侧查看字段详情。
                </p>
              </div>
              <div className="batch-answer-actions">
                <label className="ai-run-config">
                  <span>运行配置</span>
                  <select
                    value={selectedAIConfigName}
                    onChange={(event) =>
                      onSelectAIConfigForRun(event.target.value)
                    }
                    disabled={
                      aiConfigLoading || isAIDetecting || isAIBatchRunning
                    }
                  >
                    {aiConfigList.map((item) => (
                      <option key={item.name} value={item.name}>
                        {item.name}
                      </option>
                    ))}
                  </select>
                </label>
                <label className="ai-run-config">
                  <span>并发数</span>
                  <input
                    type="number"
                    min={MIN_AI_BATCH_CONCURRENCY}
                    max={MAX_AI_BATCH_CONCURRENCY}
                    step={1}
                    value={aiBatchConcurrency}
                    onChange={(event) =>
                      setAIBatchConcurrency(
                        normalizeAIBatchConcurrency(Number(event.target.value)),
                      )
                    }
                    disabled={isAIBatchRunning}
                  />
                </label>
                <button
                  type="button"
                  className="btn"
                  onClick={onSelectAllBatchRows}
                  disabled={visibleRows.length === 0 || isAIBatchRunning}
                >
                  全选可见
                </button>
                <button
                  type="button"
                  className="btn"
                  onClick={onClearBatchRows}
                  disabled={
                    batchSelectedRowIds.length === 0 || isAIBatchRunning
                  }
                >
                  清空勾选
                </button>
                <button
                  type="button"
                  className="btn btn-primary"
                  onClick={onRunSelectedBatchAIAnswer}
                  disabled={
                    aiConfigLoading ||
                    isAIDetecting ||
                    isAIBatchRunning ||
                    batchSelectedRowIds.length === 0
                  }
                >
                  {isAIBatchRunning
                    ? "AI批量回答中..."
                    : `批量回答已选 ${batchSelectedRowIds.length} 条`}
                </button>
                <div className="ai-result-target">
                  <span>保存字段：</span>
                  <strong>{aiResultFieldTitle || "未配置"}</strong>
                  <span className="ai-target-sep">|</span>
                  <span>重试：</span>
                  <strong className="ai-retry-count">
                    {aiConfig.retryCount}次
                  </strong>
                </div>
              </div>
            </section>

            {/* ─── Detail Layout ─── */}
            <section
              className={`detail-layout ${selectedRow ? "with-detail" : "list-only"}`}
            >
              {/* ─── Record List ─── */}
              <aside className="record-list">
                <div className="record-list-header">
                  <h3>数据列表</h3>
                  <p>
                    {selectedRow
                      ? "左侧为列表，右侧为当前记录详情"
                      : "默认列表模式，点击任意记录后右侧展开字段详情"}
                  </p>
                </div>
                {visibleRows.length === 0 ? (
                  <div className="record-list-empty">当前筛选条件下无数据</div>
                ) : (
                  <div className="record-list-items">
                    {visibleRows.map((row, index) => {
                      const checked = batchSelectedRowIdSet.has(row.rowId);
                      return (
                        <div
                          key={row.rowId}
                          role="button"
                          tabIndex={0}
                          className={`record-item ${selectedRowId === row.rowId ? "active" : ""} ${checked ? "batch-selected" : ""}`}
                          onClick={() =>
                            setSelectedRowId((previous) =>
                              previous === row.rowId ? null : row.rowId,
                            )
                          }
                          onKeyDown={(event) => {
                            if (event.key === "Enter" || event.key === " ") {
                              event.preventDefault();
                              setSelectedRowId((previous) =>
                                previous === row.rowId ? null : row.rowId,
                              );
                            }
                          }}
                        >
                          <div className="record-item-head">
                            <label
                              className="record-item-check"
                              onClick={(event) => event.stopPropagation()}
                            >
                              <input
                                type="checkbox"
                                checked={checked}
                                onChange={() =>
                                  onToggleBatchRowSelection(row.rowId)
                                }
                              />
                            </label>
                            <strong>第 {index + 1} 条</strong>
                          </div>
                          <span>{getRowPreviewText(row)}</span>
                        </div>
                      );
                    })}
                  </div>
                )}
              </aside>

              {/* ─── Record Detail ─── */}
              {selectedRow ? (
                <section className="record-detail">
                  <div className="record-detail-header">
                    <h3>字段详情</h3>
                    <span>点击字段左侧勾选框可控制显示/隐藏</span>
                  </div>
                  <div className="record-detail-ai-toolbar">
                    <div className="record-detail-ai-actions">
                      <label className="ai-run-config">
                        <span>运行配置</span>
                        <select
                          value={selectedAIConfigName}
                          onChange={(event) =>
                            onSelectAIConfigForRun(event.target.value)
                          }
                          disabled={
                            aiConfigLoading || isAIDetecting || isAIBatchRunning
                          }
                        >
                          {aiConfigList.map((item) => (
                            <option key={item.name} value={item.name}>
                              {item.name}
                            </option>
                          ))}
                        </select>
                      </label>
                      <button
                        type="button"
                        className="btn btn-primary"
                        onClick={onRunAIDetect}
                        disabled={
                          !selectedRow ||
                          isAIDetecting ||
                          aiConfigLoading ||
                          isAIBatchRunning
                        }
                      >
                        {isAIDetecting
                          ? `AI回答中 ${aiDetectElapsedText}`
                          : "发送AI回答"}
                      </button>
                      <button
                        type="button"
                        className="btn"
                        onClick={onSaveAIResult}
                        disabled={
                          !selectedRow ||
                          isAIDetecting ||
                          isAIBatchRunning ||
                          isSavingAIResult ||
                          !hasAISaveContent ||
                          aiConfig.resultFieldKey.trim().length === 0
                        }
                      >
                        {isSavingAIResult ? "保存中..." : "保存AI回答"}
                      </button>
                      <div className="ai-result-target">
                        <span>保存字段：</span>
                        <strong>{aiResultFieldTitle || "未配置"}</strong>
                        <span className="ai-target-sep">|</span>
                        <span>重试：</span>
                        <strong className="ai-retry-count">
                          {aiConfig.retryCount}次
                        </strong>
                      </div>
                    </div>
                    <div className="ai-stream-group">
                      <label className="ai-stream-label">
                        AI响应（流式，含思考过程）
                      </label>
                      <textarea
                        className="ai-stream-input ai-stream-input-large"
                        value={aiMergedStreamText}
                        onChange={(event) => {
                          // After manual edit, treat the merged text as final answer body.
                          setAIThinkingText("");
                          setAIResultText(event.target.value);
                        }}
                        placeholder="点击“发送AI回答”后，这里会流式展示模型输出（思考+回答），可手动编辑后再保存。"
                      />
                    </div>
                    {aiResultMessage ? (
                      <div className="ai-stream-message">{aiResultMessage}</div>
                    ) : null}
                  </div>
                  <div className="detail-fields">
                    {displayColumns.map((column) =>
                      renderDetailField(column, false),
                    )}
                    {hiddenColumns.length > 0 ? (
                      <div className="hidden-fields-section">
                        <button
                          type="button"
                          className={`hidden-fields-toggle ${showHiddenFields ? "expanded" : ""}`}
                          onClick={() => setShowHiddenFields(!showHiddenFields)}
                        >
                          <IconChevron />
                          <span>{hiddenColumns.length} 个已隐藏字段</span>
                        </button>
                        {showHiddenFields ? (
                          <div className="hidden-fields-list">
                            {hiddenColumns.map((column) =>
                              renderDetailField(column, true),
                            )}
                          </div>
                        ) : null}
                      </div>
                    ) : null}
                  </div>
                </section>
              ) : null}
            </section>
          </>
        )}
      </main>

      {/* ─── Column Selection Modal ─── */}
      {pendingFile ? (
        <div className="column-modal-mask">
          <div className="column-modal">
            <h3>
              {pendingConfigMode === "edit"
                ? "编辑字段展示与可编辑"
                : "配置字段展示与可编辑"}
            </h3>
            <p>{pendingFile.fileName}</p>
            {pendingConfigNotice ? (
              <div className="column-modal-notice">{pendingConfigNotice}</div>
            ) : null}
            <div className="column-modal-actions">
              <button
                type="button"
                className="btn"
                onClick={onPendingSelectAllDisplayColumns}
              >
                全选展示
              </button>
              <button
                type="button"
                className="btn"
                onClick={onPendingClearDisplayColumns}
              >
                清空展示
              </button>
              <button
                type="button"
                className="btn"
                onClick={onPendingClearEditableColumns}
              >
                清空可编辑
              </button>
            </div>
            <div className="column-modal-list">
              {pendingFile.columns.map((column) => {
                const checkedDisplay = pendingSelectedDisplayKeys.includes(
                  column.key,
                );
                const checkedEditable = pendingEditableColumnKeys.includes(
                  column.key,
                );
                return (
                  <div
                    key={column.key}
                    className={`column-config-row ${checkedEditable ? "editable-column-row" : ""}`}
                  >
                    <span className="column-config-name">{column.title}</span>
                    <label className="column-config-switch">
                      <input
                        type="checkbox"
                        checked={checkedDisplay}
                        onChange={() =>
                          onTogglePendingDisplayColumn(column.key)
                        }
                      />
                      <span>展示</span>
                    </label>
                    <label className="column-config-switch">
                      <input
                        type="checkbox"
                        checked={checkedEditable}
                        onChange={() =>
                          onTogglePendingEditableColumn(column.key)
                        }
                      />
                      <span>可编辑</span>
                    </label>
                  </div>
                );
              })}
            </div>
            <div className="column-modal-footer">
              <button
                type="button"
                className="btn"
                onClick={onCancelPendingFile}
              >
                取消导入
              </button>
              <button
                type="button"
                className="btn btn-primary"
                onClick={onConfirmPendingFile}
              >
                {pendingConfigMode === "edit" ? "保存配置" : "确认并保存配置"}
              </button>
            </div>
          </div>
        </div>
      ) : null}

      {/* ─── AI Config Modal ─── */}
      {isAIConfigModalOpen && activeFile ? (
        <div className="column-modal-mask">
          <div className="column-modal ai-config-modal">
            <h3>AI回答配置</h3>
            <p>{activeFile.fileName}</p>
            {aiConfigFormMessage ? (
              <div className="column-modal-notice">{aiConfigFormMessage}</div>
            ) : null}
            <div className="ai-config-form">
              <label className="ai-config-field">
                <span>配置名称（输入新名称即新增）</span>
                <input
                  type="text"
                  value={draftAIConfigName}
                  onChange={(event) => setDraftAIConfigName(event.target.value)}
                  placeholder="例如：默认配置 / 低成本模型 / 高质量模型"
                  list="ai-config-name-options"
                />
              </label>
              <datalist id="ai-config-name-options">
                {aiConfigList.map((item) => (
                  <option key={item.name} value={item.name} />
                ))}
              </datalist>
              <label className="ai-config-field">
                <span>接口类型</span>
                <select
                  value={draftAIConfig.provider}
                  onChange={(event) =>
                    setDraftAIConfig((previous) => ({
                      ...previous,
                      provider: event.target
                        .value as AIDetectConfig["provider"],
                    }))
                  }
                >
                  {AI_PROVIDER_OPTIONS.map((option) => (
                    <option key={option.value} value={option.value}>
                      {option.label}
                    </option>
                  ))}
                </select>
              </label>
              {draftAIConfig.provider === "openai" ? (
                <label className="ai-config-field">
                  <span>OpenAI兼容接口 URL</span>
                  <input
                    type="text"
                    value={draftAIConfig.url}
                    onChange={(event) =>
                      setDraftAIConfig((previous) => ({
                        ...previous,
                        url: event.target.value,
                      }))
                    }
                    placeholder="例如：https://api.openai.com/v1"
                  />
                </label>
              ) : null}
              {draftAIConfig.provider === "vertex" ? (
                <>
                  <label className="ai-config-field">
                    <span>Vertex Project</span>
                    <input
                      type="text"
                      value={draftAIConfig.vertexProject}
                      onChange={(event) =>
                        setDraftAIConfig((previous) => ({
                          ...previous,
                          vertexProject: event.target.value,
                        }))
                      }
                      placeholder="例如：my-gcp-project"
                    />
                  </label>
                  <label className="ai-config-field">
                    <span>Vertex Location</span>
                    <input
                      type="text"
                      value={draftAIConfig.vertexLocation}
                      onChange={(event) =>
                        setDraftAIConfig((previous) => ({
                          ...previous,
                          vertexLocation: event.target.value,
                        }))
                      }
                      placeholder="例如：us-central1"
                    />
                  </label>
                </>
              ) : null}
              <label className="ai-config-field">
                <span>模型</span>
                <input
                  type="text"
                  value={draftAIConfig.model}
                  onChange={(event) =>
                    setDraftAIConfig((previous) => ({
                      ...previous,
                      model: event.target.value,
                    }))
                  }
                  placeholder="例如：gpt-4.1-mini"
                />
              </label>
              <label className="ai-config-field">
                <span>Reasoning 级别</span>
                <select
                  value={draftAIConfig.reasoningEffort}
                  onChange={(event) =>
                    setDraftAIConfig((previous) => ({
                      ...previous,
                      reasoningEffort: event.target
                        .value as AIDetectConfig["reasoningEffort"],
                    }))
                  }
                >
                  {AI_REASONING_EFFORT_OPTIONS.map((option) => (
                    <option key={option} value={option}>
                      {option}
                    </option>
                  ))}
                </select>
              </label>
              <label className="ai-config-field">
                <span>失败重试次数（后端）</span>
                <input
                  type="number"
                  min={MIN_AI_RETRY_COUNT}
                  max={MAX_AI_RETRY_COUNT}
                  step={1}
                  value={draftAIConfig.retryCount}
                  onChange={(event) =>
                    setDraftAIConfig((previous) => ({
                      ...previous,
                      retryCount: normalizeAIRetryCount(
                        Number(event.target.value),
                      ),
                    }))
                  }
                />
              </label>
              {draftAIConfig.provider === "openai" ? (
                <label className="ai-config-field">
                  <span>OpenAI API Key</span>
                  <input
                    type="password"
                    value={draftAIConfig.apiKey}
                    onChange={(event) =>
                      setDraftAIConfig((previous) => ({
                        ...previous,
                        apiKey: event.target.value,
                      }))
                    }
                    placeholder="请输入 API Key"
                  />
                </label>
              ) : null}
              <label className="ai-config-field">
                <span>结果保存字段</span>
                <select
                  value={draftAIConfig.resultFieldKey}
                  onChange={(event) =>
                    onChangeDraftResultField(event.target.value)
                  }
                >
                  <option value="">请选择</option>
                  {aiResultFieldColumns.map((column) => (
                    <option key={column.key} value={column.key}>
                      {column.title}
                    </option>
                  ))}
                </select>
              </label>
              <div className="ai-config-section">
                <div className="ai-config-section-title">
                  提交回答字段（可多选）
                </div>
                <div className="ai-config-fields">
                  {aiSubmitFieldColumns.map((column) => {
                    const checked = draftAIConfig.submitFieldKeys.includes(
                      column.key,
                    );
                    return (
                      <label key={column.key} className="ai-config-field-item">
                        <input
                          type="checkbox"
                          checked={checked}
                          onChange={() =>
                            onToggleDraftAISubmitField(column.key)
                          }
                        />
                        <span>{column.title}</span>
                      </label>
                    );
                  })}
                </div>
              </div>
              <label className="ai-config-field ai-config-prompt-field">
                <span>
                  Prompt（支持变量 <code>{"{{fields_json}}"}</code> /{" "}
                  <code>{"{{fields_text}}"}</code> /{" "}
                  <code>{"{{image_fields}}"}</code>）
                </span>
                <textarea
                  value={draftAIConfig.prompt}
                  onChange={(event) =>
                    setDraftAIConfig((previous) => ({
                      ...previous,
                      prompt: event.target.value,
                    }))
                  }
                  placeholder="请输入提示词"
                />
              </label>
            </div>
            <div className="column-modal-footer">
              <button
                type="button"
                className="btn"
                onClick={onCancelAIConfigModal}
                disabled={aiConfigSaving}
              >
                取消
              </button>
              <button
                type="button"
                className="btn btn-primary"
                onClick={onSaveAIConfig}
                disabled={aiConfigSaving}
              >
                {aiConfigSaving ? "保存中..." : "保存AI回答配置"}
              </button>
            </div>
          </div>
        </div>
      ) : null}

      {/* ─── Image Lightbox ─── */}
      {previewImageSrc ? (
        <div
          className="lightbox-mask"
          onClick={() => setPreviewImageSrc(null)}
          onKeyDown={(e) => {
            if (e.key === "Escape") setPreviewImageSrc(null);
          }}
          role="button"
          tabIndex={0}
        >
          <img
            className="lightbox-image"
            src={previewImageSrc}
            alt="预览大图"
            onClick={(e) => e.stopPropagation()}
          />
          <button
            type="button"
            className="lightbox-close"
            onClick={() => setPreviewImageSrc(null)}
          >
            ×
          </button>
        </div>
      ) : null}
    </div>
  );
}

export default App;
