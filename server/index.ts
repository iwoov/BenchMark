import cors from "cors";
import express from "express";
import * as ExcelJS from "exceljs";
import multer from "multer";
import { randomUUID } from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { parseWorkbook } from "./excelParser.js";
import {
  getAIDetectConfig,
  deleteFileState,
  getColumnPrefs,
  listFileStates,
  saveAIDetectConfig,
  saveColumnPrefs,
  saveFileState,
} from "./db.js";

type ExcelJSImportLike = { Workbook: new () => ExcelJS.Workbook };
const ExcelJSRuntime: ExcelJSImportLike =
  (ExcelJS as unknown as { default?: ExcelJSImportLike }).default ??
  (ExcelJS as unknown as ExcelJSImportLike);

const app = express();
const port = 8787;

const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 25 * 1024 * 1024,
  },
});

app.use(cors());
app.use(express.json({ limit: "100mb" }));

app.get("/api/health", (_req, res) => {
  res.json({
    ok: true,
    timestamp: new Date().toISOString(),
  });
});

app.get("/api/images/local", (req, res) => {
  const pathQuery = req.query.path;
  if (typeof pathQuery !== "string" || pathQuery.trim().length === 0) {
    return res.status(400).json({ message: "path is required" });
  }

  const absolutePath = toAbsoluteImagePath(pathQuery);
  if (!absolutePath) {
    return res.status(400).json({ message: "path must be an absolute path" });
  }

  const ext = getImageExtFromPathLike(absolutePath);
  if (!ext) {
    return res.status(400).json({ message: "unsupported image extension" });
  }

  try {
    if (!fs.existsSync(absolutePath) || !fs.statSync(absolutePath).isFile()) {
      return res.status(404).json({ message: "image not found" });
    }

    res.status(200);
    res.setHeader("Content-Type", getImageMimeType(ext));
    res.setHeader("Cache-Control", "public, max-age=120");
    const stream = fs.createReadStream(absolutePath);
    stream.on("error", () => {
      if (!res.headersSent) {
        res.status(500).json({ message: "read image failed" });
      } else {
        res.end();
      }
    });
    stream.pipe(res);
    return;
  } catch {
    return res.status(500).json({ message: "read image failed" });
  }
});

function normalizeUploadedFileName(fileName: string): string {
  const decoded = Buffer.from(fileName, "latin1").toString("utf8");
  if (decoded.includes("�")) {
    return fileName;
  }
  return decoded;
}

function isNonEmptyString(value: unknown): value is string {
  return typeof value === "string" && value.trim().length > 0;
}

function normalizeOpenAIUrl(rawUrl: string): string {
  const trimmed = rawUrl.trim();
  if (trimmed.length === 0) {
    return "";
  }

  const normalized = trimmed.replace(/\/+$/, "");
  if (/\/chat\/completions$/i.test(normalized)) {
    return normalized;
  }
  if (/\/v1$/i.test(normalized)) {
    return `${normalized}/chat/completions`;
  }
  return `${normalized}/v1/chat/completions`;
}

const LOCAL_IMAGE_API_PATH = "/api/images/local";
const SUPPORTED_IMAGE_EXTENSIONS = ["png", "jpg", "jpeg", "webp"] as const;
const AI_RESPONSE_LOG_MAX_CHARS = 12000;
const AI_RESPONSE_RAW_LOG_MAX_CHARS = 6000;

function getImageMimeType(ext: string): string {
  const map: Record<string, string> = {
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    webp: "image/webp",
  };
  return map[ext.toLowerCase()] || `image/${ext}`;
}

function getImageExtFromPathLike(pathLike: string): string | null {
  const purePath = pathLike.split(/[?#]/)[0];
  const ext = path.extname(purePath).replace(".", "").toLowerCase();
  return SUPPORTED_IMAGE_EXTENSIONS.includes(
    ext as (typeof SUPPORTED_IMAGE_EXTENSIONS)[number],
  )
    ? ext
    : null;
}

function toAbsoluteImagePath(pathLike: string): string | null {
  const trimmed = pathLike.trim();
  if (!trimmed) {
    return null;
  }

  if (/^file:\/\//i.test(trimmed)) {
    try {
      return fileURLToPath(new URL(trimmed));
    } catch {
      return null;
    }
  }

  if (path.isAbsolute(trimmed) || /^[a-zA-Z]:[\\/]/.test(trimmed)) {
    return trimmed;
  }

  return null;
}

function toDataUrlFromAbsoluteImagePath(imagePath: string): string | null {
  const ext = getImageExtFromPathLike(imagePath);
  if (!ext) {
    return null;
  }

  try {
    if (!fs.existsSync(imagePath) || !fs.statSync(imagePath).isFile()) {
      return null;
    }
    const imageBuffer = fs.readFileSync(imagePath);
    return `data:${getImageMimeType(ext)};base64,${imageBuffer.toString("base64")}`;
  } catch {
    return null;
  }
}

function tryGetPathFromLocalImageApiUrl(imageUrl: string): string | null {
  const trimmed = imageUrl.trim();
  if (!trimmed) {
    return null;
  }

  const parseByUrl = (urlLike: string): string | null => {
    try {
      const url = new URL(urlLike);
      if (url.pathname !== LOCAL_IMAGE_API_PATH) {
        return null;
      }
      const rawPath = url.searchParams.get("path");
      if (!rawPath) {
        return null;
      }
      return toAbsoluteImagePath(rawPath);
    } catch {
      return null;
    }
  };

  if (trimmed.startsWith(LOCAL_IMAGE_API_PATH)) {
    return parseByUrl(`http://localhost${trimmed}`);
  }

  if (/^https?:\/\//i.test(trimmed)) {
    return parseByUrl(trimmed);
  }

  return null;
}

function normalizeImageUrlForAI(imageUrl: string): string | null {
  const trimmed = imageUrl.trim();
  if (!trimmed) {
    return null;
  }

  if (/^data:image\//i.test(trimmed)) {
    return trimmed;
  }

  const localPathFromApi = tryGetPathFromLocalImageApiUrl(trimmed);
  if (localPathFromApi) {
    return toDataUrlFromAbsoluteImagePath(localPathFromApi);
  }

  const absolutePath = toAbsoluteImagePath(trimmed);
  if (absolutePath) {
    return toDataUrlFromAbsoluteImagePath(absolutePath);
  }

  if (/^https?:\/\//i.test(trimmed)) {
    return trimmed;
  }

  return null;
}

function logAIResponseById(requestId: string, text: string): void {
  const normalized = text.replace(/\r/g, "");
  if (normalized.length <= AI_RESPONSE_LOG_MAX_CHARS) {
    // eslint-disable-next-line no-console
    console.log(
      `[AIResponse][${requestId}] len=${normalized.length}\n${normalized}`,
    );
    return;
  }

  // eslint-disable-next-line no-console
  console.log(
    `[AIResponse][${requestId}] len=${normalized.length} truncated=${AI_RESPONSE_LOG_MAX_CHARS}\n${normalized.slice(0, AI_RESPONSE_LOG_MAX_CHARS)}\n...[truncated]`,
  );
}

function logAIRawResponseById(requestId: string, text: string): void {
  const normalized = text.replace(/\r/g, "");
  if (normalized.length <= AI_RESPONSE_RAW_LOG_MAX_CHARS) {
    // eslint-disable-next-line no-console
    console.log(
      `[AIResponseRaw][${requestId}] len=${normalized.length}\n${normalized}`,
    );
    return;
  }

  // eslint-disable-next-line no-console
  console.log(
    `[AIResponseRaw][${requestId}] len=${normalized.length} truncated=${AI_RESPONSE_RAW_LOG_MAX_CHARS}\n${normalized.slice(0, AI_RESPONSE_RAW_LOG_MAX_CHARS)}\n...[truncated]`,
  );
}

function logAIThinkingById(requestId: string, text: string): void {
  const normalized = text.replace(/\r/g, "");
  if (normalized.length <= AI_RESPONSE_LOG_MAX_CHARS) {
    // eslint-disable-next-line no-console
    console.log(
      `[AIThinking][${requestId}] len=${normalized.length}\n${normalized}`,
    );
    return;
  }

  // eslint-disable-next-line no-console
  console.log(
    `[AIThinking][${requestId}] len=${normalized.length} truncated=${AI_RESPONSE_LOG_MAX_CHARS}\n${normalized.slice(0, AI_RESPONSE_LOG_MAX_CHARS)}\n...[truncated]`,
  );
}

type AIClientStreamEvent =
  | {
      type: "answer" | "thinking";
      text: string;
    }
  | {
      type: "done";
    };

function writeAIClientStreamEvent(
  res: express.Response,
  event: AIClientStreamEvent,
): void {
  res.write(`${JSON.stringify(event)}\n`);
}

function asRecord(value: unknown): Record<string, unknown> | null {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return null;
  }
  return value as Record<string, unknown>;
}

function readTextValue(value: unknown): string {
  if (typeof value === "string") {
    return value;
  }
  if (Array.isArray(value)) {
    const chunks = value
      .map((item) => readTextValue(item))
      .filter((item) => item.length > 0);
    return chunks.join("");
  }
  const objectValue = asRecord(value);
  if (!objectValue) {
    return "";
  }
  if (typeof objectValue.text === "string") {
    return objectValue.text;
  }
  if (typeof objectValue.delta === "string") {
    return objectValue.delta;
  }
  if (typeof objectValue.content === "string") {
    return objectValue.content;
  }
  if (Array.isArray(objectValue.content)) {
    return readTextValue(objectValue.content);
  }
  return "";
}

function extractContentParts(
  content: unknown,
  answerChunks: string[],
  thinkingChunks: string[],
): void {
  if (typeof content === "string") {
    if (content.length > 0) {
      answerChunks.push(content);
    }
    return;
  }
  if (!Array.isArray(content)) {
    return;
  }
  for (const part of content) {
    const partRecord = asRecord(part);
    if (!partRecord) {
      continue;
    }
    const type =
      typeof partRecord.type === "string" ? partRecord.type.toLowerCase() : "";
    const text = readTextValue(partRecord);
    if (text.length === 0) {
      continue;
    }
    if (type.includes("reasoning") || type.includes("thinking")) {
      thinkingChunks.push(text);
      continue;
    }
    answerChunks.push(text);
  }
}

function extractStreamTextPayload(payload: unknown): {
  answerText: string;
  thinkingText: string;
} {
  const root = asRecord(payload);
  if (!root) {
    return { answerText: "", thinkingText: "" };
  }

  const answerChunks: string[] = [];
  const thinkingChunks: string[] = [];

  const eventType =
    typeof root.type === "string" ? root.type.toLowerCase() : "";
  const topDelta = readTextValue(root.delta);
  const topText = readTextValue(root.text);
  if (topDelta.length > 0) {
    if (eventType.includes("reasoning") || eventType.includes("thinking")) {
      thinkingChunks.push(topDelta);
    } else if (
      eventType.includes("output_text") ||
      eventType.includes("response.text") ||
      eventType.includes(".delta")
    ) {
      answerChunks.push(topDelta);
    }
  }
  if (topText.length > 0) {
    if (eventType.includes("reasoning") || eventType.includes("thinking")) {
      thinkingChunks.push(topText);
    } else if (eventType.includes("output_text")) {
      answerChunks.push(topText);
    }
  }

  const choices = Array.isArray(root.choices) ? root.choices : [];
  const firstChoice = asRecord(choices[0]);
  if (firstChoice) {
    const delta = asRecord(firstChoice.delta);
    if (delta) {
      extractContentParts(delta.content, answerChunks, thinkingChunks);
      const deltaText = readTextValue(delta.content);
      if (!Array.isArray(delta.content) && deltaText.length > 0) {
        answerChunks.push(deltaText);
      }
      const deltaThinking =
        readTextValue(delta.reasoning_content) ||
        readTextValue(delta.reasoning) ||
        readTextValue(delta.thinking);
      if (deltaThinking.length > 0) {
        thinkingChunks.push(deltaThinking);
      }
    }

    const message = asRecord(firstChoice.message);
    if (message) {
      extractContentParts(message.content, answerChunks, thinkingChunks);
      const messageText = readTextValue(message.content);
      if (!Array.isArray(message.content) && messageText.length > 0) {
        answerChunks.push(messageText);
      }
      const messageThinking =
        readTextValue(message.reasoning) || readTextValue(message.thinking);
      if (messageThinking.length > 0) {
        thinkingChunks.push(messageThinking);
      }
    }
  }

  const response = asRecord(root.response);
  if (response) {
    const outputs = Array.isArray(response.output) ? response.output : [];
    for (const output of outputs) {
      const outputRecord = asRecord(output);
      if (!outputRecord) {
        continue;
      }
      extractContentParts(outputRecord.content, answerChunks, thinkingChunks);
    }
  }

  const outputText = root.output_text;
  if (typeof outputText === "string" && outputText.length > 0) {
    answerChunks.push(outputText);
  } else if (Array.isArray(outputText)) {
    for (const item of outputText) {
      const text = readTextValue(item);
      if (text.length > 0) {
        answerChunks.push(text);
      }
    }
  }

  return {
    answerText: answerChunks.join(""),
    thinkingText: thinkingChunks.join(""),
  };
}

type AIDetectField = {
  title: string;
  type: "text" | "image";
  value: string;
  imageUrl?: string;
};

type OpenAIMessageContentPart =
  | {
      type: "text";
      text: string;
    }
  | {
      type: "image_url";
      image_url: { url: string };
    };

type PromptBuildResult = {
  promptText: string;
  imageFields: Array<{ title: string; value: string; imageUrl: string }>;
};

function toAIDetectFields(value: unknown): AIDetectField[] | null {
  if (!Array.isArray(value)) {
    return null;
  }

  const result: AIDetectField[] = [];
  for (const item of value) {
    if (!item || typeof item !== "object" || Array.isArray(item)) {
      return null;
    }

    const candidate = item as {
      title?: unknown;
      type?: unknown;
      value?: unknown;
      imageUrl?: unknown;
    };

    if (
      typeof candidate.title !== "string" ||
      candidate.title.trim().length === 0
    ) {
      return null;
    }
    if (candidate.type !== "text" && candidate.type !== "image") {
      return null;
    }
    if (candidate.value !== undefined && typeof candidate.value !== "string") {
      return null;
    }
    if (candidate.type === "image") {
      if (
        typeof candidate.imageUrl !== "string" ||
        candidate.imageUrl.trim().length === 0
      ) {
        return null;
      }
      result.push({
        title: candidate.title.trim(),
        type: "image",
        value: typeof candidate.value === "string" ? candidate.value : "",
        imageUrl: candidate.imageUrl,
      });
      continue;
    }

    result.push({
      title: candidate.title.trim(),
      type: "text",
      value: typeof candidate.value === "string" ? candidate.value : "",
    });
  }

  return result;
}

function buildPromptContent(
  prompt: string,
  fields: AIDetectField[],
): PromptBuildResult {
  const fieldSummary: Record<string, string> = {};
  const imageFields: Array<{ title: string; value: string; imageUrl: string }> =
    [];

  fields.forEach((field) => {
    if (field.type === "image" && field.imageUrl) {
      const summary =
        field.value.trim().length > 0 ? `[图片] ${field.value}` : "[图片]";
      fieldSummary[field.title] = summary;
      imageFields.push({
        title: field.title,
        value: field.value,
        imageUrl: field.imageUrl,
      });
      return;
    }
    fieldSummary[field.title] = field.value;
  });

  const fieldsJson = JSON.stringify(fieldSummary, null, 2);
  const fieldsText = Object.entries(fieldSummary)
    .map(([key, value]) => `${key}: ${value || "-"}`)
    .join("\n");
  const imageFieldsText =
    imageFields.length > 0
      ? imageFields
          .map((field) =>
            field.value.trim().length > 0
              ? `${field.title}（说明：${field.value}）`
              : field.title,
          )
          .join("、")
      : "无";

  const hasJsonPlaceholder = prompt.includes("{{fields_json}}");
  const hasTextPlaceholder = prompt.includes("{{fields_text}}");
  const hasImagePlaceholder = prompt.includes("{{image_fields}}");

  const mergedPrompt = prompt
    .replaceAll("{{fields_json}}", fieldsJson)
    .replaceAll("{{fields_text}}", fieldsText)
    .replaceAll("{{image_fields}}", imageFieldsText);

  if (hasJsonPlaceholder || hasTextPlaceholder) {
    return {
      promptText: mergedPrompt,
      imageFields,
    };
  }

  const withImageHint =
    hasImagePlaceholder || imageFields.length === 0
      ? mergedPrompt
      : `${mergedPrompt}\n\n图片字段：${imageFieldsText}`;

  return {
    promptText: `${withImageHint.trim()}\n\n待检测字段(JSON):\n${fieldsJson}`,
    imageFields,
  };
}

app.post("/api/files/upload", upload.single("file"), async (req, res) => {
  try {
    const file = req.file;
    if (!file) {
      return res.status(400).json({
        message: "请先选择 Excel 文件",
      });
    }

    const fileId = randomUUID();
    const normalizedFileName = normalizeUploadedFileName(file.originalname);
    const parsed = await parseWorkbook(file.buffer, normalizedFileName, fileId);
    return res.json(parsed);
  } catch (error) {
    const message = error instanceof Error ? error.message : "解析 Excel 失败";
    return res.status(400).json({ message });
  }
});

app.get("/api/files", (_req, res) => {
  const files = listFileStates().map((item) => item.state);
  return res.json({ files });
});

app.put("/api/files/:fileId/state", (req, res) => {
  const { fileId } = req.params;
  const { state } = req.body as { state?: unknown };

  if (!state || typeof state !== "object") {
    return res.status(400).json({ message: "state must be an object" });
  }

  const nextState = state as { fileId?: unknown; fileName?: unknown };
  if (typeof nextState.fileId !== "string" || nextState.fileId !== fileId) {
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

  saveFileState(fileId, nextState.fileName, state);
  return res.json({ ok: true });
});

app.delete("/api/files/:fileId", (req, res) => {
  const { fileId } = req.params;
  deleteFileState(fileId);
  return res.json({ ok: true });
});

// ─── Column Preferences ───

app.get("/api/column-prefs/:fileName", (req, res) => {
  const { fileName } = req.params;
  const config = getColumnPrefs(decodeURIComponent(fileName));
  return res.json({ config });
});

app.put("/api/column-prefs/:fileName", (req, res) => {
  const { fileName } = req.params;
  const { fieldSignature, displayKeys, editableKeys } = req.body as {
    fieldSignature: unknown;
    displayKeys: unknown;
    editableKeys: unknown;
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

  saveColumnPrefs(decodeURIComponent(fileName), {
    fieldSignature,
    displayKeys,
    editableKeys,
  });
  return res.json({ ok: true });
});

// ─── AI Detection Config ───

app.get("/api/ai-config/:fileName", (req, res) => {
  const { fileName } = req.params;
  const config = getAIDetectConfig(decodeURIComponent(fileName));
  return res.json({ config });
});

app.put("/api/ai-config/:fileName", (req, res) => {
  const { fileName } = req.params;
  const { url, model, apiKey, submitFieldKeys, prompt, resultFieldKey } =
    req.body as {
      url: unknown;
      model: unknown;
      apiKey: unknown;
      submitFieldKeys: unknown;
      prompt: unknown;
      resultFieldKey?: unknown;
    };

  if (!isNonEmptyString(url)) {
    return res.status(400).json({ message: "url must be a non-empty string" });
  }
  if (!isNonEmptyString(model)) {
    return res
      .status(400)
      .json({ message: "model must be a non-empty string" });
  }
  if (!isNonEmptyString(apiKey)) {
    return res
      .status(400)
      .json({ message: "apiKey must be a non-empty string" });
  }
  if (
    !Array.isArray(submitFieldKeys) ||
    !submitFieldKeys.every((item) => typeof item === "string")
  ) {
    return res
      .status(400)
      .json({ message: "submitFieldKeys must be a string array" });
  }
  if (!isNonEmptyString(prompt)) {
    return res
      .status(400)
      .json({ message: "prompt must be a non-empty string" });
  }
  if (resultFieldKey !== undefined && typeof resultFieldKey !== "string") {
    return res.status(400).json({ message: "resultFieldKey must be a string" });
  }

  saveAIDetectConfig(decodeURIComponent(fileName), {
    url,
    model,
    apiKey,
    submitFieldKeys,
    prompt,
    resultFieldKey: typeof resultFieldKey === "string" ? resultFieldKey : "",
  });

  return res.json({ ok: true });
});

// ─── AI Detection Stream ───

app.post("/api/ai-detect/stream", async (req, res) => {
  const { url, model, apiKey, prompt, fields } = req.body as {
    url: unknown;
    model: unknown;
    apiKey: unknown;
    prompt: unknown;
    fields: unknown;
  };

  if (!isNonEmptyString(url)) {
    return res.status(400).json({ message: "url must be a non-empty string" });
  }
  if (!isNonEmptyString(model)) {
    return res
      .status(400)
      .json({ message: "model must be a non-empty string" });
  }
  if (!isNonEmptyString(apiKey)) {
    return res
      .status(400)
      .json({ message: "apiKey must be a non-empty string" });
  }
  if (!isNonEmptyString(prompt)) {
    return res
      .status(400)
      .json({ message: "prompt must be a non-empty string" });
  }

  const fieldPayload = toAIDetectFields(fields);
  if (!fieldPayload || fieldPayload.length === 0) {
    return res
      .status(400)
      .json({ message: "fields must be a non-empty array" });
  }

  const aiRequestId = randomUUID().slice(0, 8);
  const startedAt = Date.now();
  const elapsedMs = (): number => Date.now() - startedAt;

  const aiFields = fieldPayload.map((field): AIDetectField => {
    if (field.type !== "image" || !field.imageUrl) {
      return field;
    }

    const normalizedImageUrl = normalizeImageUrlForAI(field.imageUrl);
    if (normalizedImageUrl) {
      return {
        ...field,
        imageUrl: normalizedImageUrl,
      };
    }

    const fallbackValue =
      field.value.trim().length > 0
        ? `${field.value}\n[图片读取失败: ${field.imageUrl}]`
        : `[图片读取失败: ${field.imageUrl}]`;
    return {
      title: field.title,
      type: "text",
      value: fallbackValue,
    };
  });

  const aiFieldLogs = aiFields.map((field) => ({
    title: field.title,
    type: field.type,
    valuePreview: field.value.slice(0, 80),
    imageStatus:
      field.type === "image" && field.imageUrl
        ? field.imageUrl.startsWith("data:image/")
          ? "data-url"
          : "remote-url"
        : undefined,
  }));
  // eslint-disable-next-line no-console
  console.log(
    `[AIRequest][${aiRequestId}] model=${model} fields=${aiFields.length} images=${aiFields.filter((item) => item.type === "image").length} texts=${aiFields.filter((item) => item.type === "text").length}`,
  );
  // eslint-disable-next-line no-console
  console.log(
    `[AIRequestFields][${aiRequestId}] ${JSON.stringify(aiFieldLogs)}`,
  );

  const targetUrl = normalizeOpenAIUrl(url);
  try {
    // Validate URL before dispatching request.
    new URL(targetUrl);
  } catch {
    return res.status(400).json({ message: "url is invalid" });
  }

  const controller = new AbortController();
  let abortReason = "";
  let upstreamStatusCode: number | null = null;
  let streamChunkCount = 0;
  let streamTextLength = 0;
  let streamThinkingChunkCount = 0;
  let streamThinkingTextLength = 0;
  let doneByDoneToken = false;
  let doneByNaturalEnd = false;

  const abortUpstream = (reason: string): void => {
    if (controller.signal.aborted) {
      return;
    }
    abortReason = reason;
    // eslint-disable-next-line no-console
    console.log(
      `[AIAbort][${aiRequestId}] reason=${reason} elapsedMs=${elapsedMs()} reqAborted=${req.aborted} reqComplete=${req.complete} resEnded=${res.writableEnded} headersSent=${res.headersSent}`,
    );
    controller.abort();
  };

  req.on("aborted", () => {
    // eslint-disable-next-line no-console
    console.log(
      `[AIConn][${aiRequestId}] req.aborted elapsedMs=${elapsedMs()} reqComplete=${req.complete}`,
    );
    abortUpstream("req.aborted");
  });
  req.on("close", () => {
    // eslint-disable-next-line no-console
    console.log(
      `[AIConn][${aiRequestId}] req.close elapsedMs=${elapsedMs()} reqAborted=${req.aborted} reqComplete=${req.complete}`,
    );
    if (req.aborted) {
      abortUpstream("req.close(aborted)");
    }
  });
  res.on("finish", () => {
    // eslint-disable-next-line no-console
    console.log(
      `[AIConn][${aiRequestId}] res.finish elapsedMs=${elapsedMs()} status=${res.statusCode} upstreamStatus=${upstreamStatusCode ?? "-"} chunks=${streamChunkCount} chars=${streamTextLength} thinkingChunks=${streamThinkingChunkCount} thinkingChars=${streamThinkingTextLength} doneToken=${doneByDoneToken} naturalEnd=${doneByNaturalEnd}`,
    );
  });
  res.on("close", () => {
    // eslint-disable-next-line no-console
    console.log(
      `[AIConn][${aiRequestId}] res.close elapsedMs=${elapsedMs()} resEnded=${res.writableEnded} writableFinished=${res.writableFinished} chunks=${streamChunkCount} chars=${streamTextLength} thinkingChunks=${streamThinkingChunkCount} thinkingChars=${streamThinkingTextLength}`,
    );
    if (!res.writableEnded) {
      abortUpstream("res.close(before-end)");
    }
  });

  try {
    const promptContent = buildPromptContent(prompt, aiFields);
    let aiResponseText = "";
    let aiThinkingText = "";
    const userContent: string | OpenAIMessageContentPart[] =
      promptContent.imageFields.length > 0
        ? [
            {
              type: "text",
              text: promptContent.promptText,
            },
            ...promptContent.imageFields.flatMap((field) => {
              const imageLabel =
                field.value.trim().length > 0
                  ? `字段图片：${field.title}（说明：${field.value}）`
                  : `字段图片：${field.title}`;
              return [
                {
                  type: "text" as const,
                  text: imageLabel,
                },
                {
                  type: "image_url" as const,
                  image_url: {
                    url: field.imageUrl,
                  },
                },
              ];
            }),
          ]
        : promptContent.promptText;

    // eslint-disable-next-line no-console
    console.log(
      `[AIUpstream][${aiRequestId}] dispatch elapsedMs=${elapsedMs()} url=${targetUrl}`,
    );
    const upstream = await fetch(targetUrl, {
      method: "POST",
      signal: controller.signal,
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify({
        model,
        stream: true,
        messages: [{ role: "user", content: userContent }],
      }),
    });
    upstreamStatusCode = upstream.status;
    // eslint-disable-next-line no-console
    console.log(
      `[AIUpstream][${aiRequestId}] connected elapsedMs=${elapsedMs()} status=${upstream.status} ok=${upstream.ok} hasBody=${Boolean(upstream.body)}`,
    );

    if (!upstream.ok || !upstream.body) {
      const rawText = await upstream.text().catch(() => "");
      let message = "AI 检测请求失败";
      if (rawText.length > 0) {
        try {
          const payload = JSON.parse(rawText) as {
            error?: { message?: string };
            message?: string;
          };
          message = payload.error?.message ?? payload.message ?? message;
        } catch {
          message = rawText.slice(0, 400);
        }
      }
      // eslint-disable-next-line no-console
      console.log(
        `[AIResponseError][${aiRequestId}] status=${upstream.status || 500} message=${message}`,
      );
      return res.status(upstream.status || 500).json({ message });
    }

    res.status(200);
    res.setHeader("Content-Type", "application/x-ndjson; charset=utf-8");
    res.setHeader("Cache-Control", "no-cache, no-transform");
    res.setHeader("Connection", "keep-alive");
    res.setHeader("X-Accel-Buffering", "no");
    res.flushHeaders();

    const decoder = new TextDecoder();
    let buffer = "";
    let rawStreamPreview = "";

    const reader = upstream.body.getReader();
    while (true) {
      const { value, done } = await reader.read();
      if (done) {
        break;
      }
      if (!value) {
        continue;
      }

      const current = decoder.decode(value, { stream: true });
      if (rawStreamPreview.length < AI_RESPONSE_RAW_LOG_MAX_CHARS * 2) {
        rawStreamPreview += current;
      }
      buffer += current;

      const lines = buffer.split(/\r?\n/);
      buffer = lines.pop() ?? "";

      for (const line of lines) {
        const trimmed = line.trim();
        if (!trimmed.startsWith("data:")) {
          continue;
        }

        const data = trimmed.slice(5).trim();
        if (data === "[DONE]") {
          doneByDoneToken = true;
          logAIResponseById(aiRequestId, aiResponseText);
          if (aiThinkingText.trim().length > 0) {
            logAIThinkingById(aiRequestId, aiThinkingText);
          }
          writeAIClientStreamEvent(res, { type: "done" });
          res.end();
          return;
        }
        if (data.length === 0) {
          continue;
        }

        try {
          const payload = JSON.parse(data) as unknown;
          const extracted = extractStreamTextPayload(payload);
          if (extracted.thinkingText.length > 0) {
            aiThinkingText += extracted.thinkingText;
            streamThinkingChunkCount += 1;
            streamThinkingTextLength += extracted.thinkingText.length;
            writeAIClientStreamEvent(res, {
              type: "thinking",
              text: extracted.thinkingText,
            });
          }
          if (extracted.answerText.length > 0) {
            aiResponseText += extracted.answerText;
            streamChunkCount += 1;
            streamTextLength += extracted.answerText.length;
            writeAIClientStreamEvent(res, {
              type: "answer",
              text: extracted.answerText,
            });
          }
        } catch {
          // Ignore non-JSON stream chunks.
        }
      }
    }

    buffer += decoder.decode();
    if (rawStreamPreview.length < AI_RESPONSE_RAW_LOG_MAX_CHARS * 2) {
      rawStreamPreview += buffer;
    }

    if (buffer.length > 0 && buffer.includes("data:")) {
      const maybeData = buffer
        .split(/\r?\n/)
        .map((line) => line.trim())
        .find((line) => line.startsWith("data:"));
      const value = maybeData ? maybeData.slice(5).trim() : "";
      if (value && value !== "[DONE]") {
        try {
          const payload = JSON.parse(value) as unknown;
          const extracted = extractStreamTextPayload(payload);
          if (extracted.thinkingText.length > 0) {
            aiThinkingText += extracted.thinkingText;
            streamThinkingChunkCount += 1;
            streamThinkingTextLength += extracted.thinkingText.length;
            writeAIClientStreamEvent(res, {
              type: "thinking",
              text: extracted.thinkingText,
            });
          }
          if (extracted.answerText.length > 0) {
            aiResponseText += extracted.answerText;
            streamChunkCount += 1;
            streamTextLength += extracted.answerText.length;
            writeAIClientStreamEvent(res, {
              type: "answer",
              text: extracted.answerText,
            });
          }
        } catch {
          // Ignore trailing invalid chunk.
        }
      }
    }

    doneByNaturalEnd = true;
    logAIResponseById(aiRequestId, aiResponseText);
    if (aiThinkingText.trim().length > 0) {
      logAIThinkingById(aiRequestId, aiThinkingText);
    }
    if (
      aiResponseText.trim().length === 0 &&
      rawStreamPreview.trim().length > 0
    ) {
      logAIRawResponseById(aiRequestId, rawStreamPreview);
    }
    writeAIClientStreamEvent(res, { type: "done" });
    res.end();
    return;
  } catch (error) {
    if (controller.signal.aborted) {
      // eslint-disable-next-line no-console
      console.log(
        `[AIResponseAborted][${aiRequestId}] reason=${abortReason || "unknown"} elapsedMs=${elapsedMs()} upstreamStatus=${upstreamStatusCode ?? "-"} reqAborted=${req.aborted} reqComplete=${req.complete} resEnded=${res.writableEnded} chunks=${streamChunkCount} chars=${streamTextLength} thinkingChunks=${streamThinkingChunkCount} thinkingChars=${streamThinkingTextLength}`,
      );
      return;
    }
    const message = error instanceof Error ? error.message : "AI 检测请求失败";
    // eslint-disable-next-line no-console
    console.log(`[AIResponseException][${aiRequestId}] ${message}`);
    return res.status(500).json({ message });
  }
});

app.post("/api/files/export", async (req, res) => {
  const { fileName, headers, rows } = req.body as {
    fileName: unknown;
    headers: unknown;
    rows: unknown;
  };

  if (typeof fileName !== "string" || fileName.trim().length === 0) {
    return res
      .status(400)
      .json({ message: "fileName must be a non-empty string" });
  }
  if (
    !Array.isArray(headers) ||
    !headers.every((item) => typeof item === "string")
  ) {
    return res.status(400).json({ message: "headers must be a string array" });
  }
  if (
    !Array.isArray(rows) ||
    !rows.every(
      (row) =>
        Array.isArray(row) && row.every((cell) => typeof cell === "string"),
    )
  ) {
    return res.status(400).json({ message: "rows must be a 2d string array" });
  }

  try {
    const workbook = new ExcelJSRuntime.Workbook();
    const worksheet = workbook.addWorksheet("Sheet1");

    worksheet.addRow(headers);
    for (const row of rows) {
      worksheet.addRow(row);
    }

    worksheet.columns = headers.map((header, index) => {
      const maxLengthFromRows = rows.reduce((acc, row) => {
        const value = row[index] ?? "";
        return Math.max(acc, value.length);
      }, 0);
      return {
        header,
        key: `col_${index}`,
        width: Math.min(
          60,
          Math.max(12, Math.max(header.length, maxLengthFromRows) + 2),
        ),
      };
    });

    const headerRow = worksheet.getRow(1);
    headerRow.font = { bold: true };
    headerRow.commit();

    const baseName = fileName.replace(/\.[^.]+$/, "");
    const exportName = `${baseName}-导出.xlsx`;
    const encodedFileName = encodeURIComponent(exportName);
    const buffer = await workbook.xlsx.writeBuffer();

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    );
    res.setHeader(
      "Content-Disposition",
      `attachment; filename*=UTF-8''${encodedFileName}`,
    );
    const outputBuffer = Buffer.isBuffer(buffer)
      ? buffer
      : Buffer.from(buffer as ArrayBuffer);
    return res.send(outputBuffer);
  } catch (error) {
    const message = error instanceof Error ? error.message : "导出 Excel 失败";
    return res.status(500).json({ message });
  }
});

app.listen(port, () => {
  // eslint-disable-next-line no-console
  console.log(`Server running at http://localhost:${port}`);
});
