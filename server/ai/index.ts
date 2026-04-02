import express from "express";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { randomUUID } from "node:crypto";
import {
    findAIModelRouteByName,
    findAIProviderEndpointByName,
    getFileAICleaningConfig,
    getFileAIChatConfig,
    getFileAIStageConfigs,
    listAIModelRoutes,
    listAIProviderEndpoints,
    saveAIModelRoutes,
    saveFileAICleaningConfig,
    saveFileAIChatConfig,
    saveAIProviderEndpoints,
    saveFileAIStageConfigs,
    type AICleaningToolKey,
    type AIModelRouteConfig,
    type AIProviderApiType,
    type AIProviderEndpointConfig,
    type FileAICleaningConfigMap,
    type FileAICleaningToolConfig,
    type FileAIChatConfig,
    type FileAIStageConfig,
} from "../db.js";

function isNonEmptyString(value: unknown): value is string {
    return typeof value === "string" && value.trim().length > 0;
}

type AIReasoningEffort = "low" | "medium" | "high";
type AIProvider = "openai" | "gemini" | "anthropic";
type AIDetectStageKey =
    | "precheck"
    | "context_audit"
    | "independent_solving"
    | "final_verdict";
type AICleaningOutputMapping = {
    outputKey: string;
    targetFieldKey: string;
};
type AIRouteStep = {
    providerName: string;
};
type AIRoute = {
    name: string;
    model: string;
    reasoningEffort: AIReasoningEffort;
    retryCount: number;
    steps: AIRouteStep[];
};
type AIResolvedRouteStep = {
    provider: AIProviderEndpointConfig;
    route: AIRoute;
};
const AI_STAGE_ORDER: AIDetectStageKey[] = [
    "precheck",
    "context_audit",
    "independent_solving",
    "final_verdict",
];
const AI_STAGE_LABELS: Record<AIDetectStageKey, string> = {
    precheck: "Pre-check",
    context_audit: "Context Audit",
    independent_solving: "Independent Solving",
    final_verdict: "Final Verdict",
};
const AI_CLEANING_TOOL_ORDER: AICleaningToolKey[] = [
    "generate_level3_tags",
    "biochem_level1_refine",
];
const AI_CLEANING_TOOL_LABELS: Record<AICleaningToolKey, string> = {
    generate_level3_tags: "生成 level3 标签",
    biochem_level1_refine: "细分生化 level1",
};
const AI_CLEANING_TOOL_OUTPUT_KEYS: Record<AICleaningToolKey, string[]> = {
    generate_level3_tags: [
        "representation_method",
        "representation_type",
        "tags",
    ],
    biochem_level1_refine: ["discipline", "confidence", "reason"],
};
const DEFAULT_AI_RETRY_COUNT = 5;
const MIN_AI_RETRY_COUNT = 0;
const MAX_AI_RETRY_COUNT = 10;

function isAIReasoningEffort(value: unknown): value is AIReasoningEffort {
    return value === "low" || value === "medium" || value === "high";
}

function isAIProviderApiType(value: unknown): value is AIProvider {
    return value === "openai" || value === "gemini" || value === "anthropic";
}

function isOpenAICompatibleProvider(provider: AIProvider): boolean {
    return provider === "openai";
}

function isGeminiCompatibleProvider(provider: AIProvider): boolean {
    return provider === "gemini";
}

function isAnthropicCompatibleProvider(provider: AIProvider): boolean {
    return provider === "anthropic";
}

function isValidAIRetryCount(value: unknown): value is number {
    return (
        typeof value === "number" &&
        Number.isInteger(value) &&
        value >= MIN_AI_RETRY_COUNT &&
        value <= MAX_AI_RETRY_COUNT
    );
}

function normalizeAIRetryCount(value: unknown): number {
    if (isValidAIRetryCount(value)) {
        return value;
    }
    return DEFAULT_AI_RETRY_COUNT;
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

const MODEL_PROVIDER_PREFIXES = new Set([
    "openai",
    "google",
    "anthropic",
    "gemini",
    "vertex",
    "idealab",
]);

function normalizeModelName(model: string): string {
    const trimmed = model.trim();
    if (!trimmed) {
        return "";
    }
    const slashIndex = trimmed.indexOf("/");
    if (slashIndex <= 0) {
        return trimmed;
    }
    const prefix = trimmed.slice(0, slashIndex).toLowerCase();
    if (!MODEL_PROVIDER_PREFIXES.has(prefix)) {
        return trimmed;
    }
    const remainder = trimmed.slice(slashIndex + 1).trim();
    return remainder.length > 0 ? remainder : trimmed;
}

function appendQueryParam(url: URL, key: string, value: string): void {
    if (!url.searchParams.has(key)) {
        url.searchParams.set(key, value);
    }
}

function normalizeGeminiEndpoint(rawUrl: string, model: string): string {
    const trimmed = rawUrl.trim();
    if (trimmed.length === 0) {
        return "";
    }

    const withModel = trimmed.replaceAll(
        "{{model}}",
        encodeURIComponent(model),
    );
    let endpoint = withModel;
    if (!endpoint.includes(":streamGenerateContent")) {
        if (endpoint.includes("/models/")) {
            endpoint = endpoint.replace(/\/+$/, "");
            endpoint = `${endpoint}:streamGenerateContent`;
        } else {
            endpoint = endpoint.replace(/\/+$/, "");
            if (endpoint.endsWith("/v1") || endpoint.endsWith("/v1beta")) {
                endpoint = `${endpoint}/models/${encodeURIComponent(model)}:streamGenerateContent`;
            } else {
                endpoint = `${endpoint}/models/${encodeURIComponent(model)}:streamGenerateContent`;
            }
        }
    }

    try {
        const parsed = new URL(endpoint);
        appendQueryParam(parsed, "alt", "sse");
        return parsed.toString();
    } catch {
        return endpoint;
    }
}

function stripBearerPrefix(value: string): string {
    return value
        .trim()
        .replace(/^Bearer\s+/i, "")
        .trim();
}

function buildGeminiAuthHeaders(
    endpointUrl: string,
    apiKeyOrToken: string,
): {
    headers: Record<string, string>;
    mode: "x-goog-api-key" | "x-goog-api-key-bearer" | "authorization-bearer";
} {
    const token = stripBearerPrefix(apiKeyOrToken);
    const baseHeaders: Record<string, string> = {
        "Content-Type": "application/json",
    };

    // Explicit bearer token input always uses Authorization.
    if (/^Bearer\s+/i.test(apiKeyOrToken.trim())) {
        return {
            headers: {
                ...baseHeaders,
                Authorization: `Bearer ${token}`,
            },
            mode: "authorization-bearer",
        };
    }

    try {
        const parsed = new URL(endpointUrl);
        const hostname = parsed.hostname.toLowerCase();
        const pathname = parsed.pathname.toLowerCase();
        const shouldUseModelRouterApiKeyHeader =
            hostname === "routify.alibaba-inc.com" &&
            pathname.includes("/protocol/vertex/");
        if (shouldUseModelRouterApiKeyHeader) {
            return {
                headers: {
                    ...baseHeaders,
                    "x-goog-api-key": `Bearer ${token}`,
                },
                mode: "x-goog-api-key-bearer",
            };
        }
        const shouldUseBearer =
            pathname.includes("/api/vertex/") ||
            !hostname.endsWith("googleapis.com");
        if (shouldUseBearer) {
            return {
                headers: {
                    ...baseHeaders,
                    Authorization: `Bearer ${token}`,
                },
                mode: "authorization-bearer",
            };
        }
    } catch {
        // Keep fallback branch below.
    }

    return {
        headers: {
            ...baseHeaders,
            "x-goog-api-key": token,
        },
        mode: "x-goog-api-key",
    };
}

const LOCAL_IMAGE_API_PATH = "/api/images/local";
const SUPPORTED_IMAGE_EXTENSIONS = ["png", "jpg", "jpeg", "webp"] as const;
const AI_RESPONSE_LOG_MAX_CHARS = 12000;
const AI_RESPONSE_RAW_LOG_MAX_CHARS = 6000;
const SHOULD_LOG_AI_VERBOSE = process.env.DEBUG_AI_VERBOSE === "1";
const SHOULD_LOG_AI_THINKING =
    process.env.DEBUG_AI_THINKING === "1" || SHOULD_LOG_AI_VERBOSE;

function formatLogTimestamp(timestamp: number): string {
    return new Date(timestamp).toISOString();
}

function summarizeAIResponseForLog(text: string): string {
    const normalized = text.replace(/\r/g, "").trim();
    if (!normalized) {
        return "shape=empty";
    }

    try {
        const parsed = JSON.parse(normalized) as unknown;
        if (Array.isArray(parsed)) {
            return `shape=array items=${parsed.length}`;
        }
        if (parsed && typeof parsed === "object") {
            const keys = Object.keys(parsed as Record<string, unknown>);
            return `shape=object keys=${keys.length > 0 ? keys.join(",") : "-"}`;
        }
        return `shape=${typeof parsed}`;
    } catch {
        return "shape=text";
    }
}

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
    if (!SHOULD_LOG_AI_VERBOSE) {
        // eslint-disable-next-line no-console
        console.log(
            `[AIResponse][${requestId}] len=${normalized.length} ${summarizeAIResponseForLog(normalized)}`,
        );
        return;
    }
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
    if (!SHOULD_LOG_AI_VERBOSE) {
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
        return;
    }
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
    if (!SHOULD_LOG_AI_THINKING) {
        return;
    }
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

function parseUpstreamErrorMessage(rawText: string): string {
    if (rawText.length === 0) {
        return "AI 检测请求失败";
    }
    try {
        const payload = JSON.parse(rawText) as {
            error?: { message?: string };
            message?: string;
        };
        return payload.error?.message ?? payload.message ?? "AI 检测请求失败";
    } catch {
        return rawText.slice(0, 400);
    }
}

function parseUnknownUpstreamError(error: unknown): {
    status: number;
    message: string;
} {
    if (error instanceof Error) {
        const record = asRecord(error as unknown);
        const statusCandidate = Number(record?.status ?? record?.code);
        const status =
            Number.isInteger(statusCandidate) &&
            statusCandidate >= 400 &&
            statusCandidate <= 599
                ? statusCandidate
                : 500;
        return {
            status,
            message: error.message || "AI 检测请求失败",
        };
    }

    const record = asRecord(error);
    if (!record) {
        return {
            status: 500,
            message: "AI 检测请求失败",
        };
    }

    const statusCandidate = Number(record.status ?? record.code);
    const status =
        Number.isInteger(statusCandidate) &&
        statusCandidate >= 400 &&
        statusCandidate <= 599
            ? statusCandidate
            : 500;
    const message =
        (typeof record.message === "string" && record.message) ||
        (typeof record.details === "string" && record.details) ||
        "AI 检测请求失败";
    return {
        status,
        message,
    };
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
    if ("reasoning_content" in objectValue) {
        const text = readTextValue(objectValue.reasoning_content);
        if (text.length > 0) {
            return text;
        }
    }
    if ("reasoningContent" in objectValue) {
        const text = readTextValue(objectValue.reasoningContent);
        if (text.length > 0) {
            return text;
        }
    }
    if ("reasoning_text" in objectValue) {
        const text = readTextValue(objectValue.reasoning_text);
        if (text.length > 0) {
            return text;
        }
    }
    if ("reasoningText" in objectValue) {
        const text = readTextValue(objectValue.reasoningText);
        if (text.length > 0) {
            return text;
        }
    }
    if ("reasoning" in objectValue) {
        const text = readTextValue(objectValue.reasoning);
        if (text.length > 0) {
            return text;
        }
    }
    if ("thinking" in objectValue) {
        const text = readTextValue(objectValue.thinking);
        if (text.length > 0) {
            return text;
        }
    }
    if ("summary" in objectValue) {
        const text = readTextValue(objectValue.summary);
        if (text.length > 0) {
            return text;
        }
    }
    if ("output_text" in objectValue) {
        const text = readTextValue(objectValue.output_text);
        if (text.length > 0) {
            return text;
        }
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
            typeof partRecord.type === "string"
                ? partRecord.type.toLowerCase()
                : "";
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
    const rootThinking =
        readTextValue(root.reasoning_content) ||
        readTextValue(root.reasoningContent) ||
        readTextValue(root.reasoning_text) ||
        readTextValue(root.reasoningText) ||
        readTextValue(root.reasoning) ||
        readTextValue(root.thinking) ||
        readTextValue(root.summary);
    if (rootThinking.length > 0) {
        thinkingChunks.push(rootThinking);
    }

    const choices = Array.isArray(root.choices) ? root.choices : [];
    const firstChoice = asRecord(choices[0]);
    if (firstChoice) {
        const choiceThinking =
            readTextValue(firstChoice.reasoning_content) ||
            readTextValue(firstChoice.reasoningContent) ||
            readTextValue(firstChoice.reasoning_text) ||
            readTextValue(firstChoice.reasoningText) ||
            readTextValue(firstChoice.reasoning) ||
            readTextValue(firstChoice.thinking) ||
            readTextValue(firstChoice.summary);
        if (choiceThinking.length > 0) {
            thinkingChunks.push(choiceThinking);
        }

        const delta = asRecord(firstChoice.delta);
        if (delta) {
            extractContentParts(delta.content, answerChunks, thinkingChunks);
            const deltaText = readTextValue(delta.content);
            // extractContentParts already handles string/array content; avoid duplicate chunks.
            if (
                typeof delta.content !== "string" &&
                !Array.isArray(delta.content) &&
                deltaText.length > 0
            ) {
                answerChunks.push(deltaText);
            }
            const deltaThinking =
                readTextValue(delta.reasoning_content) ||
                readTextValue(delta.reasoningContent) ||
                readTextValue(delta.reasoning_text) ||
                readTextValue(delta.reasoningText) ||
                readTextValue(delta.reasoning) ||
                readTextValue(delta.thinking) ||
                readTextValue(delta.summary);
            if (deltaThinking.length > 0) {
                thinkingChunks.push(deltaThinking);
            }
        }

        const message = asRecord(firstChoice.message);
        if (message) {
            extractContentParts(message.content, answerChunks, thinkingChunks);
            const messageText = readTextValue(message.content);
            // extractContentParts already handles string/array content; avoid duplicate chunks.
            if (
                typeof message.content !== "string" &&
                !Array.isArray(message.content) &&
                messageText.length > 0
            ) {
                answerChunks.push(messageText);
            }
            const messageThinking =
                readTextValue(message.reasoning_content) ||
                readTextValue(message.reasoningContent) ||
                readTextValue(message.reasoning_text) ||
                readTextValue(message.reasoningText) ||
                readTextValue(message.reasoning) ||
                readTextValue(message.thinking) ||
                readTextValue(message.summary);
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
            extractContentParts(
                outputRecord.content,
                answerChunks,
                thinkingChunks,
            );
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

type AIChatMessage = {
    role: "user" | "assistant";
    content: string;
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

type GeminiContentPart =
    | {
          text: string;
      }
    | {
          inlineData: {
              mimeType: string;
              data: string;
          };
      }
    | {
          fileData: {
              mimeType?: string;
              fileUri: string;
          };
      };

type GeminiGenerateContentRequest = {
    systemInstruction?: {
        parts: Array<{ text: string }>;
    };
    contents: Array<{
        role: "user" | "model";
        parts: GeminiContentPart[];
    }>;
    generationConfig?: {
        thinkingConfig?: {
            includeThoughts?: boolean;
            thinkingLevel?: "low" | "medium" | "high";
        };
    };
};

function isGemini3OrLaterModel(model: string): boolean {
    const matched = /gemini-(\d+)/i.exec(model);
    if (!matched?.[1]) {
        return false;
    }
    const major = Number.parseInt(matched[1], 10);
    return Number.isFinite(major) && major >= 3;
}

function mapReasoningEffortToGeminiThinkingLevel(
    model: string,
    effort: AIReasoningEffort,
): "low" | "medium" | "high" {
    if (effort === "low") {
        return "low";
    }
    if (effort === "high") {
        return "high";
    }
    return /flash/i.test(model) ? "medium" : "high";
}

function buildGeminiThinkingConfig(
    model: string,
    effort: AIReasoningEffort,
): {
    includeThoughts: true;
    thinkingLevel?: "low" | "medium" | "high";
} {
    if (!isGemini3OrLaterModel(model)) {
        return {
            includeThoughts: true,
        };
    }
    return {
        includeThoughts: true,
        thinkingLevel: mapReasoningEffortToGeminiThinkingLevel(model, effort),
    };
}

function parseBase64DataUrl(value: string): {
    mimeType: string;
    data: string;
} | null {
    const match = value.trim().match(/^data:([^;,]+);base64,(.+)$/i);
    if (!match || match.length < 3) {
        return null;
    }
    const mimeType = match[1].trim();
    const data = match[2].trim();
    if (!mimeType || !data) {
        return null;
    }
    return {
        mimeType,
        data,
    };
}

function buildGeminiUserParts(
    promptContent: PromptBuildResult,
): GeminiContentPart[] {
    const parts: GeminiContentPart[] = [
        {
            text: promptContent.promptText,
        },
    ];

    for (const field of promptContent.imageFields) {
        const imageLabel =
            field.value.trim().length > 0
                ? `字段图片：${field.title}（说明：${field.value}）`
                : `字段图片：${field.title}`;
        parts.push({ text: imageLabel });

        const imageUrl = field.imageUrl.trim();
        const inlineData = parseBase64DataUrl(imageUrl);
        if (inlineData) {
            parts.push({
                inlineData: {
                    mimeType: inlineData.mimeType,
                    data: inlineData.data,
                },
            });
            continue;
        }

        if (imageUrl.startsWith("gs://")) {
            const ext = getImageExtFromPathLike(imageUrl);
            parts.push({
                fileData: {
                    fileUri: imageUrl,
                    mimeType: ext ? getImageMimeType(ext) : undefined,
                },
            });
            continue;
        }

        parts.push({
            text: `[图片地址（未转为 inlineData）: ${imageUrl}]`,
        });
    }

    return parts;
}

function extractGeminiStreamTextPayload(payload: unknown): {
    answerText: string;
    thinkingText: string;
} {
    const root = asRecord(payload);
    if (!root) {
        return {
            answerText: "",
            thinkingText: "",
        };
    }

    const answerChunks: string[] = [];
    const thinkingChunks: string[] = [];
    const candidates = Array.isArray(root.candidates) ? root.candidates : [];

    for (const candidate of candidates) {
        const candidateRecord = asRecord(candidate);
        if (!candidateRecord) {
            continue;
        }
        const content = asRecord(candidateRecord.content);
        const parts = Array.isArray(content?.parts) ? content.parts : [];
        for (const part of parts) {
            const partRecord = asRecord(part);
            if (!partRecord) {
                continue;
            }
            const text =
                typeof partRecord.text === "string"
                    ? partRecord.text
                    : readTextValue(partRecord.text);
            if (text.length === 0) {
                continue;
            }
            const type =
                typeof partRecord.type === "string"
                    ? partRecord.type.toLowerCase()
                    : "";
            const isThought =
                partRecord.thought === true ||
                type.includes("thought") ||
                type.includes("reasoning") ||
                type.includes("thinking");
            if (isThought) {
                thinkingChunks.push(text);
            } else {
                answerChunks.push(text);
            }
        }
    }

    return {
        answerText: answerChunks.join(""),
        thinkingText: thinkingChunks.join(""),
    };
}

function parseGeminiStreamErrorPayload(
    data: string,
): { status: number; message: string; retryable: boolean } | null {
    const trimmed = data.trim();
    if (trimmed.length === 0) {
        return null;
    }
    let payload: unknown = null;
    try {
        payload = JSON.parse(trimmed) as unknown;
    } catch {
        return null;
    }
    const root = asRecord(payload);
    if (!root) {
        return null;
    }
    const errorRecord = asRecord(root.error) ?? root;
    if (!errorRecord) {
        return null;
    }

    let statusText =
        typeof errorRecord.status === "string" ? errorRecord.status : "";
    let code =
        typeof errorRecord.code === "number" ? errorRecord.code : Number.NaN;
    let message =
        typeof errorRecord.message === "string" ? errorRecord.message : "";

    if (message.trim().startsWith("{")) {
        try {
            const innerPayload = JSON.parse(message) as unknown;
            const innerRoot = asRecord(innerPayload);
            const innerError = innerRoot
                ? (asRecord(innerRoot.error) ?? innerRoot)
                : null;
            if (innerError) {
                if (!Number.isFinite(code)) {
                    const innerCode =
                        typeof innerError.code === "number"
                            ? innerError.code
                            : Number(innerError.code);
                    if (Number.isFinite(innerCode)) {
                        code = innerCode;
                    }
                }
                if (!statusText && typeof innerError.status === "string") {
                    statusText = innerError.status;
                }
                if (
                    typeof innerError.message === "string" &&
                    innerError.message
                ) {
                    message = innerError.message;
                }
            }
        } catch {
            // Ignore nested parse errors.
        }
    }

    if (!Number.isFinite(code) && statusText === "RESOURCE_EXHAUSTED") {
        code = 429;
    }

    const status = Number.isFinite(code) ? Math.trunc(code) : 500;
    const finalMessage = message || "AI 检测请求失败";
    const retryable = status === 429 || statusText === "RESOURCE_EXHAUSTED";
    return { status, message: finalMessage, retryable };
}

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
        if (
            candidate.value !== undefined &&
            typeof candidate.value !== "string"
        ) {
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
                value:
                    typeof candidate.value === "string" ? candidate.value : "",
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

function toAIChatMessages(value: unknown): AIChatMessage[] | null {
    if (!Array.isArray(value)) {
        return null;
    }

    const result: AIChatMessage[] = [];
    for (const item of value) {
        if (!item || typeof item !== "object" || Array.isArray(item)) {
            return null;
        }

        const candidate = item as {
            role?: unknown;
            content?: unknown;
        };
        if (candidate.role !== "user" && candidate.role !== "assistant") {
            return null;
        }
        if (
            typeof candidate.content !== "string" ||
            candidate.content.trim().length === 0
        ) {
            return null;
        }
        result.push({
            role: candidate.role,
            content: candidate.content,
        });
    }

    return result;
}

function buildOpenAIUserContent(
    promptContent: PromptBuildResult,
): string | OpenAIMessageContentPart[] {
    if (promptContent.imageFields.length === 0) {
        return promptContent.promptText;
    }

    return [
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
    ];
}

function buildChatSystemPrompt(prompt: string, fields: AIDetectField[]): string {
    if (
        !prompt.includes("{{fields_json}}") &&
        !prompt.includes("{{fields_text}}") &&
        !prompt.includes("{{image_fields}}")
    ) {
        return prompt;
    }
    return buildPromptContent(prompt, fields).promptText;
}

function buildOpenAIChatMessages(
    prompt: string,
    fields: AIDetectField[],
    messages: AIChatMessage[],
): Array<{
    role: "system" | "user" | "assistant";
    content: string | OpenAIMessageContentPart[];
}> {
    const result: Array<{
        role: "system" | "user" | "assistant";
        content: string | OpenAIMessageContentPart[];
    }> = [
        {
            role: "system",
            content: buildChatSystemPrompt(prompt, fields),
        },
    ];

    if (fields.length > 0) {
        const fieldContext = buildPromptContent(
            "当前题目的固定上下文字段如下。请在后续对话中始终结合这些字段回答。\n\n{{fields_json}}",
            fields,
        );
        result.push({
            role: "user",
            content: buildOpenAIUserContent(fieldContext),
        });
    }

    messages.forEach((message) => {
        result.push({
            role: message.role,
            content: message.content,
        });
    });

    return result;
}

function buildGeminiChatContents(
    fields: AIDetectField[],
    messages: AIChatMessage[],
): GeminiGenerateContentRequest["contents"] {
    const contents: GeminiGenerateContentRequest["contents"] = [];

    if (fields.length > 0) {
        const fieldContext = buildPromptContent(
            "当前题目的固定上下文字段如下。请在后续对话中始终结合这些字段回答。\n\n{{fields_json}}",
            fields,
        );
        contents.push({
            role: "user",
            parts: buildGeminiUserParts(fieldContext),
        });
    }

    messages.forEach((message) => {
        contents.push({
            role: message.role === "assistant" ? "model" : "user",
            parts: [{ text: message.content }],
        });
    });

    return contents;
}

function buildPromptContent(
    prompt: string,
    fields: AIDetectField[],
): PromptBuildResult {
    const fieldSummary: Record<string, string> = {};
    const imageFields: Array<{
        title: string;
        value: string;
        imageUrl: string;
    }> = [];

    fields.forEach((field) => {
        if (field.type === "image" && field.imageUrl) {
            const summary =
                field.value.trim().length > 0
                    ? `[图片] ${field.value}`
                    : "[图片]";
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

type AttemptError = Error & {
    status: number;
    emitted: boolean;
};

function createAttemptError(
    message: string,
    status: number,
    emitted = false,
): AttemptError {
    const error = new Error(message) as AttemptError;
    error.status = status;
    error.emitted = emitted;
    return error;
}

function validateProviderPayload(
    item: unknown,
    index: number,
): { provider?: AIProviderEndpointConfig; message?: string } {
    if (!item || typeof item !== "object") {
        return { message: `provider at index ${index} must be an object` };
    }
    const candidate = item as {
        name?: unknown;
        apiType?: unknown;
        apiUrl?: unknown;
        apiKey?: unknown;
    };
    if (!isNonEmptyString(candidate.name)) {
        return { message: "provider name must be a non-empty string" };
    }
    if (!isAIProviderApiType(candidate.apiType)) {
        return { message: `【${candidate.name}】apiType must be openai, gemini or anthropic` };
    }
    if (!isNonEmptyString(candidate.apiUrl)) {
        return { message: `【${candidate.name}】apiUrl must be a non-empty string` };
    }
    if (!isNonEmptyString(candidate.apiKey)) {
        return { message: `【${candidate.name}】apiKey must be a non-empty string` };
    }
    return {
        provider: {
            name: candidate.name.trim(),
            apiType: candidate.apiType,
            apiUrl: candidate.apiUrl,
            apiKey: candidate.apiKey,
        },
    };
}

function validateRoutePayload(
    item: unknown,
    index: number,
    providersByName: Map<string, AIProviderEndpointConfig>,
): { route?: AIModelRouteConfig; message?: string } {
    if (!item || typeof item !== "object") {
        return { message: `route at index ${index} must be an object` };
    }
    const candidate = item as {
        name?: unknown;
        model?: unknown;
        reasoningEffort?: unknown;
        retryCount?: unknown;
        steps?: unknown;
    };
    if (!isNonEmptyString(candidate.name)) {
        return { message: "route name must be a non-empty string" };
    }
    if (!isNonEmptyString(candidate.model)) {
        return { message: `【${candidate.name}】model must be a non-empty string` };
    }
    if (
        candidate.reasoningEffort !== undefined &&
        !isAIReasoningEffort(candidate.reasoningEffort)
    ) {
        return {
            message: `【${candidate.name}】reasoningEffort must be low, medium or high`,
        };
    }
    if (
        candidate.retryCount !== undefined &&
        !isValidAIRetryCount(candidate.retryCount)
    ) {
        return {
            message: `【${candidate.name}】retryCount must be an integer between ${MIN_AI_RETRY_COUNT} and ${MAX_AI_RETRY_COUNT}`,
        };
    }
    if (!Array.isArray(candidate.steps) || candidate.steps.length === 0) {
        return { message: `【${candidate.name}】steps must be a non-empty array` };
    }
    const steps: AIRouteStep[] = [];
    let routeApiType: AIProviderApiType | null = null;
    for (const [stepIndex, step] of candidate.steps.entries()) {
        if (!step || typeof step !== "object") {
            return {
                message: `【${candidate.name}】step ${stepIndex + 1} must be an object`,
            };
        }
        const providerName = (step as { providerName?: unknown }).providerName;
        if (!isNonEmptyString(providerName)) {
            return {
                message: `【${candidate.name}】step ${stepIndex + 1} providerName must be a non-empty string`,
            };
        }
        const provider = providersByName.get(providerName.trim());
        if (!provider) {
            return {
                message: `【${candidate.name}】step ${stepIndex + 1} providerName must reference an existing provider`,
            };
        }
        if (!routeApiType) {
            routeApiType = provider.apiType;
        } else if (routeApiType !== provider.apiType) {
            return {
                message: `【${candidate.name}】all steps must use providers with the same apiType`,
            };
        }
        steps.push({ providerName: provider.name });
    }
    return {
        route: {
            name: candidate.name.trim(),
            model: candidate.model.trim(),
            reasoningEffort: isAIReasoningEffort(candidate.reasoningEffort)
                ? candidate.reasoningEffort
                : "high",
            retryCount: normalizeAIRetryCount(candidate.retryCount),
            steps,
        },
    };
}

function validateChatPayload(
    item: unknown,
    routesByName: Map<string, AIModelRouteConfig>,
    providersByName: Map<string, AIProviderEndpointConfig>,
): { chat?: FileAIChatConfig; message?: string } {
    if (!item || typeof item !== "object") {
        return { message: "chat must be an object" };
    }
    const candidate = item as {
        routeName?: unknown;
        prompt?: unknown;
        defaultSubmitFieldKeys?: unknown;
    };
    if (!isNonEmptyString(candidate.routeName)) {
        return { message: "chat routeName must be a non-empty string" };
    }
    const route = routesByName.get(candidate.routeName.trim());
    if (!route) {
        return { message: "chat routeName must reference an existing route" };
    }
    const routeApiType = getRouteApiType(route, providersByName);
    if (!routeApiType) {
        return { message: "chat route providers are invalid" };
    }
    if (routeApiType === "anthropic") {
        return { message: "chat cannot use an anthropic route yet" };
    }
    if (!isNonEmptyString(candidate.prompt)) {
        return { message: "chat prompt must be a non-empty string" };
    }
    if (
        !Array.isArray(candidate.defaultSubmitFieldKeys) ||
        !candidate.defaultSubmitFieldKeys.every((entry) => typeof entry === "string")
    ) {
        return { message: "chat defaultSubmitFieldKeys must be a string array" };
    }
    return {
        chat: {
            routeName: candidate.routeName.trim(),
            prompt: candidate.prompt,
            defaultSubmitFieldKeys: candidate.defaultSubmitFieldKeys,
        },
    };
}

function validateCleaningPayload(
    item: unknown,
    routesByName: Map<string, AIModelRouteConfig>,
    providersByName: Map<string, AIProviderEndpointConfig>,
): { cleaning?: FileAICleaningConfigMap; message?: string } {
    if (!item || typeof item !== "object") {
        return { message: "cleaning must be an object" };
    }

    const cleaning = {} as FileAICleaningConfigMap;
    for (const toolKey of AI_CLEANING_TOOL_ORDER) {
        const toolValue = (item as Record<string, unknown>)[toolKey];
        if (!toolValue || typeof toolValue !== "object") {
            return {
                message: `${AI_CLEANING_TOOL_LABELS[toolKey]} config must be an object`,
            };
        }
        const candidate = toolValue as {
            routeName?: unknown;
            submitFieldKeys?: unknown;
            prompt?: unknown;
            autoFillEnabled?: unknown;
            outputMappings?: unknown;
        };
        if (!isNonEmptyString(candidate.routeName)) {
            return {
                message: `${AI_CLEANING_TOOL_LABELS[toolKey]} routeName must be a non-empty string`,
            };
        }
        const route = routesByName.get(candidate.routeName.trim());
        if (!route) {
            return {
                message: `${AI_CLEANING_TOOL_LABELS[toolKey]} routeName must reference an existing route`,
            };
        }
        const routeApiType = getRouteApiType(route, providersByName);
        if (!routeApiType) {
            return {
                message: `${AI_CLEANING_TOOL_LABELS[toolKey]} route providers are invalid`,
            };
        }
        if (routeApiType === "anthropic") {
            return {
                message: `${AI_CLEANING_TOOL_LABELS[toolKey]} cannot use an anthropic route yet`,
            };
        }
        if (
            !Array.isArray(candidate.submitFieldKeys) ||
            !candidate.submitFieldKeys.every((entry) => typeof entry === "string")
        ) {
            return {
                message: `${AI_CLEANING_TOOL_LABELS[toolKey]} submitFieldKeys must be a string array`,
            };
        }
        if (!isNonEmptyString(candidate.prompt)) {
            return {
                message: `${AI_CLEANING_TOOL_LABELS[toolKey]} prompt must be a non-empty string`,
            };
        }
        if (!Array.isArray(candidate.outputMappings)) {
            return {
                message: `${AI_CLEANING_TOOL_LABELS[toolKey]} outputMappings must be an array`,
            };
        }
        const allowedOutputKeys = new Set(AI_CLEANING_TOOL_OUTPUT_KEYS[toolKey]);
        const mappingMap = new Map<string, AICleaningOutputMapping>();
        for (const mapping of candidate.outputMappings) {
            if (!mapping || typeof mapping !== "object") {
                return {
                    message: `${AI_CLEANING_TOOL_LABELS[toolKey]} outputMappings must contain objects`,
                };
            }
            const outputKey = (mapping as { outputKey?: unknown }).outputKey;
            const targetFieldKey = (mapping as { targetFieldKey?: unknown })
                .targetFieldKey;
            if (!isNonEmptyString(outputKey)) {
                return {
                    message: `${AI_CLEANING_TOOL_LABELS[toolKey]} outputKey must be a non-empty string`,
                };
            }
            if (!allowedOutputKeys.has(outputKey.trim())) {
                return {
                    message: `${AI_CLEANING_TOOL_LABELS[toolKey]} contains invalid outputKey: ${outputKey}`,
                };
            }
            if (
                targetFieldKey !== undefined &&
                targetFieldKey !== null &&
                typeof targetFieldKey !== "string"
            ) {
                return {
                    message: `${AI_CLEANING_TOOL_LABELS[toolKey]} targetFieldKey must be a string`,
                };
            }
            if (mappingMap.has(outputKey.trim())) {
                return {
                    message: `${AI_CLEANING_TOOL_LABELS[toolKey]} outputKey duplicated: ${outputKey}`,
                };
            }
            mappingMap.set(outputKey.trim(), {
                outputKey: outputKey.trim(),
                targetFieldKey:
                    typeof targetFieldKey === "string"
                        ? targetFieldKey.trim()
                        : "",
            });
        }
        if (mappingMap.size !== allowedOutputKeys.size) {
            return {
                message: `${AI_CLEANING_TOOL_LABELS[toolKey]} outputMappings are incomplete`,
            };
        }
        cleaning[toolKey] = {
            routeName: candidate.routeName.trim(),
            submitFieldKeys: candidate.submitFieldKeys,
            prompt: candidate.prompt,
            autoFillEnabled: candidate.autoFillEnabled === true,
            outputMappings: AI_CLEANING_TOOL_OUTPUT_KEYS[toolKey].map(
                (outputKey) => mappingMap.get(outputKey)!,
            ),
        } satisfies FileAICleaningToolConfig;
    }

    return { cleaning };
}

function getRouteApiType(
    route: AIModelRouteConfig,
    providersByName: Map<string, AIProviderEndpointConfig>,
): AIProvider | null {
    let apiType: AIProvider | null = null;
    for (const step of route.steps) {
        const provider = providersByName.get(step.providerName);
        if (!provider) {
            return null;
        }
        if (!apiType) {
            apiType = provider.apiType as AIProvider;
            continue;
        }
        if (apiType !== provider.apiType) {
            return null;
        }
    }
    return apiType;
}

function resolveRouteSteps(route: AIModelRouteConfig): {
    route: AIRoute;
    steps: AIResolvedRouteStep[];
    apiType: AIProvider;
} | null {
    const steps: AIResolvedRouteStep[] = [];
    let apiType: AIProvider | null = null;
    for (const step of route.steps) {
        const provider = findAIProviderEndpointByName(step.providerName);
        if (!provider) {
            return null;
        }
        const providerApiType = provider.apiType as AIProvider;
        if (!apiType) {
            apiType = providerApiType;
        } else if (apiType !== providerApiType) {
            return null;
        }
        steps.push({
            provider,
            route: {
                name: route.name,
                model: route.model,
                reasoningEffort: route.reasoningEffort,
                retryCount: route.retryCount,
                steps: route.steps.map((item) => ({ providerName: item.providerName })),
            },
        });
    }
    if (!apiType || steps.length === 0) {
        return null;
    }
    return {
        route: steps[0].route,
        steps,
        apiType,
    };
}

function normalizeFieldsForAI(fieldPayload: AIDetectField[]): AIDetectField[] {
    return fieldPayload.map((field): AIDetectField => {
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
}

async function runOpenAIProviderAttempt({
    provider,
    route,
    prompt,
    fields,
    messages,
    signal,
    requestId,
    onAnswerChunk,
    onThinkingChunk,
}: {
    provider: AIProviderEndpointConfig;
    route: AIRoute;
    prompt: string;
    fields: AIDetectField[];
    messages?: AIChatMessage[];
    signal?: AbortSignal;
    requestId: string;
    onAnswerChunk?: (chunk: string) => void;
    onThinkingChunk?: (chunk: string) => void;
}): Promise<{ answerText: string; thinkingText: string; emittedAny: boolean }> {
    const normalizedModel = normalizeModelName(route.model);
    const normalizedOpenAIUrl = normalizeOpenAIUrl(provider.apiUrl);
    try {
        new URL(normalizedOpenAIUrl);
    } catch {
        throw createAttemptError("url is invalid", 400);
    }

    const requestMessages = messages
        ? buildOpenAIChatMessages(prompt, fields, messages)
        : [
              {
                  role: "user" as const,
                  content: buildOpenAIUserContent(
                      buildPromptContent(prompt, fields),
                  ),
              },
          ];

    let lastFailedStatus = 500;
    let lastFailedMessage = messages ? "AI 聊天请求失败" : "AI 检测请求失败";
    const totalAttempts = route.retryCount + 1;

    for (let attempt = 1; attempt <= totalAttempts; attempt += 1) {
        let upstream: Response | null = null;
        try {
            upstream = await fetch(normalizedOpenAIUrl, {
                method: "POST",
                signal,
                headers: {
                    "Content-Type": "application/json",
                    Authorization: `Bearer ${provider.apiKey}`,
                },
                body: JSON.stringify({
                    model: normalizedModel,
                    stream: true,
                    messages: requestMessages,
                    reasoning: {
                        effort: route.reasoningEffort,
                    },
                }),
            });
        } catch (error) {
            if (signal?.aborted) {
                throw error;
            }
            const parsedError = parseUnknownUpstreamError(error);
            lastFailedStatus = parsedError.status;
            lastFailedMessage = parsedError.message;
            continue;
        }

        if (!upstream.ok || !upstream.body) {
            if (!upstream.ok) {
                const rawText = await upstream.text().catch(() => "");
                lastFailedStatus = upstream.status || 500;
                lastFailedMessage = parseUpstreamErrorMessage(rawText);
            } else {
                lastFailedStatus = 502;
                lastFailedMessage = "AI 响应流为空";
            }
            continue;
        }

        const decoder = new TextDecoder();
        let buffer = "";
        let rawStreamPreview = "";
        let answerText = "";
        let thinkingText = "";
        let emittedAny = false;
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
                    logAIResponseById(requestId, answerText);
                    if (thinkingText.trim().length > 0) {
                        logAIThinkingById(requestId, thinkingText);
                    }
                    return { answerText, thinkingText, emittedAny };
                }
                if (data.length === 0) {
                    continue;
                }
                try {
                    const payload = JSON.parse(data) as unknown;
                    const extracted = extractStreamTextPayload(payload);
                    if (extracted.thinkingText.length > 0) {
                        thinkingText += extracted.thinkingText;
                        emittedAny = true;
                        onThinkingChunk?.(extracted.thinkingText);
                    }
                    if (extracted.answerText.length > 0) {
                        answerText += extracted.answerText;
                        emittedAny = true;
                        onAnswerChunk?.(extracted.answerText);
                    }
                } catch {
                    // Ignore non-JSON chunks.
                }
            }
        }

        buffer += decoder.decode();
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
                        thinkingText += extracted.thinkingText;
                        emittedAny = true;
                        onThinkingChunk?.(extracted.thinkingText);
                    }
                    if (extracted.answerText.length > 0) {
                        answerText += extracted.answerText;
                        emittedAny = true;
                        onAnswerChunk?.(extracted.answerText);
                    }
                } catch {
                    // Ignore trailing invalid chunk.
                }
            }
        }

        logAIResponseById(requestId, answerText);
        if (thinkingText.trim().length > 0) {
            logAIThinkingById(requestId, thinkingText);
        }
        if (answerText.trim().length === 0 && rawStreamPreview.trim().length > 0) {
            logAIRawResponseById(requestId, rawStreamPreview);
        }
        return { answerText, thinkingText, emittedAny };
    }

    throw createAttemptError(lastFailedMessage, lastFailedStatus);
}

async function runGeminiProviderAttempt({
    provider,
    route,
    prompt,
    fields,
    messages,
    signal,
    requestId,
    onAnswerChunk,
    onThinkingChunk,
}: {
    provider: AIProviderEndpointConfig;
    route: AIRoute;
    prompt: string;
    fields: AIDetectField[];
    messages?: AIChatMessage[];
    signal?: AbortSignal;
    requestId: string;
    onAnswerChunk?: (chunk: string) => void;
    onThinkingChunk?: (chunk: string) => void;
}): Promise<{ answerText: string; thinkingText: string; emittedAny: boolean }> {
    const normalizedModel = normalizeModelName(route.model);
    const normalizedGeminiUrl = normalizeGeminiEndpoint(
        provider.apiUrl,
        normalizedModel,
    );
    try {
        new URL(normalizedGeminiUrl);
    } catch {
        throw createAttemptError("url is invalid", 400);
    }

    const geminiThinkingConfig = buildGeminiThinkingConfig(
        normalizedModel,
        route.reasoningEffort,
    );
    const geminiAuth = buildGeminiAuthHeaders(
        normalizedGeminiUrl,
        provider.apiKey,
    );
    const requestBody: GeminiGenerateContentRequest = {
        systemInstruction: messages
            ? {
                  parts: [
                      {
                          text: buildChatSystemPrompt(prompt, fields),
                      },
                  ],
              }
            : undefined,
        contents: messages
            ? buildGeminiChatContents(fields, messages)
            : [
                  {
                      role: "user",
                      parts: buildGeminiUserParts(
                          buildPromptContent(prompt, fields),
                      ),
                  },
              ],
        generationConfig: {
            thinkingConfig: geminiThinkingConfig,
        },
    };

    let lastFailedStatus = 500;
    let lastFailedMessage = messages ? "AI 聊天请求失败" : "AI 检测请求失败";
    const totalAttempts = route.retryCount + 1;

    for (let attempt = 1; attempt <= totalAttempts; attempt += 1) {
        let upstream: Response | null = null;
        try {
            upstream = await fetch(normalizedGeminiUrl, {
                method: "POST",
                signal,
                headers: geminiAuth.headers,
                body: JSON.stringify(requestBody),
            });
        } catch (error) {
            if (signal?.aborted) {
                throw error;
            }
            const parsedError = parseUnknownUpstreamError(error);
            lastFailedStatus = parsedError.status;
            lastFailedMessage = parsedError.message;
            continue;
        }

        if (!upstream.ok || !upstream.body) {
            if (!upstream.ok) {
                const rawText = await upstream.text().catch(() => "");
                lastFailedStatus = upstream.status || 500;
                lastFailedMessage = parseUpstreamErrorMessage(rawText);
            } else {
                lastFailedStatus = 502;
                lastFailedMessage = "AI 响应流为空";
            }
            continue;
        }

        const reader = upstream.body.getReader();
        const decoder = new TextDecoder();
        let buffer = "";
        let rawStreamPreview = "";
        let currentEventType = "";
        let answerText = "";
        let thinkingText = "";
        let emittedAny = false;
        let shouldRetry = false;
        let shouldFail = false;

        const handleGeminiData = (
            data: string,
            eventType: string,
        ): "continue" | "retry" | "fail" | "done" => {
            if (data === "[DONE]") {
                return "done";
            }
            if (eventType === "error") {
                const parsedError = parseGeminiStreamErrorPayload(data);
                if (parsedError) {
                    lastFailedStatus = parsedError.status;
                    lastFailedMessage = parsedError.message;
                    if (!emittedAny && parsedError.retryable) {
                        return "retry";
                    }
                    if (!emittedAny) {
                        return "fail";
                    }
                    return "done";
                }
            }
            if (data.length === 0) {
                return "continue";
            }
            try {
                const payload = JSON.parse(data) as unknown;
                const extracted = extractGeminiStreamTextPayload(payload);
                if (extracted.thinkingText.length > 0) {
                    thinkingText += extracted.thinkingText;
                    emittedAny = true;
                    onThinkingChunk?.(extracted.thinkingText);
                }
                if (extracted.answerText.length > 0) {
                    answerText += extracted.answerText;
                    emittedAny = true;
                    onAnswerChunk?.(extracted.answerText);
                }
            } catch {
                // Ignore invalid chunks.
            }
            return "continue";
        };

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
                if (trimmed.length === 0) {
                    currentEventType = "";
                    continue;
                }
                if (trimmed.startsWith("event:")) {
                    currentEventType = trimmed.slice(6).trim().toLowerCase();
                    continue;
                }
                if (!trimmed.startsWith("data:")) {
                    continue;
                }
                const data = trimmed.slice(5).trim();
                const action = handleGeminiData(data, currentEventType);
                currentEventType = "";
                if (action === "retry") {
                    shouldRetry = true;
                    break;
                }
                if (action === "fail") {
                    shouldFail = true;
                    break;
                }
                if (action === "done") {
                    logAIResponseById(requestId, answerText);
                    if (thinkingText.trim().length > 0) {
                        logAIThinkingById(requestId, thinkingText);
                    }
                    return { answerText, thinkingText, emittedAny };
                }
            }

            if (shouldRetry || shouldFail) {
                break;
            }
        }

        if (shouldRetry || shouldFail) {
            await reader.cancel().catch(() => {});
            if (shouldRetry) {
                continue;
            }
            break;
        }

        buffer += decoder.decode();
        if (buffer.length > 0 && buffer.includes("data:")) {
            const maybeData = buffer
                .split(/\r?\n/)
                .map((line) => line.trim())
                .find((line) => line.startsWith("data:"));
            const value = maybeData ? maybeData.slice(5).trim() : "";
            if (value && value !== "[DONE]") {
                const action = handleGeminiData(value, currentEventType);
                if (action === "retry") {
                    await reader.cancel().catch(() => {});
                    continue;
                }
                if (action === "fail") {
                    await reader.cancel().catch(() => {});
                    break;
                }
            }
        }

        logAIResponseById(requestId, answerText);
        if (thinkingText.trim().length > 0) {
            logAIThinkingById(requestId, thinkingText);
        }
        if (answerText.trim().length === 0 && rawStreamPreview.trim().length > 0) {
            logAIRawResponseById(requestId, rawStreamPreview);
        }
        return { answerText, thinkingText, emittedAny };
    }

    throw createAttemptError(lastFailedMessage, lastFailedStatus);
}

async function streamAIRouteResponse({
    req,
    res,
    resolvedRoute,
    prompt,
    fields,
    messages,
}: {
    req: express.Request;
    res: express.Response;
    resolvedRoute: {
        route: AIRoute;
        steps: AIResolvedRouteStep[];
        apiType: AIProvider;
    };
    prompt: string;
    fields: AIDetectField[];
    messages?: AIChatMessage[];
}) {
    const aiFields = normalizeFieldsForAI(fields);
    const aiRequestId = randomUUID().slice(0, 8);
    const startedAt = Date.now();
    const startedAtIso = formatLogTimestamp(startedAt);
    const failures: string[] = [];
    let headersCommitted = false;
    const controller = new AbortController();

    const ensureHeaders = () => {
        if (headersCommitted) {
            return;
        }
        res.status(200);
        res.setHeader("Content-Type", "application/x-ndjson; charset=utf-8");
        res.setHeader("Cache-Control", "no-cache, no-transform");
        res.setHeader("Connection", "keep-alive");
        res.setHeader("X-Accel-Buffering", "no");
        res.flushHeaders();
        headersCommitted = true;
    };

    const abortUpstream = () => {
        if (!controller.signal.aborted) {
            controller.abort();
        }
    };

    req.on("aborted", abortUpstream);
    req.on("close", () => {
        if (req.aborted) {
            abortUpstream();
        }
    });
    res.on("close", () => {
        if (!res.writableEnded) {
            abortUpstream();
        }
    });

    try {
        for (const [index, step] of resolvedRoute.steps.entries()) {
            if (controller.signal.aborted) {
                return;
            }
            try {
                const result = await runRouteStepAttempt({
                    step,
                    prompt,
                    fields: aiFields,
                    messages,
                    signal: controller.signal,
                    requestId: `${aiRequestId}-step-${index + 1}`,
                    onAnswerChunk: (chunk) => {
                        ensureHeaders();
                        writeAIClientStreamEvent(res, {
                            type: "answer",
                            text: chunk,
                        });
                    },
                    onThinkingChunk: (chunk) => {
                        ensureHeaders();
                        writeAIClientStreamEvent(res, {
                            type: "thinking",
                            text: chunk,
                        });
                    },
                });

                if (
                    result.answerText.trim().length === 0 &&
                    result.thinkingText.trim().length === 0
                ) {
                    failures.push(`${step.provider.name}: AI 返回为空`);
                    continue;
                }

                ensureHeaders();
                writeAIClientStreamEvent(res, { type: "done" });
                res.end();
                return;
            } catch (error) {
                if (controller.signal.aborted) {
                    return;
                }
                const parsed = parseUnknownUpstreamError(error);
                failures.push(`${step.provider.name}: ${parsed.message}`);
            }
        }

        console.log(
            `[AIResponseError][${aiRequestId}] startedAt=${startedAtIso} elapsedMs=${Date.now() - startedAt} route=${resolvedRoute.route.name} message=${failures.join(" | ")}`,
        );
        return res.status(502).json({
            message:
                failures.length > 0
                    ? failures.join(" | ")
                    : "所有模型提供商均调用失败",
        });
    } catch (error) {
        if (controller.signal.aborted) {
            return;
        }
        const parsed = parseUnknownUpstreamError(error);
        if (res.headersSent) {
            if (!res.writableEnded) {
                res.end();
            }
            return;
        }
        return res.status(parsed.status).json({ message: parsed.message });
    }
}

async function runRouteStepAttempt({
    step,
    prompt,
    fields,
    messages,
    signal,
    requestId,
    onAnswerChunk,
    onThinkingChunk,
}: {
    step: AIResolvedRouteStep;
    prompt: string;
    fields: AIDetectField[];
    messages?: AIChatMessage[];
    signal?: AbortSignal;
    requestId: string;
    onAnswerChunk?: (chunk: string) => void;
    onThinkingChunk?: (chunk: string) => void;
}) {
    if (isAnthropicCompatibleProvider(step.provider.apiType as AIProvider)) {
        throw createAttemptError("Anthropic 暂不支持测试和正式调用", 400);
    }
    if (isOpenAICompatibleProvider(step.provider.apiType as AIProvider)) {
        return runOpenAIProviderAttempt({
            provider: step.provider,
            route: step.route,
            prompt,
            fields,
            messages,
            signal,
            requestId,
            onAnswerChunk,
            onThinkingChunk,
        });
    }
    if (isGeminiCompatibleProvider(step.provider.apiType as AIProvider)) {
        return runGeminiProviderAttempt({
            provider: step.provider,
            route: step.route,
            prompt,
            fields,
            messages,
            signal,
            requestId,
            onAnswerChunk,
            onThinkingChunk,
        });
    }
    throw createAttemptError("provider is invalid", 400);
}

export const registerAIRoutes = (app: express.Express) => {
    app.get("/api/ai-config/:fileName", (req, res) => {
        const fileName = decodeURIComponent(req.params.fileName);
        return res.json({
            providers: listAIProviderEndpoints(),
            routes: listAIModelRoutes(),
            stages: getFileAIStageConfigs(fileName),
            chat: getFileAIChatConfig(fileName),
            cleaning: getFileAICleaningConfig(fileName),
        });
    });

    app.put("/api/ai-config/providers", (req, res) => {
        const { providers } = req.body as { providers?: unknown };
        if (!Array.isArray(providers)) {
            return res.status(400).json({ message: "providers must be an array" });
        }

        const nextProviders: AIProviderEndpointConfig[] = [];
        const nameSet = new Set<string>();
        for (const [index, item] of providers.entries()) {
            const validation = validateProviderPayload(item, index);
            if (!validation.provider) {
                return res.status(400).json({ message: validation.message });
            }
            if (nameSet.has(validation.provider.name)) {
                return res.status(400).json({
                    message: `provider name duplicated: ${validation.provider.name}`,
                });
            }
            nameSet.add(validation.provider.name);
            nextProviders.push(validation.provider);
        }
        if (nextProviders.length === 0) {
            return res.status(400).json({ message: "providers must not be empty" });
        }

        saveAIProviderEndpoints(nextProviders);
        return res.json({ ok: true });
    });

    app.put("/api/ai-config/routes", (req, res) => {
        const { routes } = req.body as { routes?: unknown };
        if (!Array.isArray(routes)) {
            return res.status(400).json({ message: "routes must be an array" });
        }

        const providersByName = new Map(
            listAIProviderEndpoints().map((provider) => [provider.name, provider]),
        );
        if (providersByName.size === 0) {
            return res.status(400).json({ message: "providers is required before routes" });
        }

        const nextRoutes: AIModelRouteConfig[] = [];
        const nameSet = new Set<string>();
        for (const [index, item] of routes.entries()) {
            const validation = validateRoutePayload(item, index, providersByName);
            if (!validation.route) {
                return res.status(400).json({ message: validation.message });
            }
            if (nameSet.has(validation.route.name)) {
                return res.status(400).json({
                    message: `route name duplicated: ${validation.route.name}`,
                });
            }
            nameSet.add(validation.route.name);
            nextRoutes.push(validation.route);
        }
        if (nextRoutes.length === 0) {
            return res.status(400).json({ message: "routes must not be empty" });
        }

        saveAIModelRoutes(nextRoutes);
        return res.json({ ok: true });
    });

    app.put("/api/ai-config/:fileName/stages", (req, res) => {
        const fileName = decodeURIComponent(req.params.fileName);
        const { stages } = req.body as { stages?: unknown };
        if (!stages || typeof stages !== "object") {
            return res.status(400).json({ message: "stages must be an object" });
        }

        const routes = listAIModelRoutes();
        const providersByName = new Map(
            listAIProviderEndpoints().map((provider) => [provider.name, provider]),
        );
        const routesByName = new Map(routes.map((route) => [route.name, route]));

        const stageMap = {} as Record<AIDetectStageKey, FileAIStageConfig>;
        for (const stageKey of AI_STAGE_ORDER) {
            const stageValue = (stages as Record<string, unknown>)[stageKey];
            if (!stageValue || typeof stageValue !== "object") {
                return res.status(400).json({
                    message: `${AI_STAGE_LABELS[stageKey]} config must be an object`,
                });
            }
            const stage = stageValue as {
                routeName?: unknown;
                submitFieldKeys?: unknown;
                prompt?: unknown;
            };
            if (!isNonEmptyString(stage.routeName)) {
                return res.status(400).json({
                    message: `${AI_STAGE_LABELS[stageKey]} routeName must be a non-empty string`,
                });
            }
            const route = routesByName.get(stage.routeName.trim());
            if (!route) {
                return res.status(400).json({
                    message: `${AI_STAGE_LABELS[stageKey]} routeName must reference an existing route`,
                });
            }
            const routeApiType = getRouteApiType(route, providersByName);
            if (!routeApiType) {
                return res.status(400).json({
                    message: `${AI_STAGE_LABELS[stageKey]} route providers are invalid`,
                });
            }
            if (routeApiType === "anthropic") {
                return res.status(400).json({
                    message: `${AI_STAGE_LABELS[stageKey]} cannot use an anthropic route yet`,
                });
            }
            if (
                !Array.isArray(stage.submitFieldKeys) ||
                !stage.submitFieldKeys.every((item) => typeof item === "string")
            ) {
                return res.status(400).json({
                    message: `${AI_STAGE_LABELS[stageKey]} submitFieldKeys must be a string array`,
                });
            }
            if (!isNonEmptyString(stage.prompt)) {
                return res.status(400).json({
                    message: `${AI_STAGE_LABELS[stageKey]} prompt must be a non-empty string`,
                });
            }
            stageMap[stageKey] = {
                routeName: stage.routeName.trim(),
                submitFieldKeys: stage.submitFieldKeys,
                prompt: stage.prompt,
            };
        }

        saveFileAIStageConfigs(fileName, stageMap);
        return res.json({ ok: true });
    });

    app.put("/api/ai-config/:fileName/chat", (req, res) => {
        const fileName = decodeURIComponent(req.params.fileName);
        const { chat } = req.body as { chat?: unknown };
        const routes = listAIModelRoutes();
        const providersByName = new Map(
            listAIProviderEndpoints().map((provider) => [provider.name, provider]),
        );
        const routesByName = new Map(routes.map((route) => [route.name, route]));
        const validation = validateChatPayload(
            chat,
            routesByName,
            providersByName,
        );
        if (!validation.chat) {
            return res.status(400).json({ message: validation.message });
        }

        saveFileAIChatConfig(fileName, validation.chat);
        return res.json({ ok: true });
    });

    app.put("/api/ai-config/:fileName/cleaning", (req, res) => {
        const fileName = decodeURIComponent(req.params.fileName);
        const { cleaning } = req.body as { cleaning?: unknown };
        const routes = listAIModelRoutes();
        const providersByName = new Map(
            listAIProviderEndpoints().map((provider) => [provider.name, provider]),
        );
        const routesByName = new Map(routes.map((route) => [route.name, route]));
        const validation = validateCleaningPayload(
            cleaning,
            routesByName,
            providersByName,
        );
        if (!validation.cleaning) {
            return res.status(400).json({ message: validation.message });
        }

        saveFileAICleaningConfig(fileName, validation.cleaning);
        return res.json({ ok: true });
    });

    app.post("/api/ai-config/routes/test", async (req, res) => {
        const { provider, route, stepIndex } = req.body as {
            provider?: unknown;
            route?: unknown;
            stepIndex?: unknown;
        };
        const providerValidation = validateProviderPayload(provider, 0);
        if (!providerValidation.provider) {
            return res.status(400).json({ message: providerValidation.message });
        }
        const providerMap = new Map([[providerValidation.provider.name, providerValidation.provider]]);
        const routeValidation = validateRoutePayload(route, 0, providerMap);
        if (!routeValidation.route) {
            return res.status(400).json({ message: routeValidation.message });
        }
        if (
            typeof stepIndex !== "number" ||
            !Number.isInteger(stepIndex) ||
            stepIndex < 0 ||
            stepIndex >= routeValidation.route.steps.length
        ) {
            return res.status(400).json({ message: "stepIndex is invalid" });
        }
        if (providerValidation.provider.apiType === "anthropic") {
            return res.status(400).json({ message: "Anthropic 暂不支持测试" });
        }

        const startedAt = Date.now();
        try {
            const result = await runRouteStepAttempt({
                step: {
                    provider: providerValidation.provider,
                    route: {
                        name: routeValidation.route.name,
                        model: routeValidation.route.model,
                        reasoningEffort: routeValidation.route.reasoningEffort,
                        retryCount: routeValidation.route.retryCount,
                        steps: routeValidation.route.steps,
                    },
                },
                prompt: "请直接回复：你好",
                fields: [{ title: "测试消息", type: "text", value: "你好" }],
                requestId: `route-test-${randomUUID().slice(0, 8)}`,
            });
            const preview = [result.thinkingText.trim(), result.answerText.trim()]
                .filter((item) => item.length > 0)
                .join("\n\n")
                .slice(0, 300);
            return res.json({
                ok: true,
                durationMs: Date.now() - startedAt,
                providerName: providerValidation.provider.name,
                apiType: providerValidation.provider.apiType,
                routeName: routeValidation.route.name,
                model: routeValidation.route.model,
                preview,
            });
        } catch (error) {
            const parsed = parseUnknownUpstreamError(error);
            return res.status(parsed.status).json({ message: parsed.message });
        }
    });

    app.post("/api/ai-detect/stream", async (req, res) => {
        const { routeName, prompt, fields } = req.body as {
            routeName?: unknown;
            prompt?: unknown;
            fields?: unknown;
        };
        if (!isNonEmptyString(routeName)) {
            return res.status(400).json({ message: "routeName must be a non-empty string" });
        }
        if (!isNonEmptyString(prompt)) {
            return res.status(400).json({ message: "prompt must be a non-empty string" });
        }
        const fieldPayload = toAIDetectFields(fields);
        if (!fieldPayload || fieldPayload.length === 0) {
            return res.status(400).json({ message: "fields must be a non-empty array" });
        }

        const route = findAIModelRouteByName(routeName.trim());
        if (!route) {
            return res.status(404).json({ message: "模型路由不存在" });
        }
        const resolvedRoute = resolveRouteSteps(route);
        if (!resolvedRoute) {
            return res.status(400).json({ message: "模型路由引用的提供商无效" });
        }
        if (resolvedRoute.apiType === "anthropic") {
            return res.status(400).json({ message: "Anthropic 暂不支持正式调用" });
        }

        return streamAIRouteResponse({
            req,
            res,
            resolvedRoute,
            prompt,
            fields: fieldPayload,
        });
    });

    app.post("/api/ai-chat/stream", async (req, res) => {
        const { routeName, prompt, fields, messages } = req.body as {
            routeName?: unknown;
            prompt?: unknown;
            fields?: unknown;
            messages?: unknown;
        };
        if (!isNonEmptyString(routeName)) {
            return res.status(400).json({ message: "routeName must be a non-empty string" });
        }
        if (!isNonEmptyString(prompt)) {
            return res.status(400).json({ message: "prompt must be a non-empty string" });
        }
        const fieldPayload = toAIDetectFields(fields);
        if (!fieldPayload) {
            return res.status(400).json({ message: "fields must be an array" });
        }
        const messagePayload = toAIChatMessages(messages);
        if (!messagePayload || messagePayload.length === 0) {
            return res.status(400).json({ message: "messages must be a non-empty array" });
        }

        const route = findAIModelRouteByName(routeName.trim());
        if (!route) {
            return res.status(404).json({ message: "模型路由不存在" });
        }
        const resolvedRoute = resolveRouteSteps(route);
        if (!resolvedRoute) {
            return res.status(400).json({ message: "模型路由引用的提供商无效" });
        }
        if (resolvedRoute.apiType === "anthropic") {
            return res.status(400).json({ message: "Anthropic 暂不支持正式调用" });
        }

        return streamAIRouteResponse({
            req,
            res,
            resolvedRoute,
            prompt,
            fields: fieldPayload,
            messages: messagePayload,
        });
    });
};
