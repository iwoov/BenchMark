import type {
    AIDetectConfig,
    AIDetectProfile,
    AIDetectStageConfigMap,
    NamedAIDetectProfile,
    AIDetectStageConfig,
    NamedAIDetectConfig,
    ParsedColumn,
    ParsedRow,
} from "../types";
import {
    AI_STAGE_ORDER,
    DEFAULT_AI_PROFILE_NAME,
    DEFAULT_AI_PROFILES,
    DEFAULT_AI_BATCH_CONCURRENCY,
    DEFAULT_AI_CONFIG_NAME,
    DEFAULT_AI_STAGE_CONFIGS,
    DEFAULT_AI_RETRY_COUNT,
    DEFAULT_IDEALAB_GEMINI_URL,
    DEFAULT_IDEALAB_OPENAI_URL,
    DEFAULT_MODELROUTER_GEMINI_URL,
    DEFAULT_MODELROUTER_OPENAI_URL,
    MAX_AI_BATCH_CONCURRENCY,
    MAX_AI_RETRY_COUNT,
    MIN_AI_BATCH_CONCURRENCY,
    MIN_AI_RETRY_COUNT,
} from "./constants";
import type {
    AIBatchTaskState,
    AIDetectFieldPayload,
    AIDetectStreamResult,
} from "./types";
import { getCellImageSources } from "./file-helpers";

export function getDefaultAIUrl(provider: AIDetectProfile["provider"]): string {
    if (provider === "gemini") {
        return DEFAULT_IDEALAB_GEMINI_URL;
    }
    if (provider === "modelrouter-gemini") {
        return DEFAULT_MODELROUTER_GEMINI_URL;
    }
    if (provider === "modelrouter-openai") {
        return DEFAULT_MODELROUTER_OPENAI_URL;
    }
    return DEFAULT_IDEALAB_OPENAI_URL;
}

export function isGeminiProvider(
    provider: AIDetectProfile["provider"],
): boolean {
    return provider === "gemini" || provider === "modelrouter-gemini";
}

function inferAIProviderFromUrl(
    provider: AIDetectProfile["provider"],
    url: unknown,
): AIDetectProfile["provider"] {
    if (typeof url !== "string") {
        return provider;
    }
    const trimmed = url.trim().toLowerCase();
    if (!trimmed) {
        return provider;
    }
    if (!trimmed.includes("routify.alibaba-inc.com")) {
        return provider;
    }
    if (trimmed.includes("/protocol/vertex/")) {
        return "modelrouter-gemini";
    }
    if (trimmed.includes("/protocol/openai/")) {
        return "modelrouter-openai";
    }
    return provider;
}

const MODEL_PROVIDER_PREFIXES = new Set([
    "openai",
    "google",
    "anthropic",
    "gemini",
    "vertex",
    "idealab",
]);

function splitModelId(
    value: string,
    fallbackProvider: string,
): { provider: string; name: string } {
    const trimmed = value.trim();
    if (!trimmed) {
        return { provider: fallbackProvider, name: "" };
    }
    const slashIndex = trimmed.indexOf("/");
    if (slashIndex > 0) {
        return {
            provider: trimmed.slice(0, slashIndex),
            name: trimmed.slice(slashIndex + 1),
        };
    }
    return { provider: fallbackProvider, name: trimmed };
}

function stripModelProviderPrefix(value: string): string {
    const trimmed = value.trim();
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

function composeModelId(_provider: string, name: string): string {
    const trimmedName = name.trim();
    return trimmedName;
}

export function getAIBatchTaskStatusText(task: AIBatchTaskState): string {
    if (task.status === "running") {
        return "运行中";
    }
    if (task.status === "completed") {
        return task.failed > 0 ? "已完成（含失败）" : "已完成";
    }
    return "未启动";
}

export function formatDuration(ms: number): string {
    const totalSeconds = Math.max(0, Math.floor(ms / 1000));
    const minutes = Math.floor(totalSeconds / 60)
        .toString()
        .padStart(2, "0");
    const seconds = (totalSeconds % 60).toString().padStart(2, "0");
    return `${minutes}:${seconds}`;
}

export function normalizeAIBatchConcurrency(value: unknown): number {
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

export function normalizeAIRetryCount(value: unknown): number {
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

export function composeAISaveText(
    answerText: string,
    thinkingText: string,
): string {
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

export function cloneAIDetectProfile(
    profile: AIDetectProfile,
): AIDetectProfile {
    return { ...profile };
}

export function cloneNamedAIDetectProfile(
    item: NamedAIDetectProfile,
): NamedAIDetectProfile {
    return {
        name: item.name,
        profile: cloneAIDetectProfile(item.profile),
    };
}

export function cloneAIDetectStageConfig(
    stageConfig: AIDetectStageConfig,
): AIDetectStageConfig {
    return {
        ...stageConfig,
        submitFieldKeys: [...stageConfig.submitFieldKeys],
    };
}

export function cloneAIDetectConfig(config: AIDetectConfig): AIDetectConfig {
    const profiles = config.profiles.map(cloneNamedAIDetectProfile);
    const stages = {} as AIDetectStageConfigMap;
    AI_STAGE_ORDER.forEach((stageKey) => {
        stages[stageKey] = cloneAIDetectStageConfig(config.stages[stageKey]);
    });
    return { profiles, stages };
}

export function createDefaultAIDetectConfig(): AIDetectConfig {
    const profiles = DEFAULT_AI_PROFILES.map(cloneNamedAIDetectProfile);
    const stages = {} as AIDetectStageConfigMap;
    AI_STAGE_ORDER.forEach((stageKey) => {
        stages[stageKey] = cloneAIDetectStageConfig(
            DEFAULT_AI_STAGE_CONFIGS[stageKey],
        );
    });
    return { profiles, stages };
}

function normalizeLoadedAIDetectProfile(
    value: unknown,
    fallback: AIDetectProfile,
): AIDetectProfile {
    if (!value || typeof value !== "object") {
        return cloneAIDetectProfile(fallback);
    }

    const candidate = value as Partial<AIDetectProfile>;
    const rawProvider = (value as { provider?: unknown }).provider;
    const provider = inferAIProviderFromUrl(
        rawProvider === "openai" ||
            rawProvider === "gemini" ||
            rawProvider === "modelrouter-openai" ||
            rawProvider === "modelrouter-gemini"
            ? rawProvider
            : rawProvider === "vertex"
              ? "gemini"
              : rawProvider === "idealab"
                ? "openai"
                : fallback.provider,
        candidate.url,
    );
    const rawModel =
        typeof candidate.model === "string" && candidate.model.trim().length > 0
            ? stripModelProviderPrefix(candidate.model)
            : "";
    const rawModelProvider =
        typeof candidate.modelProvider === "string"
            ? candidate.modelProvider.trim()
            : "";
    const rawModelName =
        typeof candidate.modelName === "string"
            ? stripModelProviderPrefix(candidate.modelName)
            : "";
    const resolvedModel =
        rawModel.length > 0
            ? rawModel
            : rawModelName.length > 0
              ? composeModelId(rawModelProvider, rawModelName)
              : fallback.model;
    const normalizedModel = stripModelProviderPrefix(resolvedModel);
    const derivedModel = splitModelId(
        normalizedModel,
        isGeminiProvider(provider) ? "google" : "openai",
    );

    return {
        provider,
        url:
            typeof candidate.url === "string" && candidate.url.trim().length > 0
                ? candidate.url
                : getDefaultAIUrl(provider),
        model:
            normalizedModel.trim().length > 0
                ? normalizedModel
                : stripModelProviderPrefix(fallback.model),
        modelProvider: rawModelProvider || derivedModel.provider,
        modelName: rawModelName || derivedModel.name,
        apiKey: typeof candidate.apiKey === "string" ? candidate.apiKey : "",
        reasoningEffort:
            candidate.reasoningEffort === "low" ||
            candidate.reasoningEffort === "medium" ||
            candidate.reasoningEffort === "high"
                ? candidate.reasoningEffort
                : fallback.reasoningEffort,
        retryCount: normalizeAIRetryCount(candidate.retryCount),
    };
}

function normalizeLoadedProfiles(value: unknown): NamedAIDetectProfile[] {
    if (!Array.isArray(value)) {
        return [];
    }

    const result: NamedAIDetectProfile[] = [];
    const nameCount = new Map<string, number>();

    value.forEach((item, index) => {
        if (!item || typeof item !== "object") {
            return;
        }
        const candidate = item as {
            name?: unknown;
            profile?: unknown;
        } & Partial<AIDetectProfile>;
        const rawName =
            typeof candidate.name === "string" &&
            candidate.name.trim().length > 0
                ? candidate.name.trim()
                : index === 0
                  ? DEFAULT_AI_PROFILE_NAME
                  : `接口配置 ${index + 1}`;
        const count = nameCount.get(rawName) ?? 0;
        nameCount.set(rawName, count + 1);
        const name = count > 0 ? `${rawName}-${count + 1}` : rawName;
        const profileSource =
            candidate.profile && typeof candidate.profile === "object"
                ? candidate.profile
                : candidate;
        const profile = normalizeLoadedAIDetectProfile(
            profileSource,
            DEFAULT_AI_PROFILES[0].profile,
        );
        result.push({ name, profile });
    });

    return result;
}

function normalizeLoadedAIDetectStageConfig(
    value: unknown,
    fallback: AIDetectStageConfig,
    fallbackProfileName: string,
): AIDetectStageConfig {
    if (!value || typeof value !== "object") {
        return {
            ...cloneAIDetectStageConfig(fallback),
            profileName: fallbackProfileName,
        };
    }

    const candidate = value as Partial<AIDetectStageConfig>;
    const submitFieldKeys = Array.isArray(candidate.submitFieldKeys)
        ? candidate.submitFieldKeys.filter(
              (item): item is string => typeof item === "string",
          )
        : [];
    const profileName =
        typeof candidate.profileName === "string" &&
        candidate.profileName.trim()
            ? candidate.profileName.trim()
            : fallbackProfileName;

    return {
        profileName,
        submitFieldKeys,
        prompt:
            typeof candidate.prompt === "string" &&
            candidate.prompt.trim().length > 0
                ? candidate.prompt
                : fallback.prompt,
    };
}

export function normalizeLoadedAIDetectConfig(value: unknown): AIDetectConfig {
    if (!value || typeof value !== "object") {
        return createDefaultAIDetectConfig();
    }

    const candidate = value as { stages?: unknown; profiles?: unknown };
    const normalizedProfiles = normalizeLoadedProfiles(candidate.profiles);

    if (candidate.stages && typeof candidate.stages === "object") {
        const rawStages = candidate.stages as Record<string, unknown>;
        const hasProfileName = AI_STAGE_ORDER.some((stageKey) => {
            const stageValue = rawStages[stageKey] as {
                profileName?: unknown;
            } | null;
            return (
                stageValue &&
                typeof stageValue === "object" &&
                typeof stageValue.profileName === "string"
            );
        });

        const profiles =
            normalizedProfiles.length > 0
                ? normalizedProfiles
                : DEFAULT_AI_PROFILES.map(cloneNamedAIDetectProfile);
        const fallbackProfileName =
            profiles[0]?.name ?? DEFAULT_AI_PROFILE_NAME;
        const stages = {} as AIDetectStageConfigMap;

        if (hasProfileName) {
            AI_STAGE_ORDER.forEach((stageKey) => {
                stages[stageKey] = normalizeLoadedAIDetectStageConfig(
                    rawStages[stageKey],
                    DEFAULT_AI_STAGE_CONFIGS[stageKey],
                    fallbackProfileName,
                );
            });
            return { profiles, stages };
        }

        // Legacy stage payload: provider/url/model/... per stage
        const legacyProfiles: NamedAIDetectProfile[] = [];
        const legacyStages = {} as AIDetectStageConfigMap;
        AI_STAGE_ORDER.forEach((stageKey, index) => {
            const stageValue = rawStages[stageKey];
            const profileName =
                AI_STAGE_ORDER.length > 1
                    ? `${DEFAULT_AI_PROFILE_NAME}-${index + 1}`
                    : DEFAULT_AI_PROFILE_NAME;
            legacyProfiles.push({
                name: profileName,
                profile: normalizeLoadedAIDetectProfile(
                    stageValue,
                    DEFAULT_AI_PROFILES[0].profile,
                ),
            });
            legacyStages[stageKey] = normalizeLoadedAIDetectStageConfig(
                stageValue,
                DEFAULT_AI_STAGE_CONFIGS[stageKey],
                profileName,
            );
        });
        return {
            profiles: legacyProfiles,
            stages: legacyStages,
        };
    }

    // Legacy single-config payload
    const legacyProfile: NamedAIDetectProfile = {
        name: DEFAULT_AI_PROFILE_NAME,
        profile: normalizeLoadedAIDetectProfile(
            value,
            DEFAULT_AI_PROFILES[0].profile,
        ),
    };
    const stages = {} as AIDetectStageConfigMap;
    AI_STAGE_ORDER.forEach((stageKey) => {
        stages[stageKey] = normalizeLoadedAIDetectStageConfig(
            value,
            DEFAULT_AI_STAGE_CONFIGS[stageKey],
            legacyProfile.name,
        );
    });
    return { profiles: [legacyProfile], stages };
}

export function normalizeAIConfigName(value: unknown): string {
    if (typeof value !== "string") {
        return DEFAULT_AI_CONFIG_NAME;
    }
    const trimmed = value.trim();
    return trimmed.length > 0 ? trimmed : DEFAULT_AI_CONFIG_NAME;
}

export function normalizeLoadedNamedAIDetectConfigs(
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
            profiles?: unknown;
            stages?: unknown;
        };
        const name = normalizeAIConfigName(candidate.name);
        if (usedNames.has(name)) {
            return;
        }

        const configSource =
            candidate.config && typeof candidate.config === "object"
                ? candidate.config
                : candidate.stages && typeof candidate.stages === "object"
                  ? {
                        stages: candidate.stages,
                        profiles: candidate.profiles,
                    }
                  : item;
        const config = normalizeLoadedAIDetectConfig(configSource);
        usedNames.add(name);
        result.push({ name, config });
    });

    return result;
}

function normalizeAIDetectStageConfigForColumns(
    config: AIDetectStageConfig,
    columns: ParsedColumn[],
): AIDetectStageConfig {
    const keySet = new Set(columns.map((column) => column.key));
    const submitFieldKeys = config.submitFieldKeys.filter((key) =>
        keySet.has(key),
    );
    return {
        ...config,
        submitFieldKeys,
    };
}

export function normalizeAIDetectConfigForColumns(
    config: AIDetectConfig,
    columns: ParsedColumn[],
): AIDetectConfig {
    const profiles =
        config.profiles && config.profiles.length > 0
            ? config.profiles.map(cloneNamedAIDetectProfile)
            : DEFAULT_AI_PROFILES.map(cloneNamedAIDetectProfile);
    const fallbackProfileName = profiles[0]?.name ?? DEFAULT_AI_PROFILE_NAME;
    const stageProfileNames = new Set(profiles.map((item) => item.name));

    const stages = {} as AIDetectStageConfigMap;
    AI_STAGE_ORDER.forEach((stageKey) => {
        const stageConfig =
            config.stages?.[stageKey] ?? DEFAULT_AI_STAGE_CONFIGS[stageKey];
        const normalizedStage = normalizeAIDetectStageConfigForColumns(
            stageConfig,
            columns,
        );
        const normalizedProfileName =
            normalizedStage.profileName &&
            stageProfileNames.has(normalizedStage.profileName)
                ? normalizedStage.profileName
                : fallbackProfileName;
        stages[stageKey] = {
            ...normalizedStage,
            profileName: normalizedProfileName,
        };
    });

    return { profiles, stages };
}

export function normalizeNamedAIDetectConfigsForColumns(
    configs: NamedAIDetectConfig[],
    columns: ParsedColumn[],
): NamedAIDetectConfig[] {
    return configs.map((item) => ({
        name: item.name,
        config: normalizeAIDetectConfigForColumns(item.config, columns),
    }));
}

export function pickAIConfigName(
    configs: NamedAIDetectConfig[],
    preferredName: unknown,
): string {
    if (configs.length === 0) {
        return DEFAULT_AI_CONFIG_NAME;
    }
    if (typeof preferredName === "string") {
        const trimmed = preferredName.trim();
        if (
            trimmed.length > 0 &&
            configs.some((item) => item.name === trimmed)
        ) {
            return trimmed;
        }
    }
    return configs[0].name;
}

export function buildAIDetectFieldsForRow(
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

export async function requestAIDetectResult(
    payload: {
        provider: AIDetectProfile["provider"];
        url: string;
        model: string;
        apiKey: string;
        prompt: string;
        fields: AIDetectFieldPayload[];
        reasoningEffort: AIDetectProfile["reasoningEffort"];
        retryCount: number;
    },
    options?: {
        signal?: AbortSignal;
        onAnswerChunk?: (chunk: string) => void;
        onThinkingChunk?: (chunk: string) => void;
        onChunk?: (chunk: string) => void;
    },
): Promise<AIDetectStreamResult> {
    const normalizedModel = stripModelProviderPrefix(payload.model);
    const response = await fetch("/api/ai-detect/stream", {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
        },
        signal: options?.signal,
        body: JSON.stringify({
            ...payload,
            model: normalizedModel,
        }),
    });

    if (!response.ok) {
        const errorPayload = (await response.json().catch(() => ({}))) as {
            message?: string;
        };
        throw new Error(errorPayload.message ?? "AI 回答失败");
    }
    if (!response.body) {
        throw new Error("AI 响应流为空");
    }

    const reader = response.body.getReader();
    const decoder = new TextDecoder("utf-8");
    const contentType =
        response.headers.get("content-type")?.toLowerCase() ?? "";
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
                    if (
                        event.type === "answer" &&
                        typeof event.text === "string"
                    ) {
                        answerText += event.text;
                        options?.onAnswerChunk?.(event.text);
                        options?.onChunk?.(event.text);
                        continue;
                    }
                    if (
                        event.type === "thinking" &&
                        typeof event.text === "string"
                    ) {
                        thinkingText += event.text;
                        options?.onThinkingChunk?.(event.text);
                        continue;
                    }
                    if (event.type === "done") {
                        continue;
                    }
                } catch {
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
