import type {
    AIDetectConfig,
    AIChatConfig,
    AICleaningConfigMap,
    AICleaningOutputMapping,
    AICleaningToolConfig,
    AICleaningToolKey,
    AIEvaluationTaskConfig,
    AIModelRoute,
    AIProviderApiType,
    AIStreamPhase,
    AIProviderEndpoint,
    AIDetectStageConfigMap,
    AIDetectStageConfig,
    NamedAIDetectConfig,
    ParsedColumn,
    ParsedRow,
} from "../types";
import {
    AI_CLEANING_TOOL_LABELS,
    AI_CLEANING_TOOL_ORDER,
    AI_STAGE_ORDER,
    DEFAULT_AI_CHAT_CONFIG,
    DEFAULT_AI_CLEANING_CONFIGS,
    DEFAULT_AI_EVALUATION_ATTEMPT_COUNT,
    DEFAULT_AI_EVALUATION_TASK,
    DEFAULT_AI_EVALUATION_TASK_ID,
    DEFAULT_AI_EVALUATION_TASK_NAME,
    DEFAULT_AI_EVALUATION_MAX_CONCURRENCY,
    DEFAULT_AI_PROVIDER,
    DEFAULT_AI_PROVIDER_NAME,
    DEFAULT_AI_ROUTE,
    DEFAULT_AI_ROUTE_NAME,
    DEFAULT_AI_BATCH_CONCURRENCY,
    DEFAULT_AI_CONFIG_NAME,
    DEFAULT_AI_STAGE_CONFIGS,
    DEFAULT_AI_RETRY_COUNT,
    DEFAULT_ANTHROPIC_URL,
    DEFAULT_IDEALAB_GEMINI_URL,
    DEFAULT_IDEALAB_OPENAI_URL,
    MAX_AI_EVALUATION_ATTEMPT_COUNT,
    MAX_AI_EVALUATION_MAX_CONCURRENCY,
    MAX_AI_BATCH_CONCURRENCY,
    MAX_AI_RETRY_COUNT,
    MIN_AI_EVALUATION_ATTEMPT_COUNT,
    MIN_AI_EVALUATION_MAX_CONCURRENCY,
    MIN_AI_BATCH_CONCURRENCY,
    MIN_AI_RETRY_COUNT,
} from "./constants";
import type {
    AIBatchTaskState,
    AIChatMessagePayload,
    AIDetectFieldPayload,
    AIDetectStreamResult,
} from "./types";
import { getCellImageSources } from "./file-helpers";

export function getDefaultProviderUrl(apiType: AIProviderApiType): string {
    if (apiType === "gemini") {
        return DEFAULT_IDEALAB_GEMINI_URL;
    }
    if (apiType === "anthropic") {
        return DEFAULT_ANTHROPIC_URL;
    }
    return DEFAULT_IDEALAB_OPENAI_URL;
}

export function isGeminiApiType(apiType: AIProviderApiType): boolean {
    return apiType === "gemini";
}

export function isAnthropicApiType(apiType: AIProviderApiType): boolean {
    return apiType === "anthropic";
}

const MODEL_PROVIDER_PREFIXES = new Set([
    "openai",
    "google",
    "anthropic",
    "gemini",
    "vertex",
    "idealab",
]);

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

function normalizeAIEvaluationAttemptCount(value: unknown): number {
    if (typeof value !== "number" || !Number.isInteger(value)) {
        return DEFAULT_AI_EVALUATION_ATTEMPT_COUNT;
    }
    if (value < MIN_AI_EVALUATION_ATTEMPT_COUNT) {
        return MIN_AI_EVALUATION_ATTEMPT_COUNT;
    }
    if (value > MAX_AI_EVALUATION_ATTEMPT_COUNT) {
        return MAX_AI_EVALUATION_ATTEMPT_COUNT;
    }
    return value;
}

function normalizeAIEvaluationMaxConcurrency(value: unknown): number {
    if (typeof value !== "number" || !Number.isInteger(value)) {
        return DEFAULT_AI_EVALUATION_MAX_CONCURRENCY;
    }
    if (value < MIN_AI_EVALUATION_MAX_CONCURRENCY) {
        return MIN_AI_EVALUATION_MAX_CONCURRENCY;
    }
    if (value > MAX_AI_EVALUATION_MAX_CONCURRENCY) {
        return MAX_AI_EVALUATION_MAX_CONCURRENCY;
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

export interface ParsedAIResult extends Record<string, unknown> {
    is_valid?: boolean | string;
    reason?: string;
    invalid_reason?: string;
    requires_image?: boolean | string;
    has_absolute_words?: boolean | string;
    absolute_words_details?: string;
    is_all_selected_warning?: boolean | string;
    is_consistent?: boolean | string;
    missing_info?: string;
    is_objective?: boolean | string;
    subjectivity_risk_level?: string;
    analysis?: string | Record<string, string>;
    final_answer?: string;
    ai_reasoning_step_by_step?: string;
    ai_final_answer?: string;
    can_be_solved?: boolean | string;
    unsolvable_reason?: string;
    status?: "Pass" | "Fail" | string;
    discrepancy_detail?: string;
    is_answer_consistent?: boolean | string;
    superior_answer?: "expert" | "ai" | "tie" | string | null;
    inconsistency_analysis?: string;
    has_extra_info?: boolean | string;
    extra_info_details?: string;
    is_logic_forced?: boolean | string;
    logic_flaw_details?: string;
    final_verdict?: string;
}

function extractJsonCandidates(content: string): string[] {
    const source = content.includes("【AI结果】")
        ? (content.split("【AI结果】").pop() ?? "")
        : content;

    const jsonCandidates: string[] = [];
    let depth = 0;
    let start = -1;
    let inString = false;
    let escaped = false;

    for (let i = 0; i < source.length; i += 1) {
        const char = source[i];

        if (escaped) {
            escaped = false;
            continue;
        }

        if (char === "\\") {
            escaped = inString;
            continue;
        }

        if (char === '"') {
            inString = !inString;
            continue;
        }

        if (inString) {
            continue;
        }

        if (char === "{") {
            if (depth === 0) {
                start = i;
            }
            depth += 1;
            continue;
        }

        if (char === "}" && depth > 0) {
            depth -= 1;
            if (depth === 0 && start >= 0) {
                jsonCandidates.push(source.slice(start, i + 1));
                start = -1;
            }
        }
    }

    return jsonCandidates;
}

export function parseAIResultJSON(content: string): ParsedAIResult | null {
    if (!content || content.trim().length === 0) {
        return null;
    }

    const jsonCandidates = extractJsonCandidates(content);
    for (let i = jsonCandidates.length - 1; i >= 0; i -= 1) {
        try {
            return JSON.parse(jsonCandidates[i]) as ParsedAIResult;
        } catch {
            // Ignore invalid JSON candidates and continue trying older ones.
        }
    }

    return null;
}

function readStringValue(value: unknown): string {
    return typeof value === "string" ? value.trim() : "";
}

function stringifyAnalysis(analysis: ParsedAIResult["analysis"]): string {
    if (typeof analysis === "string") {
        return analysis.trim();
    }
    if (!analysis || typeof analysis !== "object") {
        return "";
    }

    return Object.entries(analysis)
        .map(([key, value]) =>
            typeof value === "string" && value.trim().length > 0
                ? `${key}: ${value.trim()}`
                : "",
        )
        .filter((item) => item.length > 0)
        .join("\n");
}

export function readBooleanLike(value: unknown): boolean | null {
    if (typeof value === "boolean") {
        return value;
    }
    if (typeof value === "string") {
        const normalized = value.trim().toLowerCase();
        if (
            normalized === "true" ||
            normalized === "yes" ||
            normalized === "1" ||
            normalized === "是"
        ) {
            return true;
        }
        if (
            normalized === "false" ||
            normalized === "no" ||
            normalized === "0" ||
            normalized === "否"
        ) {
            return false;
        }
    }
    if (typeof value === "number") {
        if (value === 1) {
            return true;
        }
        if (value === 0) {
            return false;
        }
    }
    return null;
}

export function extractAIResultFinalAnswer(content: string): string | null {
    if (!content) {
        return null;
    }

    const parsed = parseAIResultJSON(content);
    const aiFinalAnswer = readStringValue(parsed?.ai_final_answer);
    if (aiFinalAnswer.length > 0) {
        return aiFinalAnswer;
    }
    const finalAnswer = readStringValue(parsed?.final_answer);
    if (finalAnswer.length > 0) {
        return finalAnswer;
    }

    const answerSection = content.includes("【AI结果】")
        ? (content.split("【AI结果】").pop() ?? "")
        : content;
    const directMatch =
        /(?:ai_)?final_answer\s*[:：]\s*["“”']?([^"\n\r,}]+)["“”']?/i.exec(
            answerSection,
        );
    if (directMatch?.[1]) {
        return directMatch[1].trim();
    }

    return null;
}

export function buildFinalVerdictExtraFields(
    independentSolvingResult: string,
): AIDetectFieldPayload[] {
    const trimmed = independentSolvingResult.trim();
    if (trimmed.length === 0) {
        return [];
    }

    const parsed = parseAIResultJSON(trimmed);
    const finalAnswer = extractAIResultFinalAnswer(trimmed) ?? "";
    const reasoning =
        readStringValue(parsed?.ai_reasoning_step_by_step) ||
        stringifyAnalysis(parsed?.analysis);
    const canBeSolved = readBooleanLike(parsed?.can_be_solved);
    const unsolvableReason = readStringValue(parsed?.unsolvable_reason);

    const fields: AIDetectFieldPayload[] = [
        {
            title: "AI独立解题结果（第三阶段）",
            type: "text",
            value: independentSolvingResult,
        },
    ];

    if (finalAnswer.length > 0) {
        fields.push({
            title: "AI最终答案（第三阶段）",
            type: "text",
            value: finalAnswer,
        });
    }
    if (reasoning.length > 0) {
        fields.push({
            title: "AI推理过程（第三阶段）",
            type: "text",
            value: reasoning,
        });
    }
    if (canBeSolved !== null) {
        fields.push({
            title: "AI是否可解（第三阶段）",
            type: "text",
            value: canBeSolved ? "可解" : "不可解",
        });
    }
    if (unsolvableReason.length > 0) {
        fields.push({
            title: "AI无法作答原因（第三阶段）",
            type: "text",
            value: unsolvableReason,
        });
    }

    return fields;
}

export function cloneAIProviderEndpoint(
    provider: AIProviderEndpoint,
): AIProviderEndpoint {
    return { ...provider };
}

export function cloneAIModelRoute(route: AIModelRoute): AIModelRoute {
    return {
        ...route,
        steps: route.steps.map((step) => ({ ...step })),
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

export function cloneAIChatConfig(chatConfig: AIChatConfig): AIChatConfig {
    return {
        ...chatConfig,
        defaultSubmitFieldKeys: [...chatConfig.defaultSubmitFieldKeys],
    };
}

export function cloneAICleaningOutputMapping(
    mapping: AICleaningOutputMapping,
): AICleaningOutputMapping {
    return { ...mapping };
}

export function cloneAICleaningToolConfig(
    toolConfig: AICleaningToolConfig,
): AICleaningToolConfig {
    return {
        ...toolConfig,
        submitFieldKeys: [...toolConfig.submitFieldKeys],
        outputMappings: toolConfig.outputMappings.map(
            cloneAICleaningOutputMapping,
        ),
    };
}

export function cloneAIEvaluationTaskConfig(
    config: AIEvaluationTaskConfig,
): AIEvaluationTaskConfig {
    return {
        id: config.id,
        name: config.name,
        enabled: config.enabled,
        attemptCount: config.attemptCount,
        maxConcurrency: config.maxConcurrency,
        answerGeneration: {
            ...config.answerGeneration,
            questionFieldKeys: [...config.answerGeneration.questionFieldKeys],
        },
        answerJudgment: {
            ...config.answerJudgment,
            answerFieldKeys: [...config.answerJudgment.answerFieldKeys],
        },
    };
}

export function cloneAIDetectConfig(config: AIDetectConfig): AIDetectConfig {
    const providers = config.providers.map(cloneAIProviderEndpoint);
    const routes = config.routes.map(cloneAIModelRoute);
    const stages = {} as AIDetectStageConfigMap;
    const cleaning = {} as AICleaningConfigMap;
    AI_STAGE_ORDER.forEach((stageKey) => {
        stages[stageKey] = cloneAIDetectStageConfig(config.stages[stageKey]);
    });
    AI_CLEANING_TOOL_ORDER.forEach((toolKey) => {
        cleaning[toolKey] = cloneAICleaningToolConfig(config.cleaning[toolKey]);
    });
    return {
        providers,
        routes,
        stages,
        evaluationTasks: config.evaluationTasks.map(
            cloneAIEvaluationTaskConfig,
        ),
        chat: cloneAIChatConfig(config.chat),
        cleaning,
    };
}

export function createDefaultAIDetectConfig(): AIDetectConfig {
    const stages = {} as AIDetectStageConfigMap;
    const cleaning = {} as AICleaningConfigMap;
    AI_STAGE_ORDER.forEach((stageKey) => {
        stages[stageKey] = cloneAIDetectStageConfig(
            DEFAULT_AI_STAGE_CONFIGS[stageKey],
        );
    });
    AI_CLEANING_TOOL_ORDER.forEach((toolKey) => {
        cleaning[toolKey] = cloneAICleaningToolConfig(
            DEFAULT_AI_CLEANING_CONFIGS[toolKey],
        );
    });
    return {
        providers: [cloneAIProviderEndpoint(DEFAULT_AI_PROVIDER)],
        routes: [cloneAIModelRoute(DEFAULT_AI_ROUTE)],
        stages,
        evaluationTasks: [
            cloneAIEvaluationTaskConfig(DEFAULT_AI_EVALUATION_TASK),
        ],
        chat: cloneAIChatConfig(DEFAULT_AI_CHAT_CONFIG),
        cleaning,
    };
}

function normalizeLoadedProvider(
    value: unknown,
    fallback: AIProviderEndpoint,
): AIProviderEndpoint {
    if (!value || typeof value !== "object") {
        return cloneAIProviderEndpoint(fallback);
    }

    const candidate = value as {
        name?: unknown;
        apiType?: unknown;
        apiUrl?: unknown;
        apiKey?: unknown;
        provider?: unknown;
        url?: unknown;
    };
    let apiType = fallback.apiType;
    if (
        candidate.apiType === "openai" ||
        candidate.apiType === "gemini" ||
        candidate.apiType === "anthropic"
    ) {
        apiType = candidate.apiType;
    } else if (
        candidate.provider === "gemini" ||
        candidate.provider === "modelrouter-gemini" ||
        candidate.provider === "vertex"
    ) {
        apiType = "gemini";
    } else if (candidate.provider === "anthropic") {
        apiType = "anthropic";
    } else if (
        candidate.provider === "openai" ||
        candidate.provider === "modelrouter-openai" ||
        candidate.provider === "idealab"
    ) {
        apiType = "openai";
    }

    return {
        name:
            typeof candidate.name === "string" && candidate.name.trim().length
                ? candidate.name.trim()
                : fallback.name,
        apiType,
        apiUrl:
            typeof candidate.apiUrl === "string" &&
            candidate.apiUrl.trim().length > 0
                ? candidate.apiUrl
                : typeof candidate.url === "string" &&
                    candidate.url.trim().length > 0
                  ? candidate.url
                  : getDefaultProviderUrl(apiType),
        apiKey: typeof candidate.apiKey === "string" ? candidate.apiKey : "",
    };
}

function normalizeLoadedProviders(value: unknown): AIProviderEndpoint[] {
    if (!Array.isArray(value)) {
        return [];
    }

    const result: AIProviderEndpoint[] = [];
    const nameCount = new Map<string, number>();

    value.forEach((item, index) => {
        if (!item || typeof item !== "object") {
            return;
        }
        const candidate = item as { name?: unknown };
        const rawName =
            typeof candidate.name === "string" &&
            candidate.name.trim().length > 0
                ? candidate.name.trim()
                : index === 0
                  ? DEFAULT_AI_PROVIDER_NAME
                  : `提供商 ${index + 1}`;
        const count = nameCount.get(rawName) ?? 0;
        nameCount.set(rawName, count + 1);
        const name = count > 0 ? `${rawName}-${count + 1}` : rawName;
        result.push(
            normalizeLoadedProvider(
                {
                    ...item,
                    name,
                },
                DEFAULT_AI_PROVIDER,
            ),
        );
    });

    return result;
}

function normalizeLoadedRoute(
    value: unknown,
    fallback: AIModelRoute,
    fallbackProviderName: string,
): AIModelRoute {
    if (!value || typeof value !== "object") {
        return cloneAIModelRoute(fallback);
    }

    const candidate = value as {
        name?: unknown;
        model?: unknown;
        retryCount?: unknown;
        reasoningEffort?: unknown;
        steps?: unknown;
        profileName?: unknown;
        modelName?: unknown;
    };
    const steps = Array.isArray(candidate.steps)
        ? candidate.steps
              .map((step) => {
                  if (!step || typeof step !== "object") {
                      return null;
                  }
                  const providerName = (step as { providerName?: unknown })
                      .providerName;
                  if (
                      typeof providerName !== "string" ||
                      providerName.trim().length === 0
                  ) {
                      return null;
                  }
                  return { providerName: providerName.trim() };
              })
              .filter((item): item is { providerName: string } => item !== null)
        : typeof candidate.profileName === "string" &&
            candidate.profileName.trim().length > 0
          ? [{ providerName: candidate.profileName.trim() }]
          : fallback.steps.length > 0
            ? fallback.steps.map((step) => ({ ...step }))
            : [{ providerName: fallbackProviderName }];

    return {
        name:
            typeof candidate.name === "string" && candidate.name.trim().length
                ? candidate.name.trim()
                : fallback.name,
        model:
            typeof candidate.model === "string" &&
            candidate.model.trim().length > 0
                ? stripModelProviderPrefix(candidate.model)
                : typeof candidate.modelName === "string" &&
                    candidate.modelName.trim().length > 0
                  ? stripModelProviderPrefix(candidate.modelName)
                  : fallback.model,
        reasoningEffort:
            candidate.reasoningEffort === "low" ||
            candidate.reasoningEffort === "medium" ||
            candidate.reasoningEffort === "high"
                ? candidate.reasoningEffort
                : fallback.reasoningEffort,
        retryCount: normalizeAIRetryCount(candidate.retryCount),
        steps:
            steps.length > 0 ? steps : [{ providerName: fallbackProviderName }],
    };
}

function normalizeLoadedRoutes(
    value: unknown,
    fallbackProviderName: string,
): AIModelRoute[] {
    if (!Array.isArray(value)) {
        return [];
    }

    const result: AIModelRoute[] = [];
    const nameCount = new Map<string, number>();

    value.forEach((item, index) => {
        if (!item || typeof item !== "object") {
            return;
        }
        const candidate = item as { name?: unknown };
        const rawName =
            typeof candidate.name === "string" &&
            candidate.name.trim().length > 0
                ? candidate.name.trim()
                : index === 0
                  ? DEFAULT_AI_ROUTE_NAME
                  : `模型路由 ${index + 1}`;
        const count = nameCount.get(rawName) ?? 0;
        nameCount.set(rawName, count + 1);
        const name = count > 0 ? `${rawName}-${count + 1}` : rawName;
        result.push(
            normalizeLoadedRoute(
                {
                    ...item,
                    name,
                },
                DEFAULT_AI_ROUTE,
                fallbackProviderName,
            ),
        );
    });

    return result;
}

function normalizeLoadedAIDetectStageConfig(
    value: unknown,
    fallback: AIDetectStageConfig,
    fallbackRouteName: string,
): AIDetectStageConfig {
    if (!value || typeof value !== "object") {
        return {
            ...cloneAIDetectStageConfig(fallback),
            routeName: fallbackRouteName,
        };
    }

    const candidate = value as Partial<AIDetectStageConfig> & {
        profileName?: unknown;
    };
    const submitFieldKeys = Array.isArray(candidate.submitFieldKeys)
        ? candidate.submitFieldKeys.filter(
              (item): item is string => typeof item === "string",
          )
        : [];
    const routeName =
        typeof candidate.routeName === "string" && candidate.routeName.trim()
            ? candidate.routeName.trim()
            : typeof candidate.profileName === "string" &&
                candidate.profileName.trim()
              ? candidate.profileName.trim()
              : fallbackRouteName;

    return {
        routeName,
        submitFieldKeys,
        prompt:
            typeof candidate.prompt === "string" &&
            candidate.prompt.trim().length > 0
                ? candidate.prompt
                : fallback.prompt,
    };
}

function normalizeLoadedAIChatConfig(
    value: unknown,
    fallbackRouteName: string,
): AIChatConfig {
    if (!value || typeof value !== "object") {
        return {
            ...cloneAIChatConfig(DEFAULT_AI_CHAT_CONFIG),
            routeName: fallbackRouteName,
        };
    }

    const candidate = value as Partial<AIChatConfig>;
    const defaultSubmitFieldKeys = Array.isArray(
        candidate.defaultSubmitFieldKeys,
    )
        ? candidate.defaultSubmitFieldKeys.filter(
              (item): item is string => typeof item === "string",
          )
        : [];
    const routeName =
        typeof candidate.routeName === "string" && candidate.routeName.trim()
            ? candidate.routeName.trim()
            : fallbackRouteName;

    return {
        routeName,
        defaultSubmitFieldKeys,
        prompt:
            typeof candidate.prompt === "string" &&
            candidate.prompt.trim().length > 0
                ? candidate.prompt
                : DEFAULT_AI_CHAT_CONFIG.prompt,
    };
}

function normalizeLoadedAIEvaluationTaskConfig(
    value: unknown,
    fallbackRouteName: string,
    fallbackId: string,
    fallbackName: string,
): AIEvaluationTaskConfig {
    if (!value || typeof value !== "object") {
        return {
            ...cloneAIEvaluationTaskConfig(DEFAULT_AI_EVALUATION_TASK),
            id: fallbackId,
            name: fallbackName,
            answerGeneration: {
                ...DEFAULT_AI_EVALUATION_TASK.answerGeneration,
                routeName: fallbackRouteName,
            },
            answerJudgment: {
                ...DEFAULT_AI_EVALUATION_TASK.answerJudgment,
                routeName: fallbackRouteName,
            },
        };
    }

    const candidate = value as {
        enabled?: unknown;
        attemptCount?: unknown;
        answerGeneration?: unknown;
        answerJudgment?: unknown;
    };
    const legacyStageCandidate = AI_STAGE_ORDER.map(
        (stageKey) =>
            (value as Record<string, unknown>)[stageKey] as
                | {
                      enabled?: unknown;
                      routeName?: unknown;
                      questionFieldKeys?: unknown;
                      answerFieldKeys?: unknown;
                  }
                | undefined,
    ).find((item) => item && typeof item === "object");
    const normalizeRouteName = (routeName: unknown) =>
        typeof routeName === "string" && routeName.trim().length > 0
            ? routeName.trim()
            : fallbackRouteName;
    const answerGeneration =
        candidate.answerGeneration &&
        typeof candidate.answerGeneration === "object"
            ? (candidate.answerGeneration as {
                  routeName?: unknown;
                  prompt?: unknown;
                  questionFieldKeys?: unknown;
              })
            : null;
    const answerJudgment =
        candidate.answerJudgment && typeof candidate.answerJudgment === "object"
            ? (candidate.answerJudgment as {
                  routeName?: unknown;
                  prompt?: unknown;
                  answerFieldKeys?: unknown;
              })
            : null;

    return {
        id:
            typeof (value as { id?: unknown }).id === "string" &&
            (value as { id?: string }).id!.trim().length > 0
                ? (value as { id: string }).id.trim()
                : fallbackId,
        name:
            typeof (value as { name?: unknown }).name === "string" &&
            (value as { name: string }).name.trim().length > 0
                ? (value as { name: string }).name.trim()
                : fallbackName,
        enabled:
            candidate.enabled === true ||
            (candidate.answerGeneration === undefined &&
                candidate.answerJudgment === undefined &&
                legacyStageCandidate?.enabled === true),
        attemptCount: normalizeAIEvaluationAttemptCount(candidate.attemptCount),
        maxConcurrency: normalizeAIEvaluationMaxConcurrency(
            (value as { maxConcurrency?: unknown }).maxConcurrency,
        ),
        answerGeneration: {
            routeName: normalizeRouteName(
                answerGeneration?.routeName ?? legacyStageCandidate?.routeName,
            ),
            prompt:
                typeof answerGeneration?.prompt === "string" &&
                answerGeneration.prompt.trim().length > 0
                    ? answerGeneration.prompt
                    : DEFAULT_AI_EVALUATION_TASK.answerGeneration.prompt,
            questionFieldKeys: Array.isArray(
                answerGeneration?.questionFieldKeys,
            )
                ? answerGeneration.questionFieldKeys.filter(
                      (item): item is string => typeof item === "string",
                  )
                : Array.isArray(legacyStageCandidate?.questionFieldKeys)
                  ? legacyStageCandidate.questionFieldKeys.filter(
                        (item): item is string => typeof item === "string",
                    )
                  : [],
        },
        answerJudgment: {
            routeName: normalizeRouteName(
                answerJudgment?.routeName ?? legacyStageCandidate?.routeName,
            ),
            prompt:
                typeof answerJudgment?.prompt === "string" &&
                answerJudgment.prompt.trim().length > 0
                    ? answerJudgment.prompt
                    : DEFAULT_AI_EVALUATION_TASK.answerJudgment.prompt,
            answerFieldKeys: Array.isArray(answerJudgment?.answerFieldKeys)
                ? answerJudgment.answerFieldKeys.filter(
                      (item): item is string => typeof item === "string",
                  )
                : Array.isArray(legacyStageCandidate?.answerFieldKeys)
                  ? legacyStageCandidate.answerFieldKeys.filter(
                        (item): item is string => typeof item === "string",
                    )
                  : [],
        },
    };
}

function normalizeLoadedAIEvaluationTaskList(
    value: unknown,
    fallbackRouteName: string,
): AIEvaluationTaskConfig[] {
    if (Array.isArray(value)) {
        const result = value
            .map((item, index) =>
                normalizeLoadedAIEvaluationTaskConfig(
                    item,
                    fallbackRouteName,
                    `${DEFAULT_AI_EVALUATION_TASK_ID}-${index + 1}`,
                    index === 0
                        ? DEFAULT_AI_EVALUATION_TASK_NAME
                        : `评测配置 ${index + 1}`,
                ),
            )
            .filter(
                (item, index, items) =>
                    items.findIndex((candidate) => candidate.id === item.id) ===
                    index,
            );
        return result.length > 0
            ? result
            : [
                  normalizeLoadedAIEvaluationTaskConfig(
                      null,
                      fallbackRouteName,
                      DEFAULT_AI_EVALUATION_TASK_ID,
                      DEFAULT_AI_EVALUATION_TASK_NAME,
                  ),
              ];
    }

    return [
        normalizeLoadedAIEvaluationTaskConfig(
            value,
            fallbackRouteName,
            DEFAULT_AI_EVALUATION_TASK_ID,
            DEFAULT_AI_EVALUATION_TASK_NAME,
        ),
    ];
}

function normalizeLoadedAICleaningToolConfig(
    value: unknown,
    fallback: AICleaningToolConfig,
    fallbackRouteName: string,
    toolKey: AICleaningToolKey,
): AICleaningToolConfig {
    if (!value || typeof value !== "object") {
        return {
            ...cloneAICleaningToolConfig(fallback),
            routeName: fallbackRouteName,
        };
    }

    const candidate = value as Partial<AICleaningToolConfig>;
    const submitFieldKeys = Array.isArray(candidate.submitFieldKeys)
        ? candidate.submitFieldKeys.filter(
              (item): item is string => typeof item === "string",
          )
        : [];
    const allowedOutputKeys = new Set(
        AI_CLEANING_TOOL_LABELS[toolKey].outputKeys,
    );
    const fallbackOutputMap = new Map(
        fallback.outputMappings.map((item) => [item.outputKey, item]),
    );
    const candidateOutputMap = new Map<string, AICleaningOutputMapping>();
    if (Array.isArray(candidate.outputMappings)) {
        candidate.outputMappings.forEach((item) => {
            if (!item || typeof item !== "object") {
                return;
            }
            const outputKey = (item as { outputKey?: unknown }).outputKey;
            const targetFieldKey = (item as { targetFieldKey?: unknown })
                .targetFieldKey;
            if (
                typeof outputKey !== "string" ||
                !allowedOutputKeys.has(outputKey.trim())
            ) {
                return;
            }
            candidateOutputMap.set(outputKey.trim(), {
                outputKey: outputKey.trim(),
                targetFieldKey:
                    typeof targetFieldKey === "string" ? targetFieldKey : "",
            });
        });
    }
    const outputMappings = AI_CLEANING_TOOL_LABELS[toolKey].outputKeys.map(
        (outputKey) =>
            cloneAICleaningOutputMapping(
                candidateOutputMap.get(outputKey) ??
                    fallbackOutputMap.get(outputKey) ?? {
                        outputKey,
                        targetFieldKey: "",
                    },
            ),
    );
    const routeName =
        typeof candidate.routeName === "string" && candidate.routeName.trim()
            ? candidate.routeName.trim()
            : fallbackRouteName;

    return {
        routeName,
        submitFieldKeys,
        prompt:
            typeof candidate.prompt === "string" &&
            candidate.prompt.trim().length > 0
                ? candidate.prompt
                : fallback.prompt,
        autoFillEnabled: candidate.autoFillEnabled === true,
        outputMappings,
    };
}

export function normalizeLoadedAIDetectConfig(value: unknown): AIDetectConfig {
    if (!value || typeof value !== "object") {
        return createDefaultAIDetectConfig();
    }

    const candidate = value as {
        stages?: unknown;
        evaluation?: unknown;
        evaluationTasks?: unknown;
        providers?: unknown;
        routes?: unknown;
        profiles?: unknown;
        cleaning?: unknown;
    };
    const providers = normalizeLoadedProviders(
        candidate.providers ?? candidate.profiles,
    );
    const resolvedProviders =
        providers.length > 0
            ? providers
            : [cloneAIProviderEndpoint(DEFAULT_AI_PROVIDER)];
    const routes = normalizeLoadedRoutes(
        candidate.routes,
        resolvedProviders[0]?.name ?? DEFAULT_AI_PROVIDER_NAME,
    );
    const resolvedRoutes =
        routes.length > 0 ? routes : [cloneAIModelRoute(DEFAULT_AI_ROUTE)];
    const stages = {} as AIDetectStageConfigMap;
    const cleaning = {} as AICleaningConfigMap;
    AI_STAGE_ORDER.forEach((stageKey) => {
        stages[stageKey] = normalizeLoadedAIDetectStageConfig(
            candidate.stages && typeof candidate.stages === "object"
                ? (candidate.stages as Record<string, unknown>)[stageKey]
                : null,
            DEFAULT_AI_STAGE_CONFIGS[stageKey],
            resolvedRoutes[0]?.name ?? DEFAULT_AI_ROUTE_NAME,
        );
    });
    AI_CLEANING_TOOL_ORDER.forEach((toolKey) => {
        cleaning[toolKey] = normalizeLoadedAICleaningToolConfig(
            candidate.cleaning && typeof candidate.cleaning === "object"
                ? (candidate.cleaning as Record<string, unknown>)[toolKey]
                : null,
            DEFAULT_AI_CLEANING_CONFIGS[toolKey],
            resolvedRoutes[0]?.name ?? DEFAULT_AI_ROUTE_NAME,
            toolKey,
        );
    });
    return {
        providers: resolvedProviders,
        routes: resolvedRoutes,
        stages,
        evaluationTasks: normalizeLoadedAIEvaluationTaskList(
            candidate.evaluationTasks ?? candidate.evaluation,
            resolvedRoutes[0]?.name ?? DEFAULT_AI_ROUTE_NAME,
        ),
        chat: normalizeLoadedAIChatConfig(
            (candidate as { chat?: unknown }).chat,
            resolvedRoutes[0]?.name ?? DEFAULT_AI_ROUTE_NAME,
        ),
        cleaning,
    };
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
    const providers =
        config.providers && config.providers.length > 0
            ? config.providers.map(cloneAIProviderEndpoint)
            : [cloneAIProviderEndpoint(DEFAULT_AI_PROVIDER)];
    const routes =
        config.routes && config.routes.length > 0
            ? config.routes.map(cloneAIModelRoute)
            : [cloneAIModelRoute(DEFAULT_AI_ROUTE)];
    const fallbackRouteName = routes[0]?.name ?? DEFAULT_AI_ROUTE_NAME;
    const routeNames = new Set(routes.map((item) => item.name));
    const keySet = new Set(columns.map((column) => column.key));

    const stages = {} as AIDetectStageConfigMap;
    const cleaning = {} as AICleaningConfigMap;
    AI_STAGE_ORDER.forEach((stageKey) => {
        const stageConfig =
            config.stages?.[stageKey] ?? DEFAULT_AI_STAGE_CONFIGS[stageKey];
        const normalizedStage = normalizeAIDetectStageConfigForColumns(
            stageConfig,
            columns,
        );
        const normalizedRouteName =
            normalizedStage.routeName &&
            routeNames.has(normalizedStage.routeName)
                ? normalizedStage.routeName
                : fallbackRouteName;
        stages[stageKey] = {
            ...normalizedStage,
            routeName: normalizedRouteName,
        };
    });

    const evaluationTasks = (
        config.evaluationTasks ?? [DEFAULT_AI_EVALUATION_TASK]
    ).map((task, index) => ({
        id:
            typeof task.id === "string" && task.id.trim().length > 0
                ? task.id.trim()
                : `${DEFAULT_AI_EVALUATION_TASK_ID}-${index + 1}`,
        name:
            typeof task.name === "string" && task.name.trim().length > 0
                ? task.name.trim()
                : index === 0
                  ? DEFAULT_AI_EVALUATION_TASK_NAME
                  : `评测配置 ${index + 1}`,
        enabled: task.enabled === true,
        attemptCount: normalizeAIEvaluationAttemptCount(task.attemptCount),
        maxConcurrency: normalizeAIEvaluationMaxConcurrency(
            task.maxConcurrency,
        ),
        answerGeneration: {
            routeName:
                task.answerGeneration.routeName &&
                routeNames.has(task.answerGeneration.routeName)
                    ? task.answerGeneration.routeName
                    : fallbackRouteName,
            prompt:
                typeof task.answerGeneration.prompt === "string" &&
                task.answerGeneration.prompt.trim().length > 0
                    ? task.answerGeneration.prompt
                    : DEFAULT_AI_EVALUATION_TASK.answerGeneration.prompt,
            questionFieldKeys: task.answerGeneration.questionFieldKeys.filter(
                (key) => keySet.has(key),
            ),
        },
        answerJudgment: {
            routeName:
                task.answerJudgment.routeName &&
                routeNames.has(task.answerJudgment.routeName)
                    ? task.answerJudgment.routeName
                    : fallbackRouteName,
            prompt:
                typeof task.answerJudgment.prompt === "string" &&
                task.answerJudgment.prompt.trim().length > 0
                    ? task.answerJudgment.prompt
                    : DEFAULT_AI_EVALUATION_TASK.answerJudgment.prompt,
            answerFieldKeys: task.answerJudgment.answerFieldKeys.filter((key) =>
                keySet.has(key),
            ),
        },
    }));

    const chatConfig = config.chat ?? DEFAULT_AI_CHAT_CONFIG;
    const normalizedChatRouteName =
        chatConfig.routeName && routeNames.has(chatConfig.routeName)
            ? chatConfig.routeName
            : fallbackRouteName;
    AI_CLEANING_TOOL_ORDER.forEach((toolKey) => {
        const toolConfig =
            config.cleaning?.[toolKey] ?? DEFAULT_AI_CLEANING_CONFIGS[toolKey];
        const normalizedRouteName =
            toolConfig.routeName && routeNames.has(toolConfig.routeName)
                ? toolConfig.routeName
                : fallbackRouteName;
        const outputMappings = AI_CLEANING_TOOL_LABELS[toolKey].outputKeys.map(
            (outputKey) => {
                const matched = toolConfig.outputMappings.find(
                    (item) => item.outputKey === outputKey,
                );
                return {
                    outputKey,
                    targetFieldKey:
                        matched?.targetFieldKey &&
                        keySet.has(matched.targetFieldKey)
                            ? matched.targetFieldKey
                            : "",
                };
            },
        );
        cleaning[toolKey] = {
            routeName: normalizedRouteName,
            prompt:
                typeof toolConfig.prompt === "string" &&
                toolConfig.prompt.trim().length > 0
                    ? toolConfig.prompt
                    : DEFAULT_AI_CLEANING_CONFIGS[toolKey].prompt,
            autoFillEnabled: toolConfig.autoFillEnabled === true,
            submitFieldKeys: toolConfig.submitFieldKeys.filter((key) =>
                keySet.has(key),
            ),
            outputMappings,
        };
    });

    return {
        providers,
        routes,
        stages,
        evaluationTasks,
        chat: {
            routeName: normalizedChatRouteName,
            prompt:
                typeof chatConfig.prompt === "string" &&
                chatConfig.prompt.trim().length > 0
                    ? chatConfig.prompt
                    : DEFAULT_AI_CHAT_CONFIG.prompt,
            defaultSubmitFieldKeys: chatConfig.defaultSubmitFieldKeys.filter(
                (key) => keySet.has(key),
            ),
        },
        cleaning,
    };
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
    const result: NamedAIDetectConfig[] = [];
    const usedNames = new Set<string>();
    value.forEach((item) => {
        if (!item || typeof item !== "object") {
            return;
        }
        const candidate = item as {
            name?: unknown;
            config?: unknown;
            providers?: unknown;
            routes?: unknown;
            stages?: unknown;
        };
        const name = normalizeAIConfigName(candidate.name);
        if (usedNames.has(name)) {
            return;
        }
        const configSource =
            candidate.config && typeof candidate.config === "object"
                ? candidate.config
                : candidate;
        result.push({
            name,
            config: normalizeLoadedAIDetectConfig(configSource),
        });
        usedNames.add(name);
    });
    return result;
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
    const normalizedPreferred = normalizeAIConfigName(preferredName);
    if (configs.some((item) => item.name === normalizedPreferred)) {
        return normalizedPreferred;
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

async function requestAIStreamResult(
    endpoint: string,
    payload: Record<string, unknown>,
    options?: {
        signal?: AbortSignal;
        onAnswerChunk?: (chunk: string) => void;
        onThinkingChunk?: (chunk: string) => void;
        onChunk?: (chunk: string) => void;
        onPhaseChange?: (phase: AIStreamPhase) => void;
    },
): Promise<AIDetectStreamResult> {
    options?.onPhaseChange?.("requesting");

    const response = await fetch(endpoint, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
        },
        signal: options?.signal,
        body: JSON.stringify(payload),
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
    let currentPhase: AIStreamPhase = "requesting";

    const advancePhase = (nextPhase: AIStreamPhase) => {
        if (currentPhase !== nextPhase) {
            currentPhase = nextPhase;
            options?.onPhaseChange?.(nextPhase);
        }
    };

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
                        advancePhase("outputting");
                        answerText += event.text;
                        options?.onAnswerChunk?.(event.text);
                        options?.onChunk?.(event.text);
                        continue;
                    }
                    if (
                        event.type === "thinking" &&
                        typeof event.text === "string"
                    ) {
                        advancePhase("thinking");
                        thinkingText += event.text;
                        options?.onThinkingChunk?.(event.text);
                        continue;
                    }
                    if (event.type === "done") {
                        continue;
                    }
                } catch {
                    advancePhase("outputting");
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
                    advancePhase("outputting");
                    answerText += event.text;
                    options?.onAnswerChunk?.(event.text);
                    options?.onChunk?.(event.text);
                } else if (
                    event.type === "thinking" &&
                    typeof event.text === "string"
                ) {
                    advancePhase("thinking");
                    thinkingText += event.text;
                    options?.onThinkingChunk?.(event.text);
                }
            } catch {
                advancePhase("outputting");
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
                    advancePhase("outputting");
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
                advancePhase("outputting");
                answerText += chunkText;
                options?.onAnswerChunk?.(chunkText);
                options?.onChunk?.(chunkText);
            }
        }
    }

    advancePhase("completed");

    return {
        answerText,
        thinkingText,
    };
}

export async function requestAIDetectResult(
    payload: {
        routeName: string;
        prompt: string;
        fields: AIDetectFieldPayload[];
    },
    options?: {
        signal?: AbortSignal;
        onAnswerChunk?: (chunk: string) => void;
        onThinkingChunk?: (chunk: string) => void;
        onChunk?: (chunk: string) => void;
        onPhaseChange?: (phase: AIStreamPhase) => void;
    },
): Promise<AIDetectStreamResult> {
    return requestAIStreamResult("/api/ai-detect/stream", payload, options);
}

export async function requestAIChatResult(
    payload: {
        routeName: string;
        prompt: string;
        messages: AIChatMessagePayload[];
        fields: AIDetectFieldPayload[];
    },
    options?: {
        signal?: AbortSignal;
        onAnswerChunk?: (chunk: string) => void;
        onThinkingChunk?: (chunk: string) => void;
        onChunk?: (chunk: string) => void;
        onPhaseChange?: (phase: AIStreamPhase) => void;
    },
): Promise<AIDetectStreamResult> {
    return requestAIStreamResult("/api/ai-chat/stream", payload, options);
}
