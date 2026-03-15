import type {
    AIDetectConfig,
    AIDetectProfile,
    AIDetectStageConfig,
    AIDetectStageConfigMap,
    AIDetectStageKey,
    NamedAIDetectProfile,
} from "../types";
import type { AIBatchTaskState } from "./types";

export const ALL_FILTER_VALUE = "全部";
export const QUALIFIED_TITLE_ALIASES = ["是否合格"] as const;
export const TIME_TITLE_ALIASES = ["创建时间"] as const;
export const CREATOR_TITLE_ALIASES = ["创建人"] as const;
export const INSPECTOR_TITLE_ALIASES = ["质检员"] as const;
export const FEEDBACK_TITLE_ALIASES = [
    "业务反馈意见",
    "质检员业务反馈意见",
] as const;
export const OPENSOURCE_TITLE_ALIASES = ["是否开源"] as const;

export const DEFAULT_IDEALAB_OPENAI_URL =
    "https://idealab.alibaba-inc.com/api/openai/v1/";
export const DEFAULT_IDEALAB_GEMINI_URL =
    "https://idealab.alibaba-inc.com/api/vertex/v1beta/";

export const AI_STAGE_ORDER = [
    "precheck",
    "context_audit",
    "independent_solving",
    "final_verdict",
] as const;

export const AI_RUN_ALL_KEY = "all" as const;
export const AI_RUN_ALL_LABEL = "执行全部";
export const AI_RUN_STAGE_ORDER = [...AI_STAGE_ORDER, AI_RUN_ALL_KEY] as const;

export const AI_STAGE_LABELS: Record<
    AIDetectStageKey,
    { title: string; shortTitle: string; description: string }
> = {
    precheck: {
        title: "第一阶段：逻辑自洽性检查",
        shortTitle: "Pre-check",
        description: "检查题干与选项的逻辑闭环、错别字与图文依赖性。",
    },
    context_audit: {
        title: "第二阶段：多模态一致性审计",
        shortTitle: "Context Audit",
        description: "核对图文解答是否一致，是否引入题干外信息。",
    },
    independent_solving: {
        title: "第三阶段：AI 独立解题",
        shortTitle: "Independent Solving",
        description: "忽略原解，逐项推演选项并给出独立答案。",
    },
    final_verdict: {
        title: "第四阶段：答案最终裁定",
        shortTitle: "Final Verdict",
        description: "对比 AI 结论与原答案/解答，输出最终裁定。",
    },
};

export const DEFAULT_AI_STAGE_PROMPTS: Record<AIDetectStageKey, string> = {
    precheck:
        '你是题目质检助手，负责第一阶段：元数据与逻辑自洽性检查（Pre-check）。\n输入：题目文本、选项。\n题型说明：题目可能是选择题或填空题；若是选择题，可能为单选题，也可能为多选题。\n任务：\n1. 检查错别字、标点错误、选项重复、题干描述与选项矛盾。\n2. 判断图文依赖性：仅靠文字是否足以解题。\n3. 结合题型判断题干、作答方式与选项设置是否自洽，不要默认题目一定是单选题。\n输出必须为 JSON（不要输出多余文本或 Markdown）：\n{\n  "is_valid": boolean,\n  "reason": string,\n  "requires_image": boolean\n}\n要求：\n- is_valid 为 true 时，reason 填 \"无\" 或 \"\"。\n- is_valid 为 false 时，reason 必须写清楚具体问题。\n- requires_image 为 true 表示必须结合图片才能理解或解题。\n字段内容如下：\n{{fields_json}}',
    context_audit:
        '你是多模态一致性审计助手（Context Audit）。\n输入：题目文本、选项、图片、解答过程。\n题型说明：题目可能是选择题或填空题；若是选择题，可能为单选题，也可能为多选题。\n任务：\n1. 检查图片内容是否与文本描述匹配（例如文本说“如图所示三角形 ABC”，图片里是否真为 ABC）。\n2. 检查解答过程是否引入题目未提供的前提条件。\n3. 审计时结合题型判断解答是否与题目要求一致，不要默认只能有一个正确选项。\n输出必须为 JSON（不要输出多余文本或 Markdown）：\n{\n  "is_consistent": boolean,\n  "missing_info": string\n}\n要求：\n- 若一致且无外部依赖，missing_info 填 \"无\" 或 \"\"。\n- 若不一致或依赖外部信息，missing_info 写明缺失信息或矛盾点。\n字段内容如下：\n{{fields_json}}',
    independent_solving:
        '请忽略原有解答过程，独立解题并推演每个选项（Independent Solving）。\n输入：题目文本、选项、图片。\n题型说明：题目可能是选择题或填空题；若是选择题，可能为单选题，也可能为多选题。\n任务：\n1. 若题目有选项，针对每个选项逐一推导，说明为什么对/错。\n2. 若题目是填空题，则直接根据题干独立求解，不要强行按选择题处理。\n3. 给出 AI 独立计算的最终答案，并根据题意判断是单个答案、多个答案，还是填空结果。\n输出必须为 JSON（不要输出多余文本或 Markdown）：\n{\n  "analysis": { "A": "...", "B": "...", "C": "...", "D": "..." },\n  "final_answer": string\n}\n要求：\n- analysis 的键使用题目提供的选项标签（如 A/B/C/D 或 ①②③④）；如果是填空题且没有选项，可在 analysis 中按步骤组织关键推导。\n- 若为多选题，final_answer 应明确给出全部正确选项。\n- 若为填空题，final_answer 直接填写最终结果。\n- 若无法解题，请在 analysis 说明原因，并将 final_answer 写为 \"无法确定\"。\n字段内容如下：\n{{fields_json}}',
    final_verdict:
        '请进行真题对标与最终裁定（Final Verdict）。\n输入：步骤 3 的结果 + 题目原始答案 + 题目原始解答。\n任务：\n1. 对比 AI 答案与原始答案。\n2. 如果不一致，判断是 AI 错误还是原题答案/解答错误，并说明原因。\n输出必须为 JSON（不要输出多余文本或 Markdown）：\n{\n  "status": "Pass" | "Fail",\n  "discrepancy_detail": string\n}\n要求：\n- 一致则 status = \"Pass\"，discrepancy_detail 填 \"无\" 或 \"\"。\n- 不一致则 status = \"Fail\"，discrepancy_detail 说明冲突点与责任归因。\n字段内容如下：\n{{fields_json}}',
};

export const DEFAULT_AI_PROFILE_NAME = "默认接口";

export const DEFAULT_AI_PROFILE: AIDetectProfile = {
    provider: "openai",
    url: DEFAULT_IDEALAB_OPENAI_URL,
    model: "gpt-5.2",
    modelProvider: "openai",
    modelName: "gpt-5.2",
    apiKey: "",
    reasoningEffort: "high",
    retryCount: 5,
};

export const DEFAULT_AI_PROFILES: NamedAIDetectProfile[] = [
    { name: DEFAULT_AI_PROFILE_NAME, profile: DEFAULT_AI_PROFILE },
];

const DEFAULT_AI_STAGE_BASE: Omit<AIDetectStageConfig, "prompt"> = {
    profileName: DEFAULT_AI_PROFILE_NAME,
    submitFieldKeys: [],
};

export const DEFAULT_AI_STAGE_CONFIGS: AIDetectStageConfigMap = {
    precheck: {
        ...DEFAULT_AI_STAGE_BASE,
        prompt: DEFAULT_AI_STAGE_PROMPTS.precheck,
    },
    context_audit: {
        ...DEFAULT_AI_STAGE_BASE,
        prompt: DEFAULT_AI_STAGE_PROMPTS.context_audit,
    },
    independent_solving: {
        ...DEFAULT_AI_STAGE_BASE,
        prompt: DEFAULT_AI_STAGE_PROMPTS.independent_solving,
    },
    final_verdict: {
        ...DEFAULT_AI_STAGE_BASE,
        prompt: DEFAULT_AI_STAGE_PROMPTS.final_verdict,
    },
};

export const DEFAULT_AI_CONFIG: AIDetectConfig = {
    profiles: DEFAULT_AI_PROFILES,
    stages: DEFAULT_AI_STAGE_CONFIGS,
};

export const DEFAULT_AI_CONFIG_NAME = "默认配置";
export const AI_REASONING_EFFORT_OPTIONS = ["low", "medium", "high"] as const;
export const AI_PROVIDER_OPTIONS = [
    { value: "openai", label: "OpenAI 兼容 (Idealab)" },
    { value: "gemini", label: "Gemini 原生 (Idealab)" },
] as const;
export const DEFAULT_AI_RETRY_COUNT = 5;
export const MIN_AI_RETRY_COUNT = 0;
export const MAX_AI_RETRY_COUNT = 10;
export const DEFAULT_AI_BATCH_CONCURRENCY = 4;
export const MIN_AI_BATCH_CONCURRENCY = 1;
export const MAX_AI_BATCH_CONCURRENCY = 32;
export const LIST_PAGE_SIZE_OPTIONS = [10, 20, 50] as const;

export const INITIAL_AI_BATCH_TASK: AIBatchTaskState = {
    status: "idle",
    fileId: null,
    fileName: "",
    total: 0,
    completed: 0,
    success: 0,
    failed: 0,
    message: "",
};
