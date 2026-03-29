import type {
    AIDetectConfig,
    AIChatConfig,
    AIModelRoute,
    AIProviderEndpoint,
    AIDetectStageConfig,
    AIDetectStageConfigMap,
    AIDetectStageKey,
} from "../types";
import type { AIBatchTaskState } from "./types";

export const ALL_FILTER_VALUE = "全部";
export const EMPTY_FILTER_VALUE = "__EMPTY_FILTER__";
export const EMPTY_FILTER_LABEL = "空值";
export const NON_EMPTY_FILTER_VALUE = "__NON_EMPTY_FILTER__";
export const NON_EMPTY_FILTER_LABEL = "非空值";
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
export const DEFAULT_MODELROUTER_OPENAI_URL =
    "https://routify.alibaba-inc.com/protocol/openai/v1/";
export const DEFAULT_MODELROUTER_GEMINI_URL =
    "https://routify.alibaba-inc.com/protocol/vertex/v1beta/";
export const DEFAULT_ANTHROPIC_URL = "";

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
        title: "第一阶段：基础规范与特征排查",
        shortTitle: "Pre-check",
        description: "检查题目硬伤、图片依赖、绝对化词语与全选风险。",
    },
    context_audit: {
        title: "第二阶段：客观性与主观题排查",
        shortTitle: "Subjectivity",
        description: "评估答案唯一性、主观性风险与模棱两可表述。",
    },
    independent_solving: {
        title: "第三阶段：AI 独立闭卷解答",
        shortTitle: "Blind Solve",
        description: "只基于题目信息独立作答并输出可解性判断。",
    },
    final_verdict: {
        title: "第四阶段：深度对齐与额外信息审查",
        shortTitle: "Deep Align",
        description: "核对答案一致性、额外信息依赖与解析逻辑质量。",
    },
};

export const DEFAULT_AI_STAGE_PROMPTS: Record<AIDetectStageKey, string> = {
    precheck: `第一阶段：基础规范与特征排查（Pre-check 增强版）
你是专业的题目质检专家，负责第一阶段：基础规范与特征排查（Pre-check）。

输入：题目信息 JSON（{{fields_json}}）。
说明：JSON 中通常包含题目文本、选项、专家给定的答案等字段，请你自行识别并使用。

你的任务分为四个独立的判断维度：

一、有效性检查（is_valid）
只检查以下问题：错别字/明显病句、选项内容完全重复、题干与选项明显矛盾、题型与作答方式不自洽。
注意：不要因为“需要图片”或“可能有争议”判定为无效。只有存在绝对的内容/结构硬伤时，才返回 false。

二、图片依赖性判断（requires_image）
仅靠当前提供的文字和选项，是否足以完整理解题意并作答。如果缺少关键图形信息导致无法作答，返回 true。这是独立判断，与 is_valid 无关。

三、绝对性词语排查（has_absolute_words）
检查选项文本中是否包含“必然”、“肯定”、“绝对”、“总是”、“一定”等绝对化表述。若有，提取这些词语所在的选项。

四、全选陷阱预警（is_all_selected_warning）
如果题目是选择题，且专家给定的答案为全选（如 ABCD全选），则返回 true。若非选择题或非全选，返回 false。

请严格以 JSON 格式输出，不要包含任何 Markdown 标记（如 \`\`\`json）或其他说明文本：
{
  "is_valid": boolean,
  "invalid_reason": "如果 is_valid 为 false，说明具体原因；否则为空字符串",
  "requires_image": boolean,
  "has_absolute_words": boolean,
  "absolute_words_details": "如果包含，简述哪个选项包含什么词；否则为空字符串",
  "is_all_selected_warning": boolean
}

字段内容如下：
{{fields_json}}`,
    context_audit: `第二阶段：客观性与主观题排查（Subjectivity Check）
你是专业的题目质检专家，负责第二阶段：客观性与歧义排查。

输入：题目信息 JSON（{{fields_json}}）。
说明：JSON 中通常包含题目文本、选项等字段，请你自行识别并使用。

你的核心任务是评估该题目的“答案唯一性”和“表述客观性”。
判断标准：
1. 主观性判断：题干或选项中是否包含无法量化、高度依赖个人主观认知或特定立场的表述（例如“最优美”、“通常让人感到不适”、“最好的方法”等无明确标准的概念）。
2. 伪命题/模棱两可：选项之间是否存在界限模糊，或者在学术上存在广泛争议，导致“怎么解释都有道理”的情况。

请严格以 JSON 格式输出，不要包含任何 Markdown 标记或其他说明文本：
{
  "is_objective": boolean,
  "subjectivity_risk_level": "高/中/低 (高代表极度主观，低代表完全客观有唯一解)",
  "analysis": "如果不是完全客观，请指出具体哪个表述过于主观或存在模棱两可的争议；如果客观，请填 '无'"
}

字段内容如下：
{{fields_json}}`,
    independent_solving: `第三阶段：AI 独立“闭卷”解答（Blind Solve）

你是一个参加考试的优秀学生。请根据提供的题目信息进行独立解答。
输入：题目信息 JSON（{{fields_json}}）。
说明：JSON 中通常包含题目文本、选项、图片描述或识别内容等字段，请你自行识别并使用。

你的任务：
1. 忽略任何外部提供的参考答案，仅凭上述输入信息进行解答。
2. 一步一步写出你的推理过程。
3. 给出你最终的解答或选择。

请严格以 JSON 格式输出，不要包含任何 Markdown 标记或其他说明文本：
{
  "ai_reasoning_step_by_step": "详细的解题推演过程",
  "ai_final_answer": "你的最终答案，例如 'A', 'A,B', 或填空题的最终文本",
  "can_be_solved": boolean,
  "unsolvable_reason": "如果由于信息缺失导致无法解题(can_be_solved为false)，请说明缺失了什么；否则为空"
}

字段内容如下：
{{fields_json}}`,
    final_verdict: `第四阶段：深度对齐与额外信息审查（Deep Alignment & Extra Info Review）

你是高级题目教研专家，负责第四阶段：答案对齐与解析深度审查。
输入数据：题目信息 JSON（{{fields_json}}）。
说明：JSON 中通常包含：
1. 原始题目：题干、选项
2. 专家数据：专家答案、专家解答过程
3. AI盲测数据：AI得出的答案、AI推理过程
请你自行识别并使用这些字段。

你的任务分为三个维度：

一、答案一致性（is_answer_consistent）
对比专家答案与 AI 得出的答案是否核心一致。

二、额外信息/超纲概念审查（has_extra_info）
仔细检查专家的【解答过程】。判断其推导是否引入了“额外信息”。
“额外信息”定义为：
- 不属于题目本身已知条件的设定。
- 不属于通用基础学科常识（如高中/大学基础生物、化学课本常识）。
- 属于特定学术论文中自定义的概念、前沿文献的特定结论，或极其冷门的小众知识点。
如果专家的解析必须依赖这些额外信息才能走通，则返回 true。

三、逻辑自洽与倒推审查（is_logic_forced）
分析专家的解答过程，是否存在“先知答案，强行凑过程”的逻辑倒置现象，或者逻辑链条存在明显断裂。

请严格以 JSON 格式输出，不要包含任何 Markdown 标记或其他说明文本：
{
  "is_answer_consistent": boolean,
  "has_extra_info": boolean,
  "extra_info_details": "如果 has_extra_info 为 true，详细指出专家解析中滥用了哪个不在常规知识体系内的额外概念或文献结论；否则为空",
  "is_logic_forced": boolean,
  "logic_flaw_details": "如果 is_logic_forced 为 true，指出专家解析中逻辑断裂或强行倒推的地方；否则为空",
  "final_verdict": "综合评价：'优秀' / '需修改解析' / '题目超纲需废弃' / '存在逻辑错误'"
}

字段内容如下：
{{fields_json}}`,
};

export const DEFAULT_AI_CHAT_PROMPT = `你是题目详情页中的 AI 助手。

你会收到两类信息：
1. 当前题目的固定字段上下文（可能包含题干、选项、答案、解析、图片说明等）
2. 用户和你的多轮聊天记录

请遵守以下要求：
- 优先基于当前题目上下文回答，不要脱离题目泛泛而谈。
- 如果用户的问题依赖当前字段中没有提供的信息，要明确说明缺失了什么。
- 回答尽量直接、准确、结构清晰。
- 如果字段中包含参考答案或解析，只有在用户问题确实相关时才引用，并说明依据来源于当前题目字段。
- 不要输出 JSON，也不要重复粘贴全部字段内容，除非用户明确要求。`;

export const DEFAULT_AI_PROVIDER_NAME = "默认提供商";
export const DEFAULT_AI_ROUTE_NAME = "gpt-5.4";
export const DEFAULT_AI_CONFIG_NAME = "默认配置";

export const DEFAULT_AI_PROVIDER: AIProviderEndpoint = {
    name: DEFAULT_AI_PROVIDER_NAME,
    apiType: "openai",
    apiUrl: DEFAULT_IDEALAB_OPENAI_URL,
    apiKey: "",
};

export const DEFAULT_AI_ROUTE: AIModelRoute = {
    name: DEFAULT_AI_ROUTE_NAME,
    model: "gpt-5.4-2026-03-05",
    reasoningEffort: "high",
    retryCount: 5,
    steps: [{ providerName: DEFAULT_AI_PROVIDER_NAME }],
};

const DEFAULT_AI_STAGE_BASE: Omit<AIDetectStageConfig, "prompt"> = {
    routeName: DEFAULT_AI_ROUTE_NAME,
    submitFieldKeys: [],
};

export const DEFAULT_AI_CHAT_CONFIG: AIChatConfig = {
    routeName: DEFAULT_AI_ROUTE_NAME,
    prompt: DEFAULT_AI_CHAT_PROMPT,
    defaultSubmitFieldKeys: [],
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
    providers: [DEFAULT_AI_PROVIDER],
    routes: [DEFAULT_AI_ROUTE],
    stages: DEFAULT_AI_STAGE_CONFIGS,
    chat: DEFAULT_AI_CHAT_CONFIG,
};

export const AI_REASONING_EFFORT_OPTIONS = ["low", "medium", "high"] as const;
export const AI_PROVIDER_API_TYPE_OPTIONS = [
    { value: "openai", label: "OpenAI 兼容" },
    { value: "gemini", label: "Gemini 原生" },
    { value: "anthropic", label: "Anthropic 原生" },
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
