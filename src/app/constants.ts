import type {
    AIDetectConfig,
    AIChatConfig,
    AICleaningConfigMap,
    AICleaningToolConfig,
    AICleaningToolKey,
    AIEvaluationTaskConfig,
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
export const MAX_ROW_REVIEW_COUNT = 3;
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

export const AI_CLEANING_TOOL_ORDER = [
    "generate_level3_tags",
    "level1_tag_classification",
    "biochem_level1_refine",
    "knowledge_point_tag_classification",
    "question_formatting",
] as const;

export const AI_CLEANING_TOOL_LABELS: Record<
    AICleaningToolKey,
    {
        title: string;
        shortTitle: string;
        description: string;
        outputKeys: string[];
    }
> = {
    generate_level3_tags: {
        title: "生成 level3 标签",
        shortTitle: "Level3 标签",
        description:
            "分析题目表征方法、表征类型，并输出最多 3 个代表性标签。",
        outputKeys: [
            "representation_method",
            "representation_type",
            "tags",
        ],
    },
    level1_tag_classification: {
        title: "Level1 标签分类",
        shortTitle: "Level1 分类",
        description:
            "结合题干、选项与图片，判断题目覆盖的 1-4 个核心学科标签。",
        outputKeys: ["level1", "confidence", "reason"],
    },
    biochem_level1_refine: {
        title: "细分生化 level1",
        shortTitle: "生化 Level1",
        description:
            "判断题目所属的生物学科方向，并给出置信度与判断依据。",
        outputKeys: ["discipline", "confidence", "reason"],
    },
    knowledge_point_tag_classification: {
        title: "知识点标签分类",
        shortTitle: "知识点分类",
        description:
            "直接提炼题目考察的核心知识点，并输出去重后的知识点标签。",
        outputKeys: ["knowledge_points", "reason"],
    },
    question_formatting: {
        title: "题目格式化",
        shortTitle: "题目格式化",
        description:
            "根据原题、选项和答案，规范输出 question_text、options、answer 三个字段。",
        outputKeys: ["question_text", "options", "answer"],
    },
};

export const DEFAULT_AI_CLEANING_PROMPTS: Record<AICleaningToolKey, string> = {
    generate_level3_tags: `你是一个专业的题目分析专家。请分析用户提交的题目，但不要回答题目内容。

你的任务是：
1. 分析题目涉及的表征方法（如：XRD、NMR、拉曼光谱、冷冻电镜、透射电镜、Western blot、红外光谱、质谱、荧光光谱、扫描电镜、原子力显微镜等）
2. 判断题目的表征类型（如：结构表征、成分分析、形貌观察、性能测试、生物检测等）
3. 提炼出最能代表题目特征的标签，最多 3 个

请严格按照以下 JSON 格式返回结果，不要包含 Markdown 标记或额外说明：
{
  "representation_method": "题目的表征方法描述",
  "representation_type": "题目的表征类型",
  "tags": ["XRD", "结构表征", "晶体分析"]
}

注意事项：
- 不要解答题目
- 不要输出题目答案
- 只分析题目的表征特征
- tags 数组最多包含 3 个元素

字段内容如下：
{{fields_json}}`,
    level1_tag_classification: `你是一个跨学科题目标注专家。请基于用户提交的题目字段内容，对题目进行 level1 标签分类，但不要解答题目。

你会看到题干、选项、图片说明或图片本身等信息。你的任务是判断这道题覆盖了哪些学科方向。

你只能从以下学科列表中选择，不允许输出列表外的学科名称：
- 分子生物学
- 细胞生物学
- 遗传学
- 发育生物学
- 生态学
- 生理学
- 神经生物学
- 免疫学
- 进化生物学
- 结构生物学
- 系统生物学
- 生物物理学
- 生物信息学
- 无机化学
- 分析化学
- 有机化学
- 物理化学
- 金属材料
- 无机非金属材料
- 高分子材料
- 复合材料
- 半导体材料
- 生物医用材料
- 纳米材料
- 电子信息材料
- 新能源材料

请遵守以下要求：
1. 综合题干、选项、图片和上下文信息进行判断。
2. 必须严格从上述列表中选择最匹配的 1-4 个学科标签。
3. 不要输出“生物”“化学”“材料”等列表外或过于宽泛的大类。
4. 多个学科请使用英文逗号加空格连接，例如 "分子生物学, 生物化学"。
5. 只有在题目确实体现跨学科时才返回多个学科，不要为了凑多个标签而过度扩展。
6. 若题目信息不足，也必须从列表中选择你认为最可能的 1 个学科，并在置信度中体现不确定性。
7. 不要输出题目答案，不要输出与分类无关的分析。

请严格按照以下 JSON 格式返回结果，不要包含 Markdown 标记或额外说明：
{
  "level1": "学科1, 学科2",
  "confidence": "高/中/低",
  "reason": "结合题干、选项、图片信息给出的简要分类依据"
}

补充说明：
- 若只能确定一个学科，"level1" 只返回一个标签。
- "level1" 中的每个标签都必须与上面的列表完全一致，不要改写名称，不要新增标签。

字段内容如下：
{{fields_json}}`,
    biochem_level1_refine: `你是一个专业的生物学科分类专家。请分析用户提交的生物方向题目，根据题目内容和图片判断该题目的学科研究方向。

你的任务是：
识别题目所属的学科领域，优先从以下学科中选择：
- 结构生物化学：涉及蛋白质结构、核酸结构、分子结构解析、晶体学等
- 分子生物学：涉及基因表达、DNA复制、转录翻译、基因调控等
- 细胞生物化学：涉及细胞信号转导、细胞代谢、细胞器功能、细胞周期等
- 系统生物化学：涉及代谢网络、生物系统调控、组学分析、通路分析等
- 酶学与生物催化：涉及酶反应机制、酶动力学、生物催化、酶工程等

如果题目不属于以上任何学科，请判断并返回该题目实际所属的生物学科名称。

请严格按照以下 JSON 格式返回结果，不要包含 Markdown 标记或额外说明：
{
  "discipline": "学科名称",
  "confidence": "高/中/低",
  "reason": "判断依据的简要说明"
}

字段内容如下：
{{fields_json}}`,
    knowledge_point_tag_classification: `你是一个学科知识点标注专家。请基于用户提交的题目字段内容，分析该题在每个相关学科下考察了哪些知识点，但不要解答题目。

你会看到题干、选项、图片说明或图片本身等信息。你的任务是提炼题目真正考察的知识点，重点面向生物、化学、材料相关题目。

---

【知识点的定义与描述规范】

知识点是指：学生需要掌握的学科概念、原理或规律，是解题所依托的知识基础。

知识点分为三个层次，请严格只提取"概念/原理层"：

| 层次         | 说明                             | 示例                          |
|--------------|----------------------------------|-------------------------------|
| ✅ 概念/原理层 | 学科中的核心概念、定义、原理     | 细胞膜的流动性、氧化还原反应  |
| ❌ 能力/行为层 | 描述考生需要做什么操作或判断     | 判断物质运输方式、分析实验结果 |
| ❌ 情境/应用层 | 题目包装的场景、实验或具体案例   | 某药物的跨膜运输过程          |

描述要求：
- 使用简洁名词或短语，不超过10个字
- 不使用"判断""分析""比较""计算"等动词开头
- 不描述题目的具体情境或实验名称
- 不写成完整句子

正确示例：\`跨膜运输方式\` \`酶的专一性\` \`共价键极性\` \`氧化还原反应\`
错误示例：\`判断物质运输方式\` \`分析酶的作用\` \`某实验中的氧化还原过程\`

---

【标注要求】

1. 先识别题目涉及的学科，再提炼每个学科下的核心知识点。
2. 每个学科知识点数量控制在 1-3 个。
3. 只保留题目明确涉及、能够从题干/选项/图片中得到支撑的知识点，不要过度联想。
4. "knowledge_points" 对所有学科知识点去重后平铺汇总，总数不超过 5 个。
5. 不要输出题目答案，不要输出与知识点分类无关的内容。

---

【输出格式】

请严格按照以下 JSON 格式返回结果，不要包含 Markdown 标记或额外说明：

{
  "knowledge_points_by_subject": [
    {
      "subject": "学科名称",
      "points": ["知识点1", "知识点2"]
    }
  ],
  "knowledge_points": ["知识点1", "知识点2"],
  "reason": "概括说明这些知识点为何与该题相关"
}

补充说明：
- 若信息不足，请返回空数组，并在 "reason" 中说明缺失了哪些关键信息。
- 学科名称应尽量与题目所属的 level1/二级学科表达保持一致。

---

字段内容如下：
{{fields_json}}`,
    question_formatting: `你是一个题库整理专家。现在有一些人工出题的题目文本、选项以及答案。由于是人工出题，可能在形式或需求表述上没有表达清楚（例如题目缺少对答案格式的具体要求）。

请在不改变核心考察知识点的前提下，结合给出的答案，对题目文本、选项和答案进行优化，使其符合正式题库的规范。

要求：

1. 题目文本（question_text）：

   - 应简洁清晰，剥离掉原本混在题目中的选项（如果有的话）。

   - 如果是填空题，使用适当的连续下划线（如：______）表示填空处。

   - 关键调整：请根据【原答案】的形式或内容反推题目的需求。例如，如果答案是一个数值范围，但题目中没有明确要求“求出 xxx 的范围”，请在题目中补充相关提示（如：“（请给出数值范围）”）；如果答案有特定单位或精度，请在题目末尾加上适当的要求提示，避免回答者因为题目需求不清而回答不到位。

2. 选项（options）：

   - 如果是选择题，请整理为标准的选项格式（A. xxx 换行 B. xxx ...），保持纯文本换行，不使用 Markdown 加粗或特殊符号。

   - 如果是填空/简答题，选项请输出 null 或空字符串。

3. 答案（answer）：

   - 同样进行格式规范化（如修正排版、去除非必要的冗余提示词等），确保其与优化后的题目要求完全匹配。

4. 修正所有文本中可能存在的明显排版问题（如多余换行、不规范的括号等）。

5. 注意：保留原有的 LaTeX 公式（例如 $A_3$、$M_s$ 等），不要将其转换为 Unicode 字符。

6. 必须严格以 JSON 格式输出，包含 question_text、options 和 answer 三个字段。

原题目文本：
{question_text}

原选项：
{options}

原答案：
{answer}

补充说明：
- 你收到的是题目相关字段 JSON。请自行从字段中识别“原题目文本”“原选项”“原答案”对应内容。
- 若 prompt 中的 {question_text}、{options}、{answer} 未被直接替换，请从下方字段 JSON 中提取对应内容完成任务。
- 最终只输出 JSON，不要附加任何解释。

字段内容如下：
{{fields_json}}`,
};

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

export const DEFAULT_AI_EVALUATION_GENERATION_PROMPT = `你是题目评测流程中的第一步答题模型。

你会收到题目相关字段 JSON。你的任务是：
1. 仅基于题目字段作答，不要参考标准答案。
2. 提炼关键解题依据。
3. 给出最终答案。
4. 如果信息不足，也要明确指出无法作答的原因。

请严格返回 JSON，不要输出 Markdown，不要输出额外说明文字，格式如下：
{
  "status": "answered" | "insufficient_information",
  "reasoning": "简明扼要的解题依据或无法作答原因",
  "final_answer": "最终答案；若无法作答则为空字符串",
  "confidence": "high" | "medium" | "low"
}

要求：
- "final_answer" 必须是可直接比较的最终作答结果，不要混入解释。
- 若为单选题，使用单个选项值，如 "A"。
- 若为多选题，使用统一格式如 "A,B"。
- 若为填空题，直接返回填空内容；若有多个空，使用稳定顺序返回，如 "空1: xxx；空2: yyy" 或字符串形式的 JSON 数组 "[\"xxx\",\"yyy\"]"，但同一任务内格式必须保持一致。
- 若题型无法明确识别，也要给出你认为最可比较的最终答案表达。
- 若无法根据题目字段完成作答，"status" 返回 "insufficient_information"。
- 不要补充任何 JSON 之外的内容。`;

export const DEFAULT_AI_EVALUATION_JUDGMENT_PROMPT = `你是题目评测流程中的第二步判定模型。

你会收到三类信息：
1. 题目字段 JSON
2. 第一步模型的结构化回答结果
3. 标准答案字段 JSON

你的任务是判断“第一步模型的最终答案是否正确”，并返回结构化判定结果。

请严格返回 JSON，不要输出 Markdown，不要输出额外说明文字，格式如下：
{
  "verdict": "correct" | "incorrect" | "undetermined",
  "score": 1,
  "reason": "简要说明判定依据；若无法判定也要说明原因",
  "reference_answer": "从标准答案字段中提取出的参考答案；若无法提取则为空字符串",
  "model_answer": "从第一步结果中提取出的模型答案；若无法提取则为空字符串"
}

要求：
- "verdict" 为 "correct" 时，"score" 必须为 1。
- "verdict" 为 "incorrect" 时，"score" 必须为 0。
- "verdict" 为 "undetermined" 时，"score" 也必须为 0，并在 "reason" 中说明无法判定的原因。
- 判定时优先比较最终答案，不要只比较推理过程是否相似。
- 不要补充任何 JSON 之外的内容。`;

export const AI_EVALUATION_STEP_LABELS = {
    answer_generation: {
        title: "第一步：题目作答",
        description: "基于题目字段独立生成模型回答。",
    },
    answer_judgment: {
        title: "第二步：答案判定",
        description: "结合模型回答与标准答案判断是否正确。",
    },
} as const;

export const DEFAULT_AI_EVALUATION_TASK_ID = "evaluation-task-1";
export const DEFAULT_AI_EVALUATION_TASK_NAME = "评测配置 1";

export const DEFAULT_AI_CHAT_CONFIG: AIChatConfig = {
    routeName: DEFAULT_AI_ROUTE_NAME,
    prompt: DEFAULT_AI_CHAT_PROMPT,
    defaultSubmitFieldKeys: [],
};

const DEFAULT_AI_CLEANING_TOOL_BASE: Omit<
    AICleaningToolConfig,
    "prompt" | "outputMappings"
> = {
    routeName: DEFAULT_AI_ROUTE_NAME,
    submitFieldKeys: [],
    autoFillEnabled: false,
};

export const DEFAULT_AI_CLEANING_CONFIGS: AICleaningConfigMap = {
    generate_level3_tags: {
        ...DEFAULT_AI_CLEANING_TOOL_BASE,
        prompt: DEFAULT_AI_CLEANING_PROMPTS.generate_level3_tags,
        outputMappings: AI_CLEANING_TOOL_LABELS.generate_level3_tags.outputKeys.map(
            (outputKey) => ({
                outputKey,
                targetFieldKey: "",
            }),
        ),
    },
    level1_tag_classification: {
        ...DEFAULT_AI_CLEANING_TOOL_BASE,
        prompt: DEFAULT_AI_CLEANING_PROMPTS.level1_tag_classification,
        outputMappings: AI_CLEANING_TOOL_LABELS.level1_tag_classification.outputKeys.map(
            (outputKey) => ({
                outputKey,
                targetFieldKey: outputKey === "level1" ? "level1" : "",
            }),
        ),
    },
    biochem_level1_refine: {
        ...DEFAULT_AI_CLEANING_TOOL_BASE,
        prompt: DEFAULT_AI_CLEANING_PROMPTS.biochem_level1_refine,
        outputMappings: AI_CLEANING_TOOL_LABELS.biochem_level1_refine.outputKeys.map(
            (outputKey) => ({
                outputKey,
                targetFieldKey: "",
            }),
        ),
    },
    knowledge_point_tag_classification: {
        ...DEFAULT_AI_CLEANING_TOOL_BASE,
        prompt: DEFAULT_AI_CLEANING_PROMPTS.knowledge_point_tag_classification,
        outputMappings: AI_CLEANING_TOOL_LABELS.knowledge_point_tag_classification.outputKeys.map(
            (outputKey) => ({
                outputKey,
                targetFieldKey: "",
            }),
        ),
    },
    question_formatting: {
        ...DEFAULT_AI_CLEANING_TOOL_BASE,
        prompt: DEFAULT_AI_CLEANING_PROMPTS.question_formatting,
        outputMappings: AI_CLEANING_TOOL_LABELS.question_formatting.outputKeys.map(
            (outputKey) => ({
                outputKey,
                targetFieldKey: "",
            }),
        ),
    },
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

export const DEFAULT_AI_EVALUATION_TASK: AIEvaluationTaskConfig = {
    id: DEFAULT_AI_EVALUATION_TASK_ID,
    name: DEFAULT_AI_EVALUATION_TASK_NAME,
    enabled: false,
    attemptCount: 1,
    maxConcurrency: 5,
    answerGeneration: {
        routeName: DEFAULT_AI_ROUTE_NAME,
        prompt: DEFAULT_AI_EVALUATION_GENERATION_PROMPT,
        questionFieldKeys: [],
    },
    answerJudgment: {
        routeName: DEFAULT_AI_ROUTE_NAME,
        prompt: DEFAULT_AI_EVALUATION_JUDGMENT_PROMPT,
        answerFieldKeys: [],
    },
};

export const DEFAULT_AI_CONFIG: AIDetectConfig = {
    providers: [DEFAULT_AI_PROVIDER],
    routes: [DEFAULT_AI_ROUTE],
    stages: DEFAULT_AI_STAGE_CONFIGS,
    evaluationTasks: [DEFAULT_AI_EVALUATION_TASK],
    chat: DEFAULT_AI_CHAT_CONFIG,
    cleaning: DEFAULT_AI_CLEANING_CONFIGS,
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
export const DEFAULT_AI_EVALUATION_ATTEMPT_COUNT = 1;
export const MIN_AI_EVALUATION_ATTEMPT_COUNT = 1;
export const MAX_AI_EVALUATION_ATTEMPT_COUNT = 10;
export const DEFAULT_AI_EVALUATION_MAX_CONCURRENCY = 5;
export const MIN_AI_EVALUATION_MAX_CONCURRENCY = 1;
export const MAX_AI_EVALUATION_MAX_CONCURRENCY = 10;
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
