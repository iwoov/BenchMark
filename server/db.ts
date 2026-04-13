import Database from "better-sqlite3";
import path from "node:path";
import fs from "node:fs";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const dataDir = path.resolve(__dirname, "..", "data");
const backupDir = path.join(dataDir, "backups");

// Ensure data directory exists
if (!fs.existsSync(dataDir)) {
    fs.mkdirSync(dataDir, { recursive: true });
}
if (!fs.existsSync(backupDir)) {
    fs.mkdirSync(backupDir, { recursive: true });
}

const dbPath = path.join(dataDir, "benchmark.db");
const db = new Database(dbPath);

// Enable WAL mode for better performance
db.pragma("journal_mode = WAL");

// Create table
db.exec(`
  CREATE TABLE IF NOT EXISTS column_prefs (
    file_name TEXT PRIMARY KEY,
    selected_keys TEXT NOT NULL,
    field_signature TEXT,
    editable_keys TEXT,
    filter_keys TEXT
  );
`);

db.exec(`
  CREATE TABLE IF NOT EXISTS file_states (
    file_id TEXT PRIMARY KEY,
    file_name TEXT NOT NULL,
    state_json TEXT NOT NULL,
    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
  );
`);

export const DEFAULT_AI_CONFIG_NAME = "默认配置";
const DEFAULT_AI_PROFILE_NAME = "默认接口";
type AIProvider =
    | "openai"
    | "gemini"
    | "modelrouter-openai"
    | "modelrouter-gemini";
type AIReasoningEffort = "low" | "medium" | "high";
export type AIProviderApiType = "openai" | "gemini" | "anthropic";
export type AIDetectStageKey =
    | "precheck"
    | "context_audit"
    | "independent_solving"
    | "final_verdict";
export type AICleaningToolKey =
    | "generate_level3_tags"
    | "biochem_level1_refine";
const AI_STAGE_ORDER: AIDetectStageKey[] = [
    "precheck",
    "context_audit",
    "independent_solving",
    "final_verdict",
];
const AI_CLEANING_TOOL_ORDER: AICleaningToolKey[] = [
    "generate_level3_tags",
    "biochem_level1_refine",
];
const LEGACY_STAGE_KEY: AIDetectStageKey = "independent_solving";
const DEFAULT_AI_RETRY_COUNT = 5;
const MIN_AI_RETRY_COUNT = 0;
const MAX_AI_RETRY_COUNT = 10;
const DEFAULT_AI_EVALUATION_ATTEMPT_COUNT = 1;
const MIN_AI_EVALUATION_ATTEMPT_COUNT = 1;
const MAX_AI_EVALUATION_ATTEMPT_COUNT = 10;
const DEFAULT_AI_EVALUATION_MAX_CONCURRENCY = 5;
const MIN_AI_EVALUATION_MAX_CONCURRENCY = 1;
const MAX_AI_EVALUATION_MAX_CONCURRENCY = 10;
const DEFAULT_AI_PROVIDER_NAME = "默认提供商";
const DEFAULT_AI_ROUTE_NAME = "gpt-5.4";
const DEFAULT_AI_EVALUATION_TASK_ID = "evaluation-task-1";
const DEFAULT_AI_EVALUATION_TASK_NAME = "评测配置 1";
const LEGACY_PROJECT_NAME = "专家出题表格.xlsx";
const MIGRATED_LEGACY_PROJECT_NAME = "化学/生物/材料多模态评测集";
const DEFAULT_AI_CHAT_PROMPT = `你是题目详情页中的 AI 助手。

你会收到两类信息：
1. 当前题目的固定字段上下文（可能包含题干、选项、答案、解析、图片说明等）
2. 用户和你的多轮聊天记录

请遵守以下要求：
- 优先基于当前题目上下文回答，不要脱离题目泛泛而谈。
- 如果用户的问题依赖当前字段中没有提供的信息，要明确说明缺失了什么。
- 回答尽量直接、准确、结构清晰。
- 如果字段中包含参考答案或解析，只有在用户问题确实相关时才引用，并说明依据来源于当前题目字段。
- 不要输出 JSON，也不要重复粘贴全部字段内容，除非用户明确要求。`;
const DEFAULT_AI_EVALUATION_GENERATION_PROMPT = `你是答题模型。

你会收到题目相关字段，请仅基于这些字段独立作答，不要参考标准答案。

请输出清晰的作答结论和必要推理。`;
const DEFAULT_AI_EVALUATION_JUDGMENT_PROMPT = `你是评测裁判模型。

你会收到题目字段、上一步模型的回答，以及标准答案字段。

请判断模型回答是否正确，并返回明确结论与简短依据。`;
const DEFAULT_AI_CLEANING_PROMPTS: Record<AICleaningToolKey, string> = {
    generate_level3_tags: `你是一个专业的题目分析专家。请分析用户提交的题目，但不要回答题目内容。

你的任务是：
1. 分析题目涉及的表征方法
2. 判断题目的表征类型
3. 提炼出最能代表题目特征的标签，最多 3 个

请严格按照以下 JSON 格式返回结果，不要包含 Markdown 标记或其他说明：
{
  "representation_method": "题目的表征方法描述",
  "representation_type": "题目的表征类型",
  "tags": ["XRD", "结构表征", "晶体分析"]
}

字段内容如下：
{{fields_json}}`,
    biochem_level1_refine: `你是一个专业的生物学科分类专家。请分析用户提交的生物方向题目，根据题目内容和图片判断该题目的学科研究方向。

你的任务是：
识别题目所属的学科领域，优先从以下学科中选择：
- 结构生物化学
- 分子生物学
- 细胞生物化学
- 系统生物化学
- 酶学与生物催化

请严格按照以下 JSON 格式返回结果，不要包含 Markdown 标记或其他说明：
{
  "discipline": "学科名称",
  "confidence": "高/中/低",
  "reason": "判断依据的简要说明"
}

字段内容如下：
{{fields_json}}`,
};
const AI_CLEANING_TOOL_OUTPUT_KEYS: Record<AICleaningToolKey, string[]> = {
    generate_level3_tags: [
        "representation_method",
        "representation_type",
        "tags",
    ],
    biochem_level1_refine: ["discipline", "confidence", "reason"],
};

function sanitizeBackupLabel(value: string): string {
    const normalized = value.trim().replace(/\s+/g, "_");
    const safe = normalized.replace(/[^a-zA-Z0-9._-]/g, "_");
    return safe.length > 0 ? safe.slice(0, 80) : "backup";
}

function formatBackupTimestamp(date: Date): string {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    const hours = String(date.getHours()).padStart(2, "0");
    const minutes = String(date.getMinutes()).padStart(2, "0");
    const seconds = String(date.getSeconds()).padStart(2, "0");
    return `${year}${month}${day}-${hours}${minutes}${seconds}`;
}

export async function createDatabaseBackup(label: string): Promise<string> {
    const fileName = `${formatBackupTimestamp(new Date())}-${sanitizeBackupLabel(label)}.db`;
    const destination = path.join(backupDir, fileName);
    await db.backup(destination);
    return destination;
}

function getTableColumns(tableName: string): string[] {
    const rows = db.prepare(`PRAGMA table_info(${tableName})`).all() as Array<{
        name: string;
    }>;
    return rows.map((row) => row.name);
}

function createAIDetectConfigTable(): void {
    db.exec(`
    CREATE TABLE IF NOT EXISTS ai_configs (
      file_name TEXT NOT NULL,
      config_name TEXT NOT NULL,
      provider TEXT NOT NULL DEFAULT 'openai',
      ai_url TEXT NOT NULL,
      ai_model TEXT NOT NULL,
      api_key TEXT NOT NULL,
      vertex_project TEXT NOT NULL DEFAULT '',
      vertex_location TEXT NOT NULL DEFAULT '',
      submit_field_keys TEXT NOT NULL,
      prompt TEXT NOT NULL,
      result_field_key TEXT,
      reasoning_effort TEXT NOT NULL DEFAULT 'high',
      retry_count INTEGER NOT NULL DEFAULT ${DEFAULT_AI_RETRY_COUNT},
      stages_json TEXT,
      profiles_json TEXT,
      is_active INTEGER NOT NULL DEFAULT 0,
      created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
      updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
      PRIMARY KEY (file_name, config_name)
    );
  `);
    db.exec(
        "CREATE INDEX IF NOT EXISTS idx_ai_configs_file_name ON ai_configs(file_name)",
    );
    db.exec(
        "CREATE INDEX IF NOT EXISTS idx_ai_configs_active ON ai_configs(file_name, is_active)",
    );
}

function normalizeAIDetectActiveFlag(): void {
    const rows = db
        .prepare("SELECT DISTINCT file_name FROM ai_configs")
        .all() as Array<{ file_name: string }>;
    const countStmt = db.prepare(
        "SELECT COUNT(1) AS count FROM ai_configs WHERE file_name = ? AND is_active = 1",
    );
    const activateLatestStmt = db.prepare(
        `UPDATE ai_configs
     SET is_active = CASE
       WHEN config_name = (
         SELECT config_name
         FROM ai_configs
         WHERE file_name = ?
         ORDER BY datetime(updated_at) DESC, config_name ASC
         LIMIT 1
       ) THEN 1
       ELSE 0
     END
     WHERE file_name = ?`,
    );
    const keepLatestActiveStmt = db.prepare(
        `UPDATE ai_configs
     SET is_active = CASE
       WHEN config_name = (
         SELECT config_name
         FROM ai_configs
         WHERE file_name = ? AND is_active = 1
         ORDER BY datetime(updated_at) DESC, config_name ASC
         LIMIT 1
       ) THEN 1
       ELSE 0
     END
     WHERE file_name = ?`,
    );

    for (const row of rows) {
        const countRow = countStmt.get(row.file_name) as { count: number };
        const activeCount = Number(countRow.count);
        if (activeCount === 0) {
            activateLatestStmt.run(row.file_name, row.file_name);
            continue;
        }
        if (activeCount > 1) {
            keepLatestActiveStmt.run(row.file_name, row.file_name);
        }
    }
}

function migrateLegacyAIDetectConfigTable(): void {
    db.exec("DROP TABLE IF EXISTS ai_configs_legacy");
    db.exec("ALTER TABLE ai_configs RENAME TO ai_configs_legacy");
    createAIDetectConfigTable();

    db.prepare(
        `INSERT INTO ai_configs (
      file_name,
      config_name,
      provider,
      ai_url,
      ai_model,
      api_key,
      vertex_project,
      vertex_location,
      submit_field_keys,
      prompt,
      result_field_key,
      reasoning_effort,
      retry_count,
      stages_json,
      is_active,
      created_at,
      updated_at
    )
    SELECT
      file_name,
      ?,
      'openai',
      ai_url,
      ai_model,
      api_key,
      '',
      '',
      submit_field_keys,
      prompt,
      result_field_key,
      'high',
      ${DEFAULT_AI_RETRY_COUNT},
      NULL,
      1,
      CURRENT_TIMESTAMP,
      CURRENT_TIMESTAMP
    FROM ai_configs_legacy`,
    ).run(DEFAULT_AI_CONFIG_NAME);

    db.exec("DROP TABLE ai_configs_legacy");
}

function ensureAIDetectConfigTable(): void {
    const tableExists = Boolean(
        db
            .prepare(
                "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = 'ai_configs'",
            )
            .get(),
    );

    if (!tableExists) {
        createAIDetectConfigTable();
        return;
    }

    const columns = new Set(getTableColumns("ai_configs"));
    if (!columns.has("config_name")) {
        migrateLegacyAIDetectConfigTable();
        return;
    }

    if (!columns.has("is_active")) {
        db.exec(
            "ALTER TABLE ai_configs ADD COLUMN is_active INTEGER NOT NULL DEFAULT 0",
        );
    }
    if (!columns.has("created_at")) {
        db.exec("ALTER TABLE ai_configs ADD COLUMN created_at TEXT");
        db.exec(
            "UPDATE ai_configs SET created_at = CURRENT_TIMESTAMP WHERE created_at IS NULL OR created_at = ''",
        );
    }
    if (!columns.has("updated_at")) {
        db.exec("ALTER TABLE ai_configs ADD COLUMN updated_at TEXT");
        db.exec(
            "UPDATE ai_configs SET updated_at = CURRENT_TIMESTAMP WHERE updated_at IS NULL OR updated_at = ''",
        );
    }
    if (!columns.has("reasoning_effort")) {
        db.exec(
            "ALTER TABLE ai_configs ADD COLUMN reasoning_effort TEXT NOT NULL DEFAULT 'high'",
        );
    }
    if (!columns.has("retry_count")) {
        db.exec(
            `ALTER TABLE ai_configs ADD COLUMN retry_count INTEGER NOT NULL DEFAULT ${DEFAULT_AI_RETRY_COUNT}`,
        );
    }
    if (!columns.has("stages_json")) {
        db.exec("ALTER TABLE ai_configs ADD COLUMN stages_json TEXT");
    }
    if (!columns.has("profiles_json")) {
        db.exec("ALTER TABLE ai_configs ADD COLUMN profiles_json TEXT");
    }
    if (!columns.has("provider")) {
        db.exec(
            "ALTER TABLE ai_configs ADD COLUMN provider TEXT NOT NULL DEFAULT 'openai'",
        );
    }
    if (!columns.has("vertex_project")) {
        db.exec(
            "ALTER TABLE ai_configs ADD COLUMN vertex_project TEXT NOT NULL DEFAULT ''",
        );
    }
    if (!columns.has("vertex_location")) {
        db.exec(
            "ALTER TABLE ai_configs ADD COLUMN vertex_location TEXT NOT NULL DEFAULT ''",
        );
    }
    db.exec(
        "UPDATE ai_configs SET provider = 'openai' WHERE provider IS NULL OR trim(provider) = ''",
    );
    db.exec(
        "UPDATE ai_configs SET provider = 'gemini' WHERE provider = 'vertex'",
    );
    db.exec(
        "UPDATE ai_configs SET provider = 'openai' WHERE provider = 'idealab'",
    );
    db.exec(
        "UPDATE ai_configs SET provider = 'openai' WHERE provider NOT IN ('openai', 'gemini', 'modelrouter-openai', 'modelrouter-gemini')",
    );
    db.exec(
        "UPDATE ai_configs SET vertex_project = '' WHERE vertex_project IS NULL",
    );
    db.exec(
        "UPDATE ai_configs SET vertex_location = '' WHERE vertex_location IS NULL",
    );
    db.exec(
        "UPDATE ai_configs SET reasoning_effort = 'high' WHERE reasoning_effort IS NULL OR trim(reasoning_effort) = ''",
    );
    db.exec(
        `UPDATE ai_configs
     SET retry_count = ${DEFAULT_AI_RETRY_COUNT}
     WHERE retry_count IS NULL
       OR retry_count < ${MIN_AI_RETRY_COUNT}
       OR retry_count > ${MAX_AI_RETRY_COUNT}`,
    );

    createAIDetectConfigTable();
    normalizeAIDetectActiveFlag();
}

function createAIProviderEndpointTable(): void {
    db.exec(`
    CREATE TABLE IF NOT EXISTS ai_provider_endpoints (
      name TEXT PRIMARY KEY,
      api_type TEXT NOT NULL,
      api_url TEXT NOT NULL,
      api_key TEXT NOT NULL,
      created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
      updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    );
  `);
}

function createAIModelRouteTable(): void {
    db.exec(`
    CREATE TABLE IF NOT EXISTS ai_model_routes (
      name TEXT PRIMARY KEY,
      model TEXT NOT NULL,
      reasoning_effort TEXT NOT NULL DEFAULT 'high',
      retry_count INTEGER NOT NULL DEFAULT ${DEFAULT_AI_RETRY_COUNT},
      steps_json TEXT NOT NULL,
      created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
      updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    );
  `);
}

function createFileAIStageConfigTable(): void {
    db.exec(`
    CREATE TABLE IF NOT EXISTS file_ai_stage_configs (
      file_name TEXT PRIMARY KEY,
      stages_json TEXT NOT NULL,
      chat_json TEXT,
      evaluation_json TEXT,
      cleaning_json TEXT,
      updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
    );
  `);
}

function getAICleaningResultTableName(toolKey: AICleaningToolKey): string {
    return `ai_cleaning_${toolKey}_results`;
}

function getAIEvaluationResultTableName(): string {
    return "ai_evaluation_results";
}

function createAICleaningResultTable(toolKey: AICleaningToolKey): void {
    db.exec(`
    CREATE TABLE IF NOT EXISTS ${getAICleaningResultTableName(toolKey)} (
      file_id TEXT NOT NULL,
      file_name TEXT NOT NULL,
      row_id TEXT NOT NULL,
      response_text TEXT NOT NULL,
      parsed_json_text TEXT,
      updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
      PRIMARY KEY (file_id, row_id)
    );
  `);
    db.exec(
        `CREATE INDEX IF NOT EXISTS idx_${getAICleaningResultTableName(toolKey)}_file_id ON ${getAICleaningResultTableName(toolKey)}(file_id)`,
    );
}

function createAIEvaluationResultTable(): void {
    db.exec(`
    CREATE TABLE IF NOT EXISTS ${getAIEvaluationResultTableName()} (
      file_id TEXT NOT NULL,
      file_name TEXT NOT NULL,
      row_id TEXT NOT NULL,
      task_id TEXT NOT NULL,
      attempt_index INTEGER NOT NULL,
      generation_response_text TEXT NOT NULL,
      generation_parsed_json_text TEXT,
      judgment_response_text TEXT NOT NULL,
      judgment_parsed_json_text TEXT,
      final_verdict TEXT NOT NULL,
      updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
      PRIMARY KEY (file_id, row_id, task_id, attempt_index)
    );
  `);
    db.exec(
        `CREATE INDEX IF NOT EXISTS idx_${getAIEvaluationResultTableName()}_file_id ON ${getAIEvaluationResultTableName()}(file_id)`,
    );
}

function ensureAIRoutingTables(): void {
    createAIProviderEndpointTable();
    createAIModelRouteTable();
    createFileAIStageConfigTable();
    AI_CLEANING_TOOL_ORDER.forEach((toolKey) => {
        createAICleaningResultTable(toolKey);
    });
    createAIEvaluationResultTable();
    const columns = new Set(getTableColumns("file_ai_stage_configs"));
    if (!columns.has("chat_json")) {
        db.exec("ALTER TABLE file_ai_stage_configs ADD COLUMN chat_json TEXT");
    }
    if (!columns.has("evaluation_json")) {
        db.exec(
            "ALTER TABLE file_ai_stage_configs ADD COLUMN evaluation_json TEXT",
        );
    }
    if (!columns.has("cleaning_json")) {
        db.exec("ALTER TABLE file_ai_stage_configs ADD COLUMN cleaning_json TEXT");
    }
}

function createSyncDatabaseBackup(label: string): string {
    const fileName = `${formatBackupTimestamp(new Date())}-${sanitizeBackupLabel(label)}.db`;
    const destination = path.join(backupDir, fileName);
    fs.copyFileSync(dbPath, destination);
    return destination;
}

const tableColumns = db
    .prepare("PRAGMA table_info(column_prefs)")
    .all() as Array<{ name: string }>;
const hasFieldSignatureColumn = tableColumns.some(
    (column) => column.name === "field_signature",
);
const hasEditableKeysColumn = tableColumns.some(
    (column) => column.name === "editable_keys",
);
const hasFilterKeysColumn = tableColumns.some(
    (column) => column.name === "filter_keys",
);

if (!hasFieldSignatureColumn) {
    db.exec("ALTER TABLE column_prefs ADD COLUMN field_signature TEXT");
}
if (!hasEditableKeysColumn) {
    db.exec("ALTER TABLE column_prefs ADD COLUMN editable_keys TEXT");
}
if (!hasFilterKeysColumn) {
    db.exec("ALTER TABLE column_prefs ADD COLUMN filter_keys TEXT");
}
ensureAIDetectConfigTable();
ensureAIRoutingTables();
migrateLegacyAIConfigToRoutingTables();

export interface ColumnPrefsConfig {
    fieldSignature: string;
    displayKeys: string[];
    editableKeys: string[];
    filterKeys?: string[];
}

export interface PersistedFileState {
    fileId: string;
    fileName: string;
    state: unknown;
    updatedAt: string;
}

export interface AIDetectConfig {
    profiles: NamedAIDetectProfile[];
    stages: Record<AIDetectStageKey, AIDetectStageConfig>;
}

export interface NamedAIDetectConfig extends AIDetectConfig {
    name: string;
    isActive: boolean;
    updatedAt: string;
}

export interface AIDetectProfile {
    provider: AIProvider;
    url: string;
    model: string;
    modelProvider?: string;
    modelName?: string;
    apiKey: string;
    reasoningEffort: AIReasoningEffort;
    retryCount: number;
}

export interface NamedAIDetectProfile {
    name: string;
    profile: AIDetectProfile;
}

export interface AIDetectStageConfig {
    profileName: string;
    submitFieldKeys: string[];
    prompt: string;
    resultFieldKey: string;
}

interface LegacyAIDetectStageConfig {
    provider: AIProvider;
    url: string;
    model: string;
    apiKey: string;
    submitFieldKeys: string[];
    prompt: string;
    resultFieldKey: string;
    reasoningEffort: AIReasoningEffort;
    retryCount: number;
}

export interface AIProviderEndpointConfig {
    name: string;
    apiType: AIProviderApiType;
    apiUrl: string;
    apiKey: string;
    updatedAt?: string;
}

export interface AIModelRouteStepConfig {
    providerName: string;
}

export interface AIModelRouteConfig {
    name: string;
    model: string;
    reasoningEffort: AIReasoningEffort;
    retryCount: number;
    steps: AIModelRouteStepConfig[];
    updatedAt?: string;
}

export interface FileAIStageConfig {
    routeName: string;
    submitFieldKeys: string[];
    prompt: string;
}

export interface FileAIChatConfig {
    routeName: string;
    prompt: string;
    defaultSubmitFieldKeys: string[];
}

export interface FileAIEvaluationAnswerGenerationConfig {
    routeName: string;
    prompt: string;
    questionFieldKeys: string[];
}

export interface FileAIEvaluationAnswerJudgmentConfig {
    routeName: string;
    prompt: string;
    answerFieldKeys: string[];
}

export interface FileAIEvaluationTaskConfig {
    id: string;
    name: string;
    enabled: boolean;
    attemptCount: number;
    maxConcurrency: number;
    answerGeneration: FileAIEvaluationAnswerGenerationConfig;
    answerJudgment: FileAIEvaluationAnswerJudgmentConfig;
}

export interface FileAICleaningOutputMapping {
    outputKey: string;
    targetFieldKey: string;
}

export interface FileAICleaningToolConfig {
    routeName: string;
    submitFieldKeys: string[];
    prompt: string;
    autoFillEnabled: boolean;
    outputMappings: FileAICleaningOutputMapping[];
}

export type FileAICleaningConfigMap = Record<
    AICleaningToolKey,
    FileAICleaningToolConfig
>;

export interface FileAICleaningToolResult {
    responseText: string;
    parsedJsonText?: string;
    updatedAt?: string;
}

export interface FileAIEvaluationAttemptResult {
    attemptIndex: number;
    generationResponseText: string;
    generationParsedJsonText?: string;
    judgmentResponseText: string;
    judgmentParsedJsonText?: string;
    finalVerdict: string;
    updatedAt?: string;
}

function parseJsonStringArray(value: string | null | undefined): string[] {
    if (!value) {
        return [];
    }
    try {
        const parsed = JSON.parse(value) as unknown;
        if (!Array.isArray(parsed)) {
            return [];
        }
        return parsed.filter(
            (item): item is string => typeof item === "string",
        );
    } catch {
        return [];
    }
}

function normalizeReasoningEffort(
    value: string | null | undefined,
): AIReasoningEffort {
    if (value === "low" || value === "medium" || value === "high") {
        return value;
    }
    return "high";
}

function normalizeProviderApiType(
    value: unknown,
    fallback: AIProviderApiType = "openai",
): AIProviderApiType {
    if (value === "openai" || value === "gemini" || value === "anthropic") {
        return value;
    }
    if (
        value === "modelrouter-openai" ||
        value === "idealab" ||
        value === "openai-compatible"
    ) {
        return "openai";
    }
    if (value === "modelrouter-gemini" || value === "vertex") {
        return "gemini";
    }
    return fallback;
}

function normalizeAIProvider(value: string | null | undefined): AIProvider {
    if (
        value === "openai" ||
        value === "gemini" ||
        value === "modelrouter-openai" ||
        value === "modelrouter-gemini"
    ) {
        return value;
    }
    if (value === "vertex") {
        return "gemini";
    }
    if (value === "idealab") {
        return "openai";
    }
    return "openai";
}

function inferAIProviderFromUrl(
    provider: AIProvider,
    url: unknown,
): AIProvider {
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

function normalizeRetryCount(value: number | null | undefined): number {
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

function normalizeEvaluationAttemptCount(value: unknown): number {
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

function normalizeEvaluationMaxConcurrency(value: unknown): number {
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

function normalizeRouteStepList(
    value: unknown,
    fallbackProviderName: string,
): AIModelRouteStepConfig[] {
    if (!Array.isArray(value)) {
        return [{ providerName: fallbackProviderName }];
    }

    const steps = value
        .map((item) => {
            if (!item || typeof item !== "object") {
                return null;
            }
            const providerName = (item as { providerName?: unknown })
                .providerName;
            if (
                typeof providerName !== "string" ||
                providerName.trim().length === 0
            ) {
                return null;
            }
            return {
                providerName: providerName.trim(),
            };
        })
        .filter((item): item is AIModelRouteStepConfig => item !== null);

    return steps.length > 0 ? steps : [{ providerName: fallbackProviderName }];
}

function normalizeStageProvider(
    value: unknown,
    fallback: AIProvider,
): AIProvider {
    if (
        value === "openai" ||
        value === "gemini" ||
        value === "modelrouter-openai" ||
        value === "modelrouter-gemini"
    ) {
        return value;
    }
    if (value === "vertex") {
        return "gemini";
    }
    if (value === "idealab") {
        return "openai";
    }
    return fallback;
}

function normalizeStageReasoningEffort(
    value: unknown,
    fallback: AIReasoningEffort,
): AIReasoningEffort {
    if (value === "low" || value === "medium" || value === "high") {
        return value;
    }
    return fallback;
}

function normalizeProfile(
    value: unknown,
    fallback: AIDetectProfile,
): AIDetectProfile {
    if (!value || typeof value !== "object") {
        return { ...fallback };
    }
    const candidate = value as Partial<AIDetectProfile>;
    const provider = inferAIProviderFromUrl(
        normalizeStageProvider(candidate.provider, fallback.provider),
        candidate.url,
    );
    const reasoningEffort = normalizeStageReasoningEffort(
        candidate.reasoningEffort,
        fallback.reasoningEffort,
    );
    const retryCount =
        typeof candidate.retryCount === "number"
            ? normalizeRetryCount(candidate.retryCount)
            : fallback.retryCount;
    return {
        provider,
        url:
            typeof candidate.url === "string" && candidate.url.trim().length > 0
                ? candidate.url
                : fallback.url,
        model:
            typeof candidate.model === "string" &&
            candidate.model.trim().length > 0
                ? candidate.model
                : fallback.model,
        modelProvider:
            typeof candidate.modelProvider === "string"
                ? candidate.modelProvider
                : fallback.modelProvider,
        modelName:
            typeof candidate.modelName === "string"
                ? candidate.modelName
                : fallback.modelName,
        apiKey:
            typeof candidate.apiKey === "string"
                ? candidate.apiKey
                : fallback.apiKey,
        reasoningEffort,
        retryCount,
    };
}

function parseProfilesJson(
    value: string | null | undefined,
    fallback: AIDetectProfile,
): NamedAIDetectProfile[] {
    if (!value) {
        return [{ name: DEFAULT_AI_PROFILE_NAME, profile: fallback }];
    }
    try {
        const parsed = JSON.parse(value) as unknown;
        if (!Array.isArray(parsed)) {
            throw new Error("invalid");
        }
        const nameCount = new Map<string, number>();
        const profiles: NamedAIDetectProfile[] = [];
        parsed.forEach((item, index) => {
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
            profiles.push({
                name,
                profile: normalizeProfile(profileSource, fallback),
            });
        });
        return profiles.length > 0
            ? profiles
            : [{ name: DEFAULT_AI_PROFILE_NAME, profile: fallback }];
    } catch {
        return [{ name: DEFAULT_AI_PROFILE_NAME, profile: fallback }];
    }
}

function normalizeStageConfig(
    value: unknown,
    fallback: AIDetectStageConfig,
    fallbackProfileName: string,
): AIDetectStageConfig {
    if (!value || typeof value !== "object") {
        return {
            ...fallback,
            profileName: fallbackProfileName,
            submitFieldKeys: [...fallback.submitFieldKeys],
        };
    }

    const candidate = value as Partial<AIDetectStageConfig>;
    const submitFieldKeys = Array.isArray(candidate.submitFieldKeys)
        ? candidate.submitFieldKeys.filter(
              (item): item is string => typeof item === "string",
          )
        : [...fallback.submitFieldKeys];
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
        resultFieldKey:
            typeof candidate.resultFieldKey === "string"
                ? candidate.resultFieldKey
                : fallback.resultFieldKey,
    };
}

function normalizeLegacyStageConfig(
    value: unknown,
    fallback: LegacyAIDetectStageConfig,
): LegacyAIDetectStageConfig {
    if (!value || typeof value !== "object") {
        return {
            ...fallback,
            submitFieldKeys: [...fallback.submitFieldKeys],
        };
    }

    const candidate = value as Partial<LegacyAIDetectStageConfig>;
    const provider = inferAIProviderFromUrl(
        normalizeStageProvider(candidate.provider, fallback.provider),
        candidate.url,
    );
    const submitFieldKeys = Array.isArray(candidate.submitFieldKeys)
        ? candidate.submitFieldKeys.filter(
              (item): item is string => typeof item === "string",
          )
        : [...fallback.submitFieldKeys];
    const reasoningEffort = normalizeStageReasoningEffort(
        candidate.reasoningEffort,
        fallback.reasoningEffort,
    );
    const retryCount =
        typeof candidate.retryCount === "number"
            ? normalizeRetryCount(candidate.retryCount)
            : fallback.retryCount;

    return {
        provider,
        url:
            typeof candidate.url === "string" && candidate.url.trim().length > 0
                ? candidate.url
                : fallback.url,
        model:
            typeof candidate.model === "string" &&
            candidate.model.trim().length > 0
                ? candidate.model
                : fallback.model,
        apiKey:
            typeof candidate.apiKey === "string"
                ? candidate.apiKey
                : fallback.apiKey,
        submitFieldKeys,
        prompt:
            typeof candidate.prompt === "string" &&
            candidate.prompt.trim().length > 0
                ? candidate.prompt
                : fallback.prompt,
        resultFieldKey:
            typeof candidate.resultFieldKey === "string"
                ? candidate.resultFieldKey
                : fallback.resultFieldKey,
        reasoningEffort,
        retryCount,
    };
}

function parseStageConfigMap(
    value: string | null | undefined,
    fallback: AIDetectStageConfig,
    fallbackProfileName: string,
): Record<AIDetectStageKey, AIDetectStageConfig> {
    if (!value) {
        const stages = {} as Record<AIDetectStageKey, AIDetectStageConfig>;
        AI_STAGE_ORDER.forEach((stageKey) => {
            stages[stageKey] = normalizeStageConfig(
                null,
                fallback,
                fallbackProfileName,
            );
        });
        return stages;
    }

    try {
        const parsed = JSON.parse(value) as unknown;
        if (!parsed || typeof parsed !== "object") {
            throw new Error("invalid");
        }
        const rawStages = parsed as Record<string, unknown>;
        const stages = {} as Record<AIDetectStageKey, AIDetectStageConfig>;
        AI_STAGE_ORDER.forEach((stageKey) => {
            stages[stageKey] = normalizeStageConfig(
                rawStages[stageKey],
                fallback,
                fallbackProfileName,
            );
        });
        return stages;
    } catch {
        const stages = {} as Record<AIDetectStageKey, AIDetectStageConfig>;
        AI_STAGE_ORDER.forEach((stageKey) => {
            stages[stageKey] = normalizeStageConfig(
                null,
                fallback,
                fallbackProfileName,
            );
        });
        return stages;
    }
}

function parseLegacyStageConfigMap(
    value: string | null | undefined,
    fallback: LegacyAIDetectStageConfig,
): Record<AIDetectStageKey, LegacyAIDetectStageConfig> {
    if (!value) {
        const stages = {} as Record<
            AIDetectStageKey,
            LegacyAIDetectStageConfig
        >;
        AI_STAGE_ORDER.forEach((stageKey) => {
            stages[stageKey] = normalizeLegacyStageConfig(null, fallback);
        });
        return stages;
    }

    try {
        const parsed = JSON.parse(value) as unknown;
        if (!parsed || typeof parsed !== "object") {
            throw new Error("invalid");
        }
        const rawStages = parsed as Record<string, unknown>;
        const stages = {} as Record<
            AIDetectStageKey,
            LegacyAIDetectStageConfig
        >;
        AI_STAGE_ORDER.forEach((stageKey) => {
            stages[stageKey] = normalizeLegacyStageConfig(
                rawStages[stageKey],
                fallback,
            );
        });
        return stages;
    } catch {
        const stages = {} as Record<
            AIDetectStageKey,
            LegacyAIDetectStageConfig
        >;
        AI_STAGE_ORDER.forEach((stageKey) => {
            stages[stageKey] = normalizeLegacyStageConfig(null, fallback);
        });
        return stages;
    }
}

/**
 * Get saved column preferences for a given file name.
 * Returns null if not found.
 */
export function getColumnPrefs(fileName: string): ColumnPrefsConfig | null {
    const row = db
        .prepare(
            "SELECT selected_keys, field_signature, editable_keys, filter_keys FROM column_prefs WHERE file_name = ?",
        )
        .get(fileName) as
        | {
              selected_keys: string;
              field_signature: string | null;
              editable_keys: string | null;
              filter_keys: string | null;
          }
        | undefined;

    if (!row) {
        return null;
    }

    return {
        fieldSignature: row.field_signature ?? "",
        displayKeys: parseJsonStringArray(row.selected_keys),
        editableKeys: parseJsonStringArray(row.editable_keys),
        filterKeys:
            row.filter_keys === null || row.filter_keys === undefined
                ? undefined
                : parseJsonStringArray(row.filter_keys),
    };
}

/**
 * Save (upsert) column preferences for a given file name.
 */
export function saveColumnPrefs(
    fileName: string,
    config: ColumnPrefsConfig,
): void {
    db.prepare(
        `INSERT INTO column_prefs (file_name, selected_keys, field_signature, editable_keys, filter_keys)
     VALUES (?, ?, ?, ?, ?)
     ON CONFLICT(file_name) DO UPDATE SET
       selected_keys = excluded.selected_keys,
       field_signature = excluded.field_signature,
       editable_keys = excluded.editable_keys,
       filter_keys = excluded.filter_keys`,
    ).run(
        fileName,
        JSON.stringify(config.displayKeys),
        config.fieldSignature,
        JSON.stringify(config.editableKeys),
        JSON.stringify(config.filterKeys ?? []),
    );
}

export function listFileStates(): PersistedFileState[] {
    const rows = db
        .prepare(
            "SELECT file_id, file_name, state_json, updated_at FROM file_states ORDER BY datetime(updated_at) DESC",
        )
        .all() as Array<{
        file_id: string;
        file_name: string;
        state_json: string;
        updated_at: string;
    }>;

    return rows
        .map((row) => {
            try {
                return {
                    fileId: row.file_id,
                    fileName: row.file_name,
                    state: JSON.parse(row.state_json) as unknown,
                    updatedAt: row.updated_at,
                };
            } catch {
                return null;
            }
        })
        .filter((row): row is PersistedFileState => row !== null);
}

export function getFileState(fileId: string): PersistedFileState | null {
    const row = db
        .prepare(
            "SELECT file_id, file_name, state_json, updated_at FROM file_states WHERE file_id = ?",
        )
        .get(fileId) as
        | {
              file_id: string;
              file_name: string;
              state_json: string;
              updated_at: string;
          }
        | undefined;

    if (!row) {
        return null;
    }

    try {
        return {
            fileId: row.file_id,
            fileName: row.file_name,
            state: JSON.parse(row.state_json) as unknown,
            updatedAt: row.updated_at,
        };
    } catch {
        return null;
    }
}

function getFileStateByName(fileName: string): PersistedFileState | null {
    const row = db
        .prepare(
            `SELECT file_id, file_name, state_json, updated_at
             FROM file_states
             WHERE file_name = ?
             ORDER BY datetime(updated_at) DESC
             LIMIT 1`,
        )
        .get(fileName) as
        | {
              file_id: string;
              file_name: string;
              state_json: string;
              updated_at: string;
          }
        | undefined;

    if (!row) {
        return null;
    }

    try {
        return {
            fileId: row.file_id,
            fileName: row.file_name,
            state: JSON.parse(row.state_json) as unknown,
            updatedAt: row.updated_at,
        };
    } catch {
        return null;
    }
}

function readFileStateJson(fileId: string): unknown {
    const row = db
        .prepare("SELECT state_json FROM file_states WHERE file_id = ?")
        .get(fileId) as { state_json: string } | undefined;

    if (!row) {
        return null;
    }

    try {
        return JSON.parse(row.state_json) as unknown;
    } catch {
        return null;
    }
}

export function isProjectNameInUse(
    fileName: string,
    excludeFileId?: string,
): boolean {
    const owner = getFileStateByName(fileName);
    if (owner && owner.fileId !== excludeFileId) {
        return true;
    }

    if (!owner) {
        const hasStageConfig = Boolean(
            db
                .prepare(
                    "SELECT 1 FROM file_ai_stage_configs WHERE file_name = ? LIMIT 1",
                )
                .get(fileName),
        );
        if (hasStageConfig) {
            return true;
        }

        const hasAIConfig = Boolean(
            db.prepare("SELECT 1 FROM ai_configs WHERE file_name = ? LIMIT 1").get(
                fileName,
            ),
        );
        if (hasAIConfig) {
            return true;
        }
    }

    return false;
}

function renameProjectReferences(
    fileId: string,
    previousFileName: string,
    nextFileName: string,
): PersistedFileState | null {
    const existingState = readFileStateJson(fileId);
    if (!existingState || typeof existingState !== "object") {
        return null;
    }

    const nextState = {
        ...(existingState as Record<string, unknown>),
        fileName: nextFileName,
    };

    const tx = db.transaction(() => {
        db.prepare(
            `UPDATE file_states
             SET file_name = ?,
                 state_json = ?,
                 updated_at = CURRENT_TIMESTAMP
             WHERE file_id = ?`,
        ).run(nextFileName, JSON.stringify(nextState), fileId);

        db.prepare("UPDATE ai_configs SET file_name = ? WHERE file_name = ?").run(
            nextFileName,
            previousFileName,
        );

        db.prepare(
            `UPDATE file_ai_stage_configs
             SET file_name = ?,
                 updated_at = CURRENT_TIMESTAMP
             WHERE file_name = ?`,
        ).run(nextFileName, previousFileName);

        AI_CLEANING_TOOL_ORDER.forEach((toolKey) => {
            db.prepare(
                `UPDATE ${getAICleaningResultTableName(toolKey)}
                 SET file_name = ?,
                     updated_at = CURRENT_TIMESTAMP
                 WHERE file_id = ?`,
            ).run(nextFileName, fileId);
        });
    });

    tx();
    return getFileState(fileId);
}

export function saveFileState(
    fileId: string,
    fileName: string,
    state: unknown,
): void {
    db.prepare(
        `INSERT INTO file_states (file_id, file_name, state_json, updated_at)
     VALUES (?, ?, ?, CURRENT_TIMESTAMP)
     ON CONFLICT(file_id) DO UPDATE SET
       file_name = excluded.file_name,
       state_json = excluded.state_json,
       updated_at = CURRENT_TIMESTAMP`,
    ).run(fileId, fileName, JSON.stringify(state));
}

export function renameProject(
    fileId: string,
    nextFileName: string,
): PersistedFileState | null {
    const existing = getFileState(fileId);
    if (!existing) {
        return null;
    }

    if (existing.fileName === nextFileName) {
        return existing;
    }

    return renameProjectReferences(fileId, existing.fileName, nextFileName);
}

export function updateFileStateAIResults(
    fileId: string,
    stageKey: AIDetectStageKey,
    results: Array<{ rowId: string; resultText: string }>,
): number | null {
    if (results.length === 0) {
        return 0;
    }
    const row = db
        .prepare("SELECT state_json FROM file_states WHERE file_id = ?")
        .get(fileId) as { state_json: string } | undefined;
    if (!row) {
        return null;
    }
    let parsedState: unknown = null;
    try {
        parsedState = JSON.parse(row.state_json) as unknown;
    } catch {
        return null;
    }
    if (!parsedState || typeof parsedState !== "object") {
        return null;
    }
    const state = parsedState as {
        rows?: Array<Record<string, unknown>>;
        clientStateVersion?: unknown;
    };
    if (!Array.isArray(state.rows)) {
        return null;
    }

    const rowMap = new Map<string, Record<string, unknown>>();
    state.rows.forEach((item) => {
        const rowId = item?.rowId;
        if (typeof rowId === "string") {
            rowMap.set(rowId, item);
        }
    });

    let updatedCount = 0;
    results.forEach(({ rowId, resultText }) => {
        const target = rowMap.get(rowId);
        if (!target) {
            return;
        }
        const rawAIResults = target.aiResults;
        const aiResults =
            rawAIResults && typeof rawAIResults === "object"
                ? (rawAIResults as Record<string, string>)
                : {};
        aiResults[stageKey] = resultText;
        target.aiResults = aiResults;
        updatedCount += 1;
    });

    if (updatedCount === 0) {
        return 0;
    }

    const currentVersion =
        typeof state.clientStateVersion === "number" &&
        Number.isFinite(state.clientStateVersion)
            ? Math.trunc(state.clientStateVersion)
            : 0;
    state.clientStateVersion = Math.max(Date.now(), currentVersion + 1);

    db.prepare(
        "UPDATE file_states SET state_json = ?, updated_at = CURRENT_TIMESTAMP WHERE file_id = ?",
    ).run(JSON.stringify(state), fileId);

    return updatedCount;
}

export function deleteFileState(fileId: string): void {
    db.prepare("DELETE FROM file_states WHERE file_id = ?").run(fileId);
}

function normalizeConfigName(name: string): string {
    const trimmed = name.trim();
    return trimmed.length > 0 ? trimmed : DEFAULT_AI_CONFIG_NAME;
}

export function listAIDetectConfigs(fileName: string): {
    configs: NamedAIDetectConfig[];
    activeConfigName: string;
} {
    const rows = db
        .prepare(
            `SELECT
         config_name,
         provider,
         ai_url,
         ai_model,
         api_key,
         vertex_project,
         vertex_location,
         submit_field_keys,
         prompt,
         result_field_key,
         reasoning_effort,
         retry_count,
         stages_json,
         profiles_json,
         is_active,
         updated_at
       FROM ai_configs
       WHERE file_name = ?
       ORDER BY is_active DESC, datetime(updated_at) DESC, config_name ASC`,
        )
        .all(fileName) as Array<{
        config_name: string;
        provider: string | null;
        ai_url: string;
        ai_model: string;
        api_key: string;
        vertex_project: string | null;
        vertex_location: string | null;
        submit_field_keys: string;
        prompt: string;
        result_field_key: string | null;
        reasoning_effort: string | null;
        retry_count: number | null;
        stages_json: string | null;
        profiles_json: string | null;
        is_active: number;
        updated_at: string;
    }>;

    const configs = rows.map((row) => {
        const legacyStageConfig: LegacyAIDetectStageConfig = {
            provider: normalizeAIProvider(row.provider),
            url: row.ai_url,
            model: row.ai_model,
            apiKey: row.api_key,
            submitFieldKeys: parseJsonStringArray(row.submit_field_keys),
            prompt: row.prompt,
            resultFieldKey: row.result_field_key ?? "",
            reasoningEffort: normalizeReasoningEffort(row.reasoning_effort),
            retryCount: normalizeRetryCount(row.retry_count),
        };
        const fallbackProfile: AIDetectProfile = {
            provider: legacyStageConfig.provider,
            url: legacyStageConfig.url,
            model: legacyStageConfig.model,
            apiKey: legacyStageConfig.apiKey,
            reasoningEffort: legacyStageConfig.reasoningEffort,
            retryCount: legacyStageConfig.retryCount,
        };
        const fallbackProfileName = DEFAULT_AI_PROFILE_NAME;
        let profiles = parseProfilesJson(row.profiles_json, fallbackProfile);
        let stages: Record<AIDetectStageKey, AIDetectStageConfig> | null = null;

        if (row.stages_json) {
            try {
                const parsed = JSON.parse(row.stages_json) as unknown;
                if (!parsed || typeof parsed !== "object") {
                    throw new Error("invalid");
                }
                const rawStages = parsed as Record<string, unknown>;
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

                if (hasProfileName) {
                    const fallbackStage: AIDetectStageConfig = {
                        profileName: profiles[0]?.name ?? fallbackProfileName,
                        submitFieldKeys: legacyStageConfig.submitFieldKeys,
                        prompt: legacyStageConfig.prompt,
                        resultFieldKey: legacyStageConfig.resultFieldKey,
                    };
                    stages = parseStageConfigMap(
                        row.stages_json,
                        fallbackStage,
                        profiles[0]?.name ?? fallbackProfileName,
                    );
                } else {
                    const legacyStages = parseLegacyStageConfigMap(
                        row.stages_json,
                        legacyStageConfig,
                    );
                    const derivedProfiles: NamedAIDetectProfile[] = [];
                    const derivedStages = {} as Record<
                        AIDetectStageKey,
                        AIDetectStageConfig
                    >;
                    AI_STAGE_ORDER.forEach((stageKey, index) => {
                        const legacyStage = legacyStages[stageKey];
                        const profileName =
                            AI_STAGE_ORDER.length > 1
                                ? `${DEFAULT_AI_PROFILE_NAME}-${index + 1}`
                                : DEFAULT_AI_PROFILE_NAME;
                        derivedProfiles.push({
                            name: profileName,
                            profile: {
                                provider: legacyStage.provider,
                                url: legacyStage.url,
                                model: legacyStage.model,
                                apiKey: legacyStage.apiKey,
                                reasoningEffort: legacyStage.reasoningEffort,
                                retryCount: legacyStage.retryCount,
                            },
                        });
                        derivedStages[stageKey] = {
                            profileName,
                            submitFieldKeys: legacyStage.submitFieldKeys,
                            prompt: legacyStage.prompt,
                            resultFieldKey: legacyStage.resultFieldKey,
                        };
                    });
                    profiles = derivedProfiles;
                    stages = derivedStages;
                }
            } catch {
                stages = null;
            }
        }

        if (!stages) {
            const fallbackStage: AIDetectStageConfig = {
                profileName: profiles[0]?.name ?? fallbackProfileName,
                submitFieldKeys: legacyStageConfig.submitFieldKeys,
                prompt: legacyStageConfig.prompt,
                resultFieldKey: legacyStageConfig.resultFieldKey,
            };
            stages = parseStageConfigMap(
                null,
                fallbackStage,
                profiles[0]?.name ?? fallbackProfileName,
            );
        }
        return {
            name: row.config_name,
            profiles,
            stages,
            isActive: row.is_active === 1,
            updatedAt: row.updated_at,
        };
    });
    const activeConfigName =
        configs.find((config) => config.isActive)?.name ??
        configs[0]?.name ??
        "";

    return {
        configs,
        activeConfigName,
    };
}

export function saveAIDetectConfig(
    fileName: string,
    configName: string,
    config: AIDetectConfig,
    options?: {
        setActive?: boolean;
    },
): void {
    const normalizedName = normalizeConfigName(configName);
    const shouldSetActive = options?.setActive !== false;
    const stagesJson = JSON.stringify(config.stages);
    const profilesJson = JSON.stringify(config.profiles ?? []);
    const legacyStage =
        config.stages[LEGACY_STAGE_KEY] ?? config.stages[AI_STAGE_ORDER[0]];
    const fallbackProfile =
        config.profiles?.[0]?.profile ??
        ({
            provider: "openai",
            url: "",
            model: "",
            apiKey: "",
            reasoningEffort: "high",
            retryCount: DEFAULT_AI_RETRY_COUNT,
        } as AIDetectProfile);
    const legacyProfile =
        config.profiles?.find((item) => item.name === legacyStage.profileName)
            ?.profile ?? fallbackProfile;

    const tx = db.transaction(() => {
        db.prepare(
            `INSERT INTO ai_configs (
         file_name,
         config_name,
         provider,
         ai_url,
         ai_model,
         api_key,
         vertex_project,
         vertex_location,
         submit_field_keys,
         prompt,
         result_field_key,
         reasoning_effort,
         retry_count,
         stages_json,
         profiles_json,
         is_active,
         created_at,
         updated_at
       )
       VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)
       ON CONFLICT(file_name, config_name) DO UPDATE SET
         provider = excluded.provider,
         ai_url = excluded.ai_url,
         ai_model = excluded.ai_model,
         api_key = excluded.api_key,
         vertex_project = excluded.vertex_project,
         vertex_location = excluded.vertex_location,
         submit_field_keys = excluded.submit_field_keys,
         prompt = excluded.prompt,
         result_field_key = excluded.result_field_key,
         reasoning_effort = excluded.reasoning_effort,
         retry_count = excluded.retry_count,
         stages_json = excluded.stages_json,
         profiles_json = excluded.profiles_json,
         is_active = CASE
           WHEN excluded.is_active = 1 THEN 1
           ELSE ai_configs.is_active
         END,
         updated_at = CURRENT_TIMESTAMP`,
        ).run(
            fileName,
            normalizedName,
            normalizeAIProvider(legacyProfile.provider),
            legacyProfile.url,
            legacyProfile.model,
            legacyProfile.apiKey,
            "",
            "",
            JSON.stringify(legacyStage.submitFieldKeys),
            legacyStage.prompt,
            legacyStage.resultFieldKey || null,
            legacyProfile.reasoningEffort,
            normalizeRetryCount(legacyProfile.retryCount),
            stagesJson,
            profilesJson,
            shouldSetActive ? 1 : 0,
        );

        const activeCountRow = db
            .prepare(
                "SELECT COUNT(1) AS count FROM ai_configs WHERE file_name = ? AND is_active = 1",
            )
            .get(fileName) as { count: number };
        const activeCount = Number(activeCountRow.count);
        if (shouldSetActive || activeCount === 0) {
            db.prepare(
                `UPDATE ai_configs
         SET is_active = CASE WHEN config_name = ? THEN 1 ELSE 0 END
         WHERE file_name = ?`,
            ).run(normalizedName, fileName);
        }
    });

    tx();
}

export function setAIDetectActiveConfig(
    fileName: string,
    configName: string,
): boolean {
    const normalizedName = normalizeConfigName(configName);
    const row = db
        .prepare(
            "SELECT 1 FROM ai_configs WHERE file_name = ? AND config_name = ? LIMIT 1",
        )
        .get(fileName, normalizedName);
    if (!row) {
        return false;
    }

    db.prepare(
        `UPDATE ai_configs
     SET is_active = CASE WHEN config_name = ? THEN 1 ELSE 0 END,
         updated_at = CASE WHEN config_name = ? THEN CURRENT_TIMESTAMP ELSE updated_at END
     WHERE file_name = ?`,
    ).run(normalizedName, normalizedName, fileName);

    return true;
}

function normalizeProviderEndpointConfig(
    value: unknown,
    fallback?: AIProviderEndpointConfig,
): AIProviderEndpointConfig {
    const defaultValue =
        fallback ??
        ({
            name: DEFAULT_AI_PROVIDER_NAME,
            apiType: "openai",
            apiUrl: "",
            apiKey: "",
        } as AIProviderEndpointConfig);
    if (!value || typeof value !== "object") {
        return { ...defaultValue };
    }
    const candidate = value as Partial<AIProviderEndpointConfig>;
    return {
        name:
            typeof candidate.name === "string" && candidate.name.trim().length > 0
                ? candidate.name.trim()
                : defaultValue.name,
        apiType: normalizeProviderApiType(candidate.apiType, defaultValue.apiType),
        apiUrl:
            typeof candidate.apiUrl === "string" ? candidate.apiUrl : defaultValue.apiUrl,
        apiKey:
            typeof candidate.apiKey === "string" ? candidate.apiKey : defaultValue.apiKey,
        updatedAt:
            typeof candidate.updatedAt === "string"
                ? candidate.updatedAt
                : defaultValue.updatedAt,
    };
}

function normalizeModelRouteConfig(
    value: unknown,
    fallbackProviderName: string,
    fallback?: AIModelRouteConfig,
): AIModelRouteConfig {
    const defaultValue =
        fallback ??
        ({
            name: DEFAULT_AI_ROUTE_NAME,
            model: "",
            reasoningEffort: "high",
            retryCount: DEFAULT_AI_RETRY_COUNT,
            steps: [{ providerName: fallbackProviderName }],
        } as AIModelRouteConfig);
    if (!value || typeof value !== "object") {
        return {
            ...defaultValue,
            steps: defaultValue.steps.map((step) => ({ ...step })),
        };
    }
    const candidate = value as Partial<AIModelRouteConfig>;
    return {
        name:
            typeof candidate.name === "string" && candidate.name.trim().length > 0
                ? candidate.name.trim()
                : defaultValue.name,
        model:
            typeof candidate.model === "string" && candidate.model.trim().length > 0
                ? candidate.model
                : defaultValue.model,
        reasoningEffort:
            candidate.reasoningEffort === "low" ||
            candidate.reasoningEffort === "medium" ||
            candidate.reasoningEffort === "high"
                ? candidate.reasoningEffort
                : defaultValue.reasoningEffort,
        retryCount: normalizeRetryCount(
            typeof candidate.retryCount === "number" ? candidate.retryCount : null,
        ),
        steps: normalizeRouteStepList(candidate.steps, fallbackProviderName),
        updatedAt:
            typeof candidate.updatedAt === "string"
                ? candidate.updatedAt
                : defaultValue.updatedAt,
    };
}

function normalizeFileStageConfig(
    value: unknown,
    fallbackRouteName: string,
    fallback: AIDetectStageConfig,
): FileAIStageConfig {
    if (!value || typeof value !== "object") {
        return {
            routeName: fallbackRouteName,
            submitFieldKeys: [...fallback.submitFieldKeys],
            prompt: fallback.prompt,
        };
    }
    const candidate = value as Partial<FileAIStageConfig> & {
        profileName?: unknown;
        routeName?: unknown;
    };
    const submitFieldKeys = Array.isArray(candidate.submitFieldKeys)
        ? candidate.submitFieldKeys.filter(
              (item): item is string => typeof item === "string",
          )
        : [...fallback.submitFieldKeys];
    const routeName =
        typeof candidate.routeName === "string" && candidate.routeName.trim().length > 0
            ? candidate.routeName.trim()
            : typeof candidate.profileName === "string" &&
                candidate.profileName.trim().length > 0
              ? candidate.profileName.trim()
              : fallbackRouteName;
    return {
        routeName,
        submitFieldKeys,
        prompt:
            typeof candidate.prompt === "string" && candidate.prompt.trim().length > 0
                ? candidate.prompt
                : fallback.prompt,
    };
}

function parseFileStageConfigMap(
    value: string | null | undefined,
    fallbackRouteName: string,
): Record<AIDetectStageKey, FileAIStageConfig> {
    let parsed: Record<string, unknown> = {};
    if (value) {
        try {
            const candidate = JSON.parse(value) as unknown;
            if (candidate && typeof candidate === "object") {
                parsed = candidate as Record<string, unknown>;
            }
        } catch {
            parsed = {};
        }
    }
    const stages = {} as Record<AIDetectStageKey, FileAIStageConfig>;
    AI_STAGE_ORDER.forEach((stageKey) => {
        stages[stageKey] = normalizeFileStageConfig(
            parsed[stageKey],
            fallbackRouteName,
            {
                profileName: fallbackRouteName,
                submitFieldKeys: [],
                prompt: "",
                resultFieldKey: "",
            },
        );
    });
    return stages;
}

function normalizeFileChatConfig(
    value: unknown,
    fallbackRouteName: string,
): FileAIChatConfig {
    if (!value || typeof value !== "object") {
        return {
            routeName: fallbackRouteName,
            prompt: DEFAULT_AI_CHAT_PROMPT,
            defaultSubmitFieldKeys: [],
        };
    }
    const candidate = value as {
        routeName?: unknown;
        prompt?: unknown;
        defaultSubmitFieldKeys?: unknown;
    };
    return {
        routeName:
            typeof candidate.routeName === "string" &&
            candidate.routeName.trim().length > 0
                ? candidate.routeName.trim()
                : fallbackRouteName,
        prompt:
            typeof candidate.prompt === "string" &&
            candidate.prompt.trim().length > 0
                ? candidate.prompt
                : DEFAULT_AI_CHAT_PROMPT,
        defaultSubmitFieldKeys: Array.isArray(candidate.defaultSubmitFieldKeys)
            ? candidate.defaultSubmitFieldKeys.filter(
                  (item): item is string => typeof item === "string",
              )
            : [],
    };
}

function parseFileChatConfig(
    value: string | null | undefined,
    fallbackRouteName: string,
): FileAIChatConfig {
    if (!value) {
        return normalizeFileChatConfig(null, fallbackRouteName);
    }
    try {
        return normalizeFileChatConfig(
            JSON.parse(value) as unknown,
            fallbackRouteName,
        );
    } catch {
        return normalizeFileChatConfig(null, fallbackRouteName);
    }
}

function normalizeFileEvaluationTaskConfig(
    value: unknown,
    fallbackRouteName: string,
    fallbackId: string,
    fallbackName: string,
): FileAIEvaluationTaskConfig {
    if (!value || typeof value !== "object") {
        return {
            id: fallbackId,
            name: fallbackName,
            enabled: false,
            attemptCount: DEFAULT_AI_EVALUATION_ATTEMPT_COUNT,
            maxConcurrency: DEFAULT_AI_EVALUATION_MAX_CONCURRENCY,
            answerGeneration: {
                routeName: fallbackRouteName,
                prompt: DEFAULT_AI_EVALUATION_GENERATION_PROMPT,
                questionFieldKeys: [],
            },
            answerJudgment: {
                routeName: fallbackRouteName,
                prompt: DEFAULT_AI_EVALUATION_JUDGMENT_PROMPT,
                answerFieldKeys: [],
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
    const answerGeneration =
        candidate.answerGeneration && typeof candidate.answerGeneration === "object"
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
    const generationRouteName =
        answerGeneration?.routeName ?? legacyStageCandidate?.routeName;
    const judgmentRouteName =
        answerJudgment?.routeName ?? legacyStageCandidate?.routeName;
    return {
        id:
            typeof (value as { id?: unknown }).id === "string" &&
            (value as { id: string }).id.trim().length > 0
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
        attemptCount: normalizeEvaluationAttemptCount(candidate.attemptCount),
        maxConcurrency: normalizeEvaluationMaxConcurrency(
            (value as { maxConcurrency?: unknown }).maxConcurrency,
        ),
        answerGeneration: {
            routeName:
                typeof generationRouteName === "string" &&
                generationRouteName.trim().length > 0
                    ? generationRouteName.trim()
                    : fallbackRouteName,
            prompt:
                typeof answerGeneration?.prompt === "string" &&
                answerGeneration.prompt.trim().length > 0
                    ? answerGeneration.prompt
                    : DEFAULT_AI_EVALUATION_GENERATION_PROMPT,
            questionFieldKeys: Array.isArray(answerGeneration?.questionFieldKeys)
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
            routeName:
                typeof judgmentRouteName === "string" &&
                judgmentRouteName.trim().length > 0
                    ? judgmentRouteName.trim()
                    : fallbackRouteName,
            prompt:
                typeof answerJudgment?.prompt === "string" &&
                answerJudgment.prompt.trim().length > 0
                    ? answerJudgment.prompt
                    : DEFAULT_AI_EVALUATION_JUDGMENT_PROMPT,
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

function parseFileEvaluationConfigMap(
    value: string | null | undefined,
    fallbackRouteName: string,
): FileAIEvaluationTaskConfig[] {
    if (!value) {
        return [
            normalizeFileEvaluationTaskConfig(
                null,
                fallbackRouteName,
                DEFAULT_AI_EVALUATION_TASK_ID,
                DEFAULT_AI_EVALUATION_TASK_NAME,
            ),
        ];
    }
    try {
        const parsed = JSON.parse(value) as unknown;
        if (Array.isArray(parsed)) {
            const tasks = parsed.map((item, index) =>
                normalizeFileEvaluationTaskConfig(
                    item,
                    fallbackRouteName,
                    `${DEFAULT_AI_EVALUATION_TASK_ID}-${index + 1}`,
                    index === 0
                        ? DEFAULT_AI_EVALUATION_TASK_NAME
                        : `评测配置 ${index + 1}`,
                ),
            );
            return tasks.length > 0
                ? tasks
                : [
                      normalizeFileEvaluationTaskConfig(
                          null,
                          fallbackRouteName,
                          DEFAULT_AI_EVALUATION_TASK_ID,
                          DEFAULT_AI_EVALUATION_TASK_NAME,
                      ),
                  ];
        }
        return [
            normalizeFileEvaluationTaskConfig(
                parsed,
                fallbackRouteName,
                DEFAULT_AI_EVALUATION_TASK_ID,
                DEFAULT_AI_EVALUATION_TASK_NAME,
            ),
        ];
    } catch {
        return [
            normalizeFileEvaluationTaskConfig(
                null,
                fallbackRouteName,
                DEFAULT_AI_EVALUATION_TASK_ID,
                DEFAULT_AI_EVALUATION_TASK_NAME,
            ),
        ];
    }
}

function normalizeFileCleaningToolConfig(
    value: unknown,
    fallbackRouteName: string,
    toolKey: AICleaningToolKey,
): FileAICleaningToolConfig {
    const fallbackOutputMappings = AI_CLEANING_TOOL_OUTPUT_KEYS[toolKey].map(
        (outputKey) => ({
            outputKey,
            targetFieldKey: "",
        }),
    );
    if (!value || typeof value !== "object") {
        return {
            routeName: fallbackRouteName,
            submitFieldKeys: [],
            prompt: DEFAULT_AI_CLEANING_PROMPTS[toolKey],
            autoFillEnabled: false,
            outputMappings: fallbackOutputMappings,
        };
    }

    const candidate = value as Partial<FileAICleaningToolConfig>;
    const allowedOutputKeys = new Set(AI_CLEANING_TOOL_OUTPUT_KEYS[toolKey]);
    const outputMappingMap = new Map<string, FileAICleaningOutputMapping>();
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
            outputMappingMap.set(outputKey.trim(), {
                outputKey: outputKey.trim(),
                targetFieldKey:
                    typeof targetFieldKey === "string" ? targetFieldKey : "",
            });
        });
    }

    return {
        routeName:
            typeof candidate.routeName === "string" &&
            candidate.routeName.trim().length > 0
                ? candidate.routeName.trim()
                : fallbackRouteName,
        submitFieldKeys: Array.isArray(candidate.submitFieldKeys)
            ? candidate.submitFieldKeys.filter(
                  (item): item is string => typeof item === "string",
              )
            : [],
        prompt:
            typeof candidate.prompt === "string" &&
            candidate.prompt.trim().length > 0
                ? candidate.prompt
                : DEFAULT_AI_CLEANING_PROMPTS[toolKey],
        autoFillEnabled:
            (candidate as { autoFillEnabled?: unknown }).autoFillEnabled ===
            true,
        outputMappings: AI_CLEANING_TOOL_OUTPUT_KEYS[toolKey].map(
            (outputKey) =>
                outputMappingMap.get(outputKey) ?? {
                    outputKey,
                    targetFieldKey: "",
                },
        ),
    };
}

function parseFileCleaningConfigMap(
    value: string | null | undefined,
    fallbackRouteName: string,
): FileAICleaningConfigMap {
    let parsed: Record<string, unknown> = {};
    if (value) {
        try {
            const candidate = JSON.parse(value) as unknown;
            if (candidate && typeof candidate === "object") {
                parsed = candidate as Record<string, unknown>;
            }
        } catch {
            parsed = {};
        }
    }
    const cleaning = {} as FileAICleaningConfigMap;
    AI_CLEANING_TOOL_ORDER.forEach((toolKey) => {
        cleaning[toolKey] = normalizeFileCleaningToolConfig(
            parsed[toolKey],
            fallbackRouteName,
            toolKey,
        );
    });
    return cleaning;
}

export function listAIProviderEndpoints(): AIProviderEndpointConfig[] {
    const rows = db
        .prepare(
            `SELECT name, api_type, api_url, api_key, updated_at
             FROM ai_provider_endpoints
             ORDER BY datetime(updated_at) DESC, name ASC`,
        )
        .all() as Array<{
        name: string;
        api_type: string;
        api_url: string;
        api_key: string;
        updated_at: string;
    }>;
    return rows.map((row) =>
        normalizeProviderEndpointConfig({
            name: row.name,
            apiType: row.api_type,
            apiUrl: row.api_url,
            apiKey: row.api_key,
            updatedAt: row.updated_at,
        }),
    );
}

export function saveAIProviderEndpoints(
    providers: AIProviderEndpointConfig[],
): void {
    const tx = db.transaction(() => {
        db.prepare("DELETE FROM ai_provider_endpoints").run();
        const stmt = db.prepare(
            `INSERT INTO ai_provider_endpoints (
                name, api_type, api_url, api_key, created_at, updated_at
            ) VALUES (?, ?, ?, ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)`,
        );
        providers.forEach((provider) => {
            const normalized = normalizeProviderEndpointConfig(provider);
            stmt.run(
                normalized.name,
                normalized.apiType,
                normalized.apiUrl,
                normalized.apiKey,
            );
        });
    });
    tx();
}

export function listAIModelRoutes(): AIModelRouteConfig[] {
    const rows = db
        .prepare(
            `SELECT name, model, reasoning_effort, retry_count, steps_json, updated_at
             FROM ai_model_routes
             ORDER BY datetime(updated_at) DESC, name ASC`,
        )
        .all() as Array<{
        name: string;
        model: string;
        reasoning_effort: string;
        retry_count: number;
        steps_json: string;
        updated_at: string;
    }>;

    return rows.map((row) =>
        normalizeModelRouteConfig(
            {
                name: row.name,
                model: row.model,
                reasoningEffort: normalizeReasoningEffort(row.reasoning_effort),
                retryCount: row.retry_count,
                steps: JSON.parse(row.steps_json) as unknown,
                updatedAt: row.updated_at,
            },
            DEFAULT_AI_PROVIDER_NAME,
        ),
    );
}

export function saveAIModelRoutes(routes: AIModelRouteConfig[]): void {
    const tx = db.transaction(() => {
        db.prepare("DELETE FROM ai_model_routes").run();
        const stmt = db.prepare(
            `INSERT INTO ai_model_routes (
                name, model, reasoning_effort, retry_count, steps_json, created_at, updated_at
            ) VALUES (?, ?, ?, ?, ?, CURRENT_TIMESTAMP, CURRENT_TIMESTAMP)`,
        );
        routes.forEach((route) => {
            const fallbackProviderName =
                route.steps[0]?.providerName ?? DEFAULT_AI_PROVIDER_NAME;
            const normalized = normalizeModelRouteConfig(
                route,
                fallbackProviderName,
            );
            stmt.run(
                normalized.name,
                normalized.model,
                normalized.reasoningEffort,
                normalized.retryCount,
                JSON.stringify(normalized.steps),
            );
        });
    });
    tx();
}

export function getFileAIStageConfigs(
    fileName: string,
): Record<AIDetectStageKey, FileAIStageConfig> {
    const row = db
        .prepare(
            `SELECT stages_json
             FROM file_ai_stage_configs
             WHERE file_name = ?`,
        )
        .get(fileName) as { stages_json: string } | undefined;

    return parseFileStageConfigMap(row?.stages_json, DEFAULT_AI_ROUTE_NAME);
}

export function getFileAIChatConfig(fileName: string): FileAIChatConfig {
    const row = db
        .prepare(
            `SELECT chat_json
             FROM file_ai_stage_configs
             WHERE file_name = ?`,
        )
        .get(fileName) as { chat_json: string | null } | undefined;

    return parseFileChatConfig(row?.chat_json, DEFAULT_AI_ROUTE_NAME);
}

export function getFileAIEvaluationConfig(
    fileName: string,
): FileAIEvaluationTaskConfig[] {
    const row = db
        .prepare(
            `SELECT evaluation_json
             FROM file_ai_stage_configs
             WHERE file_name = ?`,
        )
        .get(fileName) as { evaluation_json: string | null } | undefined;

    return parseFileEvaluationConfigMap(
        row?.evaluation_json,
        DEFAULT_AI_ROUTE_NAME,
    );
}

export function getFileAICleaningConfig(fileName: string): FileAICleaningConfigMap {
    const row = db
        .prepare(
            `SELECT cleaning_json
             FROM file_ai_stage_configs
             WHERE file_name = ?`,
        )
        .get(fileName) as { cleaning_json: string | null } | undefined;

    return parseFileCleaningConfigMap(
        row?.cleaning_json,
        DEFAULT_AI_ROUTE_NAME,
    );
}

export function saveFileAIStageConfigs(
    fileName: string,
    stages: Record<AIDetectStageKey, FileAIStageConfig>,
): void {
    db.prepare(
        `INSERT INTO file_ai_stage_configs (file_name, stages_json, updated_at)
         VALUES (?, ?, CURRENT_TIMESTAMP)
         ON CONFLICT(file_name) DO UPDATE SET
           stages_json = excluded.stages_json,
           updated_at = CURRENT_TIMESTAMP`,
    ).run(fileName, JSON.stringify(stages));
}

export function saveFileAIChatConfig(
    fileName: string,
    chat: FileAIChatConfig,
): void {
    const existingRow = db
        .prepare(
            `SELECT stages_json, evaluation_json, cleaning_json
             FROM file_ai_stage_configs
             WHERE file_name = ?`,
        )
        .get(fileName) as
        | {
              stages_json: string;
              evaluation_json: string | null;
              cleaning_json: string | null;
          }
        | undefined;
    db.prepare(
        `INSERT INTO file_ai_stage_configs (
             file_name,
             stages_json,
             chat_json,
             evaluation_json,
             cleaning_json,
             updated_at
         )
         VALUES (?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
         ON CONFLICT(file_name) DO UPDATE SET
           chat_json = excluded.chat_json,
           updated_at = CURRENT_TIMESTAMP`,
    ).run(
        fileName,
        existingRow?.stages_json ??
            JSON.stringify(parseFileStageConfigMap(null, DEFAULT_AI_ROUTE_NAME)),
        JSON.stringify(chat),
        existingRow?.evaluation_json ?? null,
        existingRow?.cleaning_json ?? null,
    );
}

export function saveFileAIEvaluationConfig(
    fileName: string,
    evaluation: FileAIEvaluationTaskConfig[],
): void {
    const existingRow = db
        .prepare(
            `SELECT stages_json, chat_json, cleaning_json
             FROM file_ai_stage_configs
             WHERE file_name = ?`,
        )
        .get(fileName) as
        | {
              stages_json: string;
              chat_json: string | null;
              cleaning_json: string | null;
          }
        | undefined;
    db.prepare(
        `INSERT INTO file_ai_stage_configs (
             file_name,
             stages_json,
             chat_json,
             evaluation_json,
             cleaning_json,
             updated_at
         )
         VALUES (?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
         ON CONFLICT(file_name) DO UPDATE SET
           evaluation_json = excluded.evaluation_json,
           updated_at = CURRENT_TIMESTAMP`,
    ).run(
        fileName,
        existingRow?.stages_json ??
            JSON.stringify(parseFileStageConfigMap(null, DEFAULT_AI_ROUTE_NAME)),
        existingRow?.chat_json ?? null,
        JSON.stringify(evaluation),
        existingRow?.cleaning_json ?? null,
    );
}

export function saveFileAICleaningConfig(
    fileName: string,
    cleaning: FileAICleaningConfigMap,
): void {
    const existingRow = db
        .prepare(
            `SELECT stages_json, chat_json, evaluation_json
             FROM file_ai_stage_configs
             WHERE file_name = ?`,
        )
        .get(fileName) as
        | {
              stages_json: string;
              chat_json: string | null;
              evaluation_json: string | null;
          }
        | undefined;
    db.prepare(
        `INSERT INTO file_ai_stage_configs (
             file_name,
             stages_json,
             chat_json,
             evaluation_json,
             cleaning_json,
             updated_at
         )
         VALUES (?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
         ON CONFLICT(file_name) DO UPDATE SET
           cleaning_json = excluded.cleaning_json,
           updated_at = CURRENT_TIMESTAMP`,
    ).run(
        fileName,
        existingRow?.stages_json ??
            JSON.stringify(parseFileStageConfigMap(null, DEFAULT_AI_ROUTE_NAME)),
        existingRow?.chat_json ?? null,
        existingRow?.evaluation_json ?? null,
        JSON.stringify(cleaning),
    );
}

function normalizeFileAICleaningToolResult(
    value: {
        response_text?: unknown;
        parsed_json_text?: unknown;
        updated_at?: unknown;
    },
): FileAICleaningToolResult | null {
    if (typeof value.response_text !== "string") {
        return null;
    }
    return {
        responseText: value.response_text,
        parsedJsonText:
            typeof value.parsed_json_text === "string"
                ? value.parsed_json_text
                : undefined,
        updatedAt:
            typeof value.updated_at === "string" ? value.updated_at : undefined,
    };
}

function normalizeFileAIEvaluationAttemptResult(
    value: {
        attempt_index?: unknown;
        generation_response_text?: unknown;
        generation_parsed_json_text?: unknown;
        judgment_response_text?: unknown;
        judgment_parsed_json_text?: unknown;
        final_verdict?: unknown;
        updated_at?: unknown;
    },
): FileAIEvaluationAttemptResult | null {
    if (
        typeof value.attempt_index !== "number" ||
        typeof value.generation_response_text !== "string" ||
        typeof value.judgment_response_text !== "string" ||
        typeof value.final_verdict !== "string"
    ) {
        return null;
    }
    return {
        attemptIndex: value.attempt_index,
        generationResponseText: value.generation_response_text,
        generationParsedJsonText:
            typeof value.generation_parsed_json_text === "string"
                ? value.generation_parsed_json_text
                : undefined,
        judgmentResponseText: value.judgment_response_text,
        judgmentParsedJsonText:
            typeof value.judgment_parsed_json_text === "string"
                ? value.judgment_parsed_json_text
                : undefined,
        finalVerdict: value.final_verdict,
        updatedAt:
            typeof value.updated_at === "string" ? value.updated_at : undefined,
    };
}

export function listFileAICleaningResults(
    fileId: string,
): Partial<Record<AICleaningToolKey, Record<string, FileAICleaningToolResult>>> {
    const result: Partial<
        Record<AICleaningToolKey, Record<string, FileAICleaningToolResult>>
    > = {};
    AI_CLEANING_TOOL_ORDER.forEach((toolKey) => {
        const rows = db
            .prepare(
                `SELECT row_id, response_text, parsed_json_text, updated_at
                 FROM ${getAICleaningResultTableName(toolKey)}
                 WHERE file_id = ?
                 ORDER BY datetime(updated_at) DESC, row_id ASC`,
            )
            .all(fileId) as Array<{
            row_id: string;
            response_text: string;
            parsed_json_text: string | null;
            updated_at: string;
        }>;
        if (rows.length === 0) {
            return;
        }
        const rowMap: Record<string, FileAICleaningToolResult> = {};
        rows.forEach((row) => {
            const normalized = normalizeFileAICleaningToolResult(row);
            if (normalized) {
                rowMap[row.row_id] = normalized;
            }
        });
        if (Object.keys(rowMap).length > 0) {
            result[toolKey] = rowMap;
        }
    });
    return result;
}

export function listFileAIEvaluationResults(
    fileId: string,
): Record<string, Record<string, FileAIEvaluationAttemptResult[]>> {
    const rows = db
        .prepare(
            `SELECT row_id, task_id, attempt_index, generation_response_text, generation_parsed_json_text, judgment_response_text, judgment_parsed_json_text, final_verdict, updated_at
             FROM ${getAIEvaluationResultTableName()}
             WHERE file_id = ?
             ORDER BY task_id ASC, row_id ASC, attempt_index ASC`,
        )
        .all(fileId) as Array<{
        row_id: string;
        task_id: string;
        attempt_index: number;
        generation_response_text: string;
        generation_parsed_json_text: string | null;
        judgment_response_text: string;
        judgment_parsed_json_text: string | null;
        final_verdict: string;
        updated_at: string;
    }>;
    const result: Record<string, Record<string, FileAIEvaluationAttemptResult[]>> =
        {};
    rows.forEach((row) => {
        const normalized = normalizeFileAIEvaluationAttemptResult(row);
        if (!normalized) {
            return;
        }
        const taskMap = (result[row.task_id] ??= {});
        const attempts = (taskMap[row.row_id] ??= []);
        attempts.push(normalized);
    });
    return result;
}

export function saveFileAICleaningToolResult(
    fileId: string,
    fileName: string,
    rowId: string,
    toolKey: AICleaningToolKey,
    responseText: string,
    parsedJsonText?: string,
): void {
    db.prepare(
        `INSERT INTO ${getAICleaningResultTableName(toolKey)} (
             file_id,
             file_name,
             row_id,
             response_text,
             parsed_json_text,
             updated_at
         )
         VALUES (?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
         ON CONFLICT(file_id, row_id) DO UPDATE SET
           file_name = excluded.file_name,
           response_text = excluded.response_text,
           parsed_json_text = excluded.parsed_json_text,
           updated_at = CURRENT_TIMESTAMP`,
    ).run(fileId, fileName, rowId, responseText, parsedJsonText ?? null);
}

export function saveFileAIEvaluationAttemptResult(
    fileId: string,
    fileName: string,
    rowId: string,
    taskId: string,
    attemptIndex: number,
    generationResponseText: string,
    generationParsedJsonText: string | undefined,
    judgmentResponseText: string,
    judgmentParsedJsonText: string | undefined,
    finalVerdict: string,
): void {
    db.prepare(
        `INSERT INTO ${getAIEvaluationResultTableName()} (
             file_id,
             file_name,
             row_id,
             task_id,
             attempt_index,
             generation_response_text,
             generation_parsed_json_text,
             judgment_response_text,
             judgment_parsed_json_text,
             final_verdict,
             updated_at
         )
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
         ON CONFLICT(file_id, row_id, task_id, attempt_index) DO UPDATE SET
           file_name = excluded.file_name,
           generation_response_text = excluded.generation_response_text,
           generation_parsed_json_text = excluded.generation_parsed_json_text,
           judgment_response_text = excluded.judgment_response_text,
           judgment_parsed_json_text = excluded.judgment_parsed_json_text,
           final_verdict = excluded.final_verdict,
           updated_at = CURRENT_TIMESTAMP`,
    ).run(
        fileId,
        fileName,
        rowId,
        taskId,
        attemptIndex,
        generationResponseText,
        generationParsedJsonText ?? null,
        judgmentResponseText,
        judgmentParsedJsonText ?? null,
        finalVerdict,
    );
}

export function deleteFileAICleaningResults(fileId: string): void {
    AI_CLEANING_TOOL_ORDER.forEach((toolKey) => {
        db.prepare(
            `DELETE FROM ${getAICleaningResultTableName(toolKey)} WHERE file_id = ?`,
        ).run(fileId);
    });
}

export function deleteFileAIEvaluationResults(fileId: string): void {
    db.prepare(
        `DELETE FROM ${getAIEvaluationResultTableName()} WHERE file_id = ?`,
    ).run(fileId);
}

export function findAIProviderEndpointByName(
    name: string,
): AIProviderEndpointConfig | null {
    const row = db
        .prepare(
            `SELECT name, api_type, api_url, api_key, updated_at
             FROM ai_provider_endpoints
             WHERE name = ?`,
        )
        .get(name) as
        | {
              name: string;
              api_type: string;
              api_url: string;
              api_key: string;
              updated_at: string;
          }
        | undefined;
    if (!row) {
        return null;
    }
    return normalizeProviderEndpointConfig({
        name: row.name,
        apiType: row.api_type,
        apiUrl: row.api_url,
        apiKey: row.api_key,
        updatedAt: row.updated_at,
    });
}

export function findAIModelRouteByName(name: string): AIModelRouteConfig | null {
    const row = db
        .prepare(
            `SELECT name, model, reasoning_effort, retry_count, steps_json, updated_at
             FROM ai_model_routes
             WHERE name = ?`,
        )
        .get(name) as
        | {
              name: string;
              model: string;
              reasoning_effort: string;
              retry_count: number;
              steps_json: string;
              updated_at: string;
          }
        | undefined;
    if (!row) {
        return null;
    }
    return normalizeModelRouteConfig(
        {
            name: row.name,
            model: row.model,
            reasoningEffort: normalizeReasoningEffort(row.reasoning_effort),
            retryCount: row.retry_count,
            steps: JSON.parse(row.steps_json) as unknown,
            updatedAt: row.updated_at,
        },
        DEFAULT_AI_PROVIDER_NAME,
    );
}

function legacyProviderToApiType(provider: AIProvider): AIProviderApiType {
    if (provider === "gemini" || provider === "modelrouter-gemini") {
        return "gemini";
    }
    return "openai";
}

function buildLegacyMigrationName(fileName: string, itemName: string): string {
    const trimmedItem = itemName.trim() || "default";
    return `${fileName} / ${trimmedItem}`;
}

function migrateLegacyAIConfigToRoutingTables(): void {
    const providerCountRow = db
        .prepare("SELECT COUNT(1) AS count FROM ai_provider_endpoints")
        .get() as { count: number };
    const routeCountRow = db
        .prepare("SELECT COUNT(1) AS count FROM ai_model_routes")
        .get() as { count: number };
    const stageCountRow = db
        .prepare("SELECT COUNT(1) AS count FROM file_ai_stage_configs")
        .get() as { count: number };
    if (
        Number(providerCountRow.count) > 0 ||
        Number(routeCountRow.count) > 0 ||
        Number(stageCountRow.count) > 0
    ) {
        return;
    }

    const fileRows = db
        .prepare("SELECT DISTINCT file_name FROM ai_configs ORDER BY file_name ASC")
        .all() as Array<{ file_name: string }>;
    if (fileRows.length === 0) {
        return;
    }

    createSyncDatabaseBackup("ai-routing-migration");

    const providers = new Map<string, AIProviderEndpointConfig>();
    const routes = new Map<string, AIModelRouteConfig>();
    const stagesByFile = new Map<
        string,
        Record<AIDetectStageKey, FileAIStageConfig>
    >();

    fileRows.forEach(({ file_name: fileName }) => {
        const { configs, activeConfigName } = listAIDetectConfigs(fileName);
        const activeConfig =
            configs.find((item) => item.name === activeConfigName) ?? configs[0];
        if (!activeConfig) {
            return;
        }

        const routeNameByProfileName = new Map<string, string>();
        activeConfig.profiles.forEach((profileItem) => {
            const providerName = buildLegacyMigrationName(
                fileName,
                profileItem.name,
            );
            providers.set(providerName, {
                name: providerName,
                apiType: legacyProviderToApiType(profileItem.profile.provider),
                apiUrl: profileItem.profile.url,
                apiKey: profileItem.profile.apiKey,
            });
            routes.set(providerName, {
                name: providerName,
                model: profileItem.profile.model,
                reasoningEffort: profileItem.profile.reasoningEffort,
                retryCount: profileItem.profile.retryCount,
                steps: [{ providerName }],
            });
            routeNameByProfileName.set(profileItem.name, providerName);
        });

        const stageMap = {} as Record<AIDetectStageKey, FileAIStageConfig>;
        AI_STAGE_ORDER.forEach((stageKey) => {
            const stage = activeConfig.stages[stageKey];
            const routeName =
                routeNameByProfileName.get(stage.profileName) ??
                Array.from(routeNameByProfileName.values())[0] ??
                DEFAULT_AI_ROUTE_NAME;
            stageMap[stageKey] = {
                routeName,
                submitFieldKeys: [...stage.submitFieldKeys],
                prompt: stage.prompt,
            };
        });
        stagesByFile.set(fileName, stageMap);
    });

    saveAIProviderEndpoints(Array.from(providers.values()));
    saveAIModelRoutes(Array.from(routes.values()));
    stagesByFile.forEach((stages, fileName) => {
        saveFileAIStageConfigs(fileName, stages);
    });
}

function migrateLegacyProjectName(): void {
    const legacyRows = db
        .prepare("SELECT file_id FROM file_states WHERE file_name = ?")
        .all(LEGACY_PROJECT_NAME) as Array<{ file_id: string }>;

    if (legacyRows.length !== 1) {
        return;
    }

    if (isProjectNameInUse(MIGRATED_LEGACY_PROJECT_NAME)) {
        return;
    }

    renameProjectReferences(
        legacyRows[0].file_id,
        LEGACY_PROJECT_NAME,
        MIGRATED_LEGACY_PROJECT_NAME,
    );
}

migrateLegacyProjectName();

export default db;
