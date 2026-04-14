export type CellType = "text" | "image";

export interface ParsedCell {
    type: CellType;
    value?: string;
    src?: string;
    srcList?: string[];
}

export interface ParsedColumn {
    key: string;
    title: string;
    editable: boolean;
    required: boolean;
}

export interface ParsedRow {
    rowId: string;
    enabled: boolean;
    values: Record<string, ParsedCell>;
    aiResults?: Partial<Record<AIDetectStageKey, string>>;
    cleaningResults?: Partial<Record<AICleaningToolKey, AICleaningToolResult>>;
    evaluationResults?: Record<string, AIEvaluationAttemptResult[]>;
}

export interface ParsedFile {
    fileId: string;
    fileName: string;
    sourceFileName?: string;
    updatedAt?: string;
    columns: ParsedColumn[];
    rows: ParsedRow[];
    level1Options: string[];
    level2Options: string[];
}

export interface FilterCondition {
    id: string;
    columnKey: string;
    value: string;
}

export type StatisticsChartType = "bar" | "pie" | "line" | "table";

export interface StatisticsConfig {
    selectedFieldKeys: string[];
    chartTypeByField: Record<string, StatisticsChartType>;
}

export interface FileViewState extends ParsedFile {
    selectedDisplayColumnKeys: string[];
    selectedEditableColumnKeys: string[];
    filterConditions: FilterCondition[];
    statisticsConfig: StatisticsConfig;
    /** Row count from server summary; always available even before rows are loaded. */
    rowCount?: number;
    /** Whether full row data has been fetched from the server. */
    detailLoaded?: boolean;
}

export type AIDetectStageKey =
    | "precheck"
    | "context_audit"
    | "independent_solving"
    | "final_verdict";

export type AIDetectRunKey = AIDetectStageKey | "all";

export type AIProviderApiType = "openai" | "gemini" | "anthropic";

export interface AIProviderEndpoint {
    name: string;
    apiType: AIProviderApiType;
    apiUrl: string;
    apiKey: string;
}

export interface AIModelRouteStep {
    providerName: string;
}

export interface AIModelRoute {
    name: string;
    model: string;
    reasoningEffort: "low" | "medium" | "high";
    retryCount: number;
    steps: AIModelRouteStep[];
}

export interface AIDetectStageConfig {
    routeName: string;
    submitFieldKeys: string[];
    prompt: string;
}

export interface AIEvaluationAnswerGenerationConfig {
    routeName: string;
    prompt: string;
    questionFieldKeys: string[];
}

export interface AIEvaluationAnswerJudgmentConfig {
    routeName: string;
    prompt: string;
    answerFieldKeys: string[];
}

export interface AIEvaluationTaskConfig {
    id: string;
    name: string;
    enabled: boolean;
    attemptCount: number;
    maxConcurrency: number;
    answerGeneration: AIEvaluationAnswerGenerationConfig;
    answerJudgment: AIEvaluationAnswerJudgmentConfig;
}

export interface AIEvaluationAttemptResult {
    attemptIndex: number;
    generationResponseText: string;
    generationParsedJsonText?: string;
    judgmentResponseText: string;
    judgmentParsedJsonText?: string;
    finalVerdict: string;
    updatedAt?: string;
}

export interface AIChatConfig {
    routeName: string;
    prompt: string;
    defaultSubmitFieldKeys: string[];
}

export type AICleaningToolKey =
    | "generate_level3_tags"
    | "biochem_level1_refine"
    | "question_formatting";

export interface AICleaningOutputMapping {
    outputKey: string;
    targetFieldKey: string;
}

export interface AICleaningToolResult {
    responseText: string;
    parsedJsonText?: string;
    updatedAt?: string;
}

export interface AICleaningToolConfig {
    routeName: string;
    submitFieldKeys: string[];
    prompt: string;
    autoFillEnabled: boolean;
    outputMappings: AICleaningOutputMapping[];
}

export type AICleaningConfigMap = Record<
    AICleaningToolKey,
    AICleaningToolConfig
>;

export type AIBatchToolKey = AIDetectRunKey | AICleaningToolKey;

export type AIDetectStageConfigMap = Record<
    AIDetectStageKey,
    AIDetectStageConfig
>;

export interface AIDetectConfig {
    providers: AIProviderEndpoint[];
    routes: AIModelRoute[];
    stages: AIDetectStageConfigMap;
    evaluationTasks: AIEvaluationTaskConfig[];
    chat: AIChatConfig;
    cleaning: AICleaningConfigMap;
}

export interface NamedAIDetectConfig {
    name: string;
    config: AIDetectConfig;
}

export type AIStreamPhase =
    | "requesting"
    | "thinking"
    | "outputting"
    | "completed";
