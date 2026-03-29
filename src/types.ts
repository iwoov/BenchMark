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
    values: Record<string, ParsedCell>;
    aiResults?: Partial<Record<AIDetectStageKey, string>>;
}

export interface ParsedFile {
    fileId: string;
    fileName: string;
    sourceFileName?: string;
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

export interface FileViewState extends ParsedFile {
    selectedDisplayColumnKeys: string[];
    selectedEditableColumnKeys: string[];
    filterConditions: FilterCondition[];
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

export interface AIChatConfig {
    routeName: string;
    prompt: string;
    defaultSubmitFieldKeys: string[];
}

export type AIDetectStageConfigMap = Record<
    AIDetectStageKey,
    AIDetectStageConfig
>;

export interface AIDetectConfig {
    providers: AIProviderEndpoint[];
    routes: AIModelRoute[];
    stages: AIDetectStageConfigMap;
    chat: AIChatConfig;
}

export interface NamedAIDetectConfig {
    name: string;
    config: AIDetectConfig;
}
