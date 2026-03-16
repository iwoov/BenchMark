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

export interface FileViewState extends ParsedFile {
    selectedDisplayColumnKeys: string[];
    selectedEditableColumnKeys: string[];
    selectedFilterColumnKeys: string[];
    columnFilterValues: Record<string, string>;
}

export type AIDetectStageKey =
    | "precheck"
    | "context_audit"
    | "independent_solving"
    | "final_verdict";

export type AIDetectRunKey = AIDetectStageKey | "all";

export interface AIDetectProfile {
    provider: "openai" | "gemini" | "modelrouter-openai" | "modelrouter-gemini";
    url: string;
    model: string;
    modelProvider?: string;
    modelName?: string;
    apiKey: string;
    reasoningEffort: "low" | "medium" | "high";
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
}

export type AIDetectStageConfigMap = Record<
    AIDetectStageKey,
    AIDetectStageConfig
>;

export interface AIDetectConfig {
    profiles: NamedAIDetectProfile[];
    stages: AIDetectStageConfigMap;
}

export interface NamedAIDetectConfig {
    name: string;
    config: AIDetectConfig;
}
