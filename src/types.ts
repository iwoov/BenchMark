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
}

export interface ParsedFile {
  fileId: string;
  fileName: string;
  columns: ParsedColumn[];
  rows: ParsedRow[];
  level1Options: string[];
  level2Options: string[];
}

export interface FileViewState extends ParsedFile {
  selectedDisplayColumnKeys: string[];
  selectedEditableColumnKeys: string[];
  level1Filter: string;
  level2Filter: string;
  timeFilter: string;
}

export interface AIDetectConfig {
  provider: "openai" | "vertex";
  url: string;
  model: string;
  apiKey: string;
  vertexProject: string;
  vertexLocation: string;
  submitFieldKeys: string[];
  prompt: string;
  resultFieldKey: string;
  reasoningEffort: "low" | "medium" | "high";
  retryCount: number;
}

export interface NamedAIDetectConfig {
  name: string;
  config: AIDetectConfig;
}
