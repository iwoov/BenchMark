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
}

export interface AIDetectConfig {
  url: string;
  model: string;
  apiKey: string;
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
