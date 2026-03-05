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

export interface ParsedWorkbook {
  fileId: string;
  fileName: string;
  columns: ParsedColumn[];
  rows: ParsedRow[];
  level1Options: string[];
  level2Options: string[];
}
