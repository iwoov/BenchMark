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
    aiResults?: Record<string, string>;
    cleaningResults?: Record<
        string,
        {
            responseText: string;
            parsedJsonText?: string;
            updatedAt?: string;
        }
    >;
    evaluationResults?: Record<
        string,
        Array<{
            attemptIndex: number;
            generationResponseText: string;
            generationParsedJsonText?: string;
            judgmentResponseText: string;
            judgmentParsedJsonText?: string;
            finalVerdict: string;
            updatedAt?: string;
            generationLatencyMs?: number;
            judgmentLatencyMs?: number;
            generationInputTokens?: number;
            generationOutputTokens?: number;
            judgmentInputTokens?: number;
            judgmentOutputTokens?: number;
            generationFinishReason?: string;
            judgmentFinishReason?: string;
        }>
    >;
}

export interface ParsedWorkbook {
    fileId: string;
    fileName: string;
    sourceFileName?: string;
    columns: ParsedColumn[];
    rows: ParsedRow[];
    level1Options: string[];
    level2Options: string[];
}
