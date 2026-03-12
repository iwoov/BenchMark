export interface ColumnPrefsConfig {
  fieldSignature: string;
  displayKeys: string[];
  editableKeys: string[];
  filterKeys?: string[];
}

export type MainSection = "list" | "detail" | "settings";
export type SettingsSection = "fields" | "ai";

export interface RouteState {
  section: MainSection;
  settingsSection: SettingsSection;
  rowId?: string | null;
}

export interface AIDetectFieldPayload {
  title: string;
  type: "text" | "image";
  value: string;
  imageUrl?: string;
  imageUrls?: string[];
}

export interface AIDetectStreamResult {
  answerText: string;
  thinkingText: string;
}

export type AIBatchTaskStatus = "idle" | "running" | "completed";

export interface AIBatchTaskState {
  status: AIBatchTaskStatus;
  fileId: string | null;
  fileName: string;
  total: number;
  completed: number;
  success: number;
  failed: number;
  message: string;
}
