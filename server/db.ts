import Database from "better-sqlite3";
import path from "node:path";
import fs from "node:fs";
import { fileURLToPath } from "node:url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const dataDir = path.resolve(__dirname, "..", "data");

// Ensure data directory exists
if (!fs.existsSync(dataDir)) {
  fs.mkdirSync(dataDir, { recursive: true });
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
    editable_keys TEXT
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

db.exec(`
  CREATE TABLE IF NOT EXISTS ai_configs (
    file_name TEXT PRIMARY KEY,
    ai_url TEXT NOT NULL,
    ai_model TEXT NOT NULL,
    api_key TEXT NOT NULL,
    submit_field_keys TEXT NOT NULL,
    prompt TEXT NOT NULL,
    result_field_key TEXT
  );
`);

const tableColumns = db
  .prepare("PRAGMA table_info(column_prefs)")
  .all() as Array<{ name: string }>;
const hasFieldSignatureColumn = tableColumns.some(
  (column) => column.name === "field_signature",
);
const hasEditableKeysColumn = tableColumns.some(
  (column) => column.name === "editable_keys",
);

if (!hasFieldSignatureColumn) {
  db.exec("ALTER TABLE column_prefs ADD COLUMN field_signature TEXT");
}
if (!hasEditableKeysColumn) {
  db.exec("ALTER TABLE column_prefs ADD COLUMN editable_keys TEXT");
}

export interface ColumnPrefsConfig {
  fieldSignature: string;
  displayKeys: string[];
  editableKeys: string[];
}

export interface PersistedFileState {
  fileId: string;
  fileName: string;
  state: unknown;
  updatedAt: string;
}

export interface AIDetectConfig {
  url: string;
  model: string;
  apiKey: string;
  submitFieldKeys: string[];
  prompt: string;
  resultFieldKey: string;
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
    return parsed.filter((item): item is string => typeof item === "string");
  } catch {
    return [];
  }
}

/**
 * Get saved column preferences for a given file name.
 * Returns null if not found.
 */
export function getColumnPrefs(fileName: string): ColumnPrefsConfig | null {
  const row = db
    .prepare(
      "SELECT selected_keys, field_signature, editable_keys FROM column_prefs WHERE file_name = ?",
    )
    .get(fileName) as
    | {
        selected_keys: string;
        field_signature: string | null;
        editable_keys: string | null;
      }
    | undefined;

  if (!row) {
    return null;
  }

  return {
    fieldSignature: row.field_signature ?? "",
    displayKeys: parseJsonStringArray(row.selected_keys),
    editableKeys: parseJsonStringArray(row.editable_keys),
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
    `INSERT INTO column_prefs (file_name, selected_keys, field_signature, editable_keys)
     VALUES (?, ?, ?, ?)
     ON CONFLICT(file_name) DO UPDATE SET
       selected_keys = excluded.selected_keys,
       field_signature = excluded.field_signature,
       editable_keys = excluded.editable_keys`,
  ).run(
    fileName,
    JSON.stringify(config.displayKeys),
    config.fieldSignature,
    JSON.stringify(config.editableKeys),
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

export function deleteFileState(fileId: string): void {
  db.prepare("DELETE FROM file_states WHERE file_id = ?").run(fileId);
}

export function getAIDetectConfig(fileName: string): AIDetectConfig | null {
  const row = db
    .prepare(
      `SELECT ai_url, ai_model, api_key, submit_field_keys, prompt, result_field_key
       FROM ai_configs
       WHERE file_name = ?`,
    )
    .get(fileName) as
    | {
        ai_url: string;
        ai_model: string;
        api_key: string;
        submit_field_keys: string;
        prompt: string;
        result_field_key: string | null;
      }
    | undefined;

  if (!row) {
    return null;
  }

  return {
    url: row.ai_url,
    model: row.ai_model,
    apiKey: row.api_key,
    submitFieldKeys: parseJsonStringArray(row.submit_field_keys),
    prompt: row.prompt,
    resultFieldKey: row.result_field_key ?? "",
  };
}

export function saveAIDetectConfig(
  fileName: string,
  config: AIDetectConfig,
): void {
  db.prepare(
    `INSERT INTO ai_configs (file_name, ai_url, ai_model, api_key, submit_field_keys, prompt, result_field_key)
     VALUES (?, ?, ?, ?, ?, ?, ?)
     ON CONFLICT(file_name) DO UPDATE SET
       ai_url = excluded.ai_url,
       ai_model = excluded.ai_model,
       api_key = excluded.api_key,
       submit_field_keys = excluded.submit_field_keys,
       prompt = excluded.prompt,
       result_field_key = excluded.result_field_key`,
  ).run(
    fileName,
    config.url,
    config.model,
    config.apiKey,
    JSON.stringify(config.submitFieldKeys),
    config.prompt,
    config.resultFieldKey || null,
  );
}

export default db;
