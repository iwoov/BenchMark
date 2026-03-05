import * as ExcelJS from "exceljs";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import type {
  ParsedCell,
  ParsedColumn,
  ParsedRow,
  ParsedWorkbook,
} from "./types.js";

type ExcelJSImportLike = { Workbook: new () => ExcelJS.Workbook };
const ExcelJSRuntime: ExcelJSImportLike =
  (ExcelJS as unknown as { default?: ExcelJSImportLike }).default ??
  (ExcelJS as unknown as ExcelJSImportLike);

const LEVEL1_ALIASES = ["level1"];
const LEVEL2_ALIASES = ["level2"];
const IMAGE_EXTENSIONS = ["png", "jpg", "jpeg", "webp"];
const IMAGE_SEARCH_MAX_DEPTH = 6;
const IMAGE_SEARCH_SKIP_DIRS = new Set([
  ".git",
  "node_modules",
  "Library",
  ".Trash",
]);

const imageRootCache = new Map<string, string | null>();
const hyperlinkPathCache = new Map<string, string | null>();

interface ColumnRuntimeMeta extends ParsedColumn {
  colNumber: number;
}

function normalizeHeaderTitle(value: string): string {
  return value.replace(/\s+/g, "").toLowerCase();
}

function matchesHeader(title: string, aliases: string[]): boolean {
  const normalizedTitle = normalizeHeaderTitle(title);
  return aliases.some(
    (alias) => normalizeHeaderTitle(alias) === normalizedTitle,
  );
}

function isRequiredFilterHeader(title: string): boolean {
  return (
    matchesHeader(title, LEVEL1_ALIASES) || matchesHeader(title, LEVEL2_ALIASES)
  );
}

function normalizeCellText(value: ExcelJS.CellValue): string {
  if (value === null || value === undefined) {
    return "";
  }

  if (typeof value === "string") {
    return value.trim();
  }

  if (typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }

  if (value instanceof Date) {
    return value.toISOString();
  }

  if (typeof value === "object") {
    if ("richText" in value && Array.isArray(value.richText)) {
      return value.richText
        .map((item) => item.text)
        .join("")
        .trim();
    }

    if ("hyperlink" in value) {
      const text =
        "text" in value && typeof value.text === "string" ? value.text : "";
      if (text) {
        return text.trim();
      }
      return typeof value.hyperlink === "string" ? value.hyperlink.trim() : "";
    }

    if ("formula" in value) {
      const formulaResult = "result" in value ? value.result : "";
      return normalizeCellText(formulaResult as ExcelJS.CellValue);
    }

    if ("error" in value) {
      return "";
    }

    if ("text" in value && typeof value.text === "string") {
      return value.text.trim();
    }

    const str = String(value);
    if (str === "[object Object]") {
      return "";
    }
    return str.trim();
  }

  return String(value).trim();
}

function detectHeaderRow(worksheet: ExcelJS.Worksheet): number {
  const probeLimit = Math.max(1, Math.min(worksheet.rowCount, 25));
  let bestIndex = 1;
  let bestScore = -1;

  for (let rowIndex = 1; rowIndex <= probeLimit; rowIndex += 1) {
    const row = worksheet.getRow(rowIndex);
    const rowValues = Array.isArray(row.values) ? row.values.slice(1) : [];
    const texts = rowValues
      .map((item: ExcelJS.CellValue | undefined) =>
        normalizeCellText((item ?? null) as ExcelJS.CellValue),
      )
      .filter((value: string) => Boolean(value));

    const hasFilterHeaders = texts.some((value) =>
      isRequiredFilterHeader(value),
    )
      ? 1
      : 0;
    const score = texts.length + hasFilterHeaders * 100;

    if (score > bestScore) {
      bestScore = score;
      bestIndex = rowIndex;
    }
  }

  return bestIndex;
}

function buildColumns(headerRow: ExcelJS.Row): ColumnRuntimeMeta[] {
  const columns: ColumnRuntimeMeta[] = [];

  headerRow.eachCell({ includeEmpty: false }, (cell, colNumber) => {
    const title = normalizeCellText(cell.value);
    if (!title) {
      return;
    }

    columns.push({
      key: `col_${colNumber}`,
      title,
      editable: false,
      required: isRequiredFilterHeader(title),
      colNumber,
    });
  });

  return columns;
}

function getImageExtFromPathLike(pathLike: string): string | null {
  const purePath = pathLike.split(/[?#]/)[0];
  const ext = path.extname(purePath).replace(".", "").toLowerCase();
  return IMAGE_EXTENSIONS.includes(ext) ? ext : null;
}

function isPathExistingFile(filePath: string): boolean {
  try {
    return fs.existsSync(filePath) && fs.statSync(filePath).isFile();
  } catch {
    return false;
  }
}

function isPathExistingDir(dirPath: string): boolean {
  try {
    return fs.existsSync(dirPath) && fs.statSync(dirPath).isDirectory();
  } catch {
    return false;
  }
}

function safeDecodeURIComponent(value: string): string {
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
}

function normalizeRelativePathSegments(pathLike: string): string[] {
  return pathLike
    .replaceAll("\\", "/")
    .split("/")
    .map((segment) => safeDecodeURIComponent(segment.trim()))
    .filter((segment) => segment.length > 0);
}

function getImageSearchRoots(): string[] {
  const roots: string[] = [];
  const envRoot = process.env.BENCHMARK_IMAGE_ROOT?.trim();
  if (envRoot) {
    roots.push(path.resolve(envRoot));
  }

  const cwd = process.cwd();
  roots.push(cwd);

  const home = process.env.HOME?.trim();
  if (home) {
    roots.push(path.resolve(home, "Downloads"));
    roots.push(path.resolve(home, "Desktop"));
    roots.push(path.resolve(home, "Documents"));
    roots.push(path.resolve(home));
  }

  return Array.from(new Set(roots)).filter((item) => isPathExistingDir(item));
}

function findDirectoryByName(
  rootPath: string,
  targetName: string,
  depth: number,
): string | null {
  if (depth < 0 || !isPathExistingDir(rootPath)) {
    return null;
  }

  let entries: fs.Dirent[] = [];
  try {
    entries = fs.readdirSync(rootPath, { withFileTypes: true });
  } catch {
    return null;
  }

  for (const entry of entries) {
    if (!entry.isDirectory()) {
      continue;
    }
    if (entry.name === targetName) {
      return path.join(rootPath, entry.name);
    }
  }

  if (depth === 0) {
    return null;
  }

  for (const entry of entries) {
    if (!entry.isDirectory()) {
      continue;
    }
    if (IMAGE_SEARCH_SKIP_DIRS.has(entry.name)) {
      continue;
    }

    const nextPath = path.join(rootPath, entry.name);
    const found = findDirectoryByName(nextPath, targetName, depth - 1);
    if (found) {
      return found;
    }
  }

  return null;
}

function locateImageRootDirectory(rootDirName: string): string | null {
  if (imageRootCache.has(rootDirName)) {
    return imageRootCache.get(rootDirName) ?? null;
  }

  for (const searchRoot of getImageSearchRoots()) {
    const found = findDirectoryByName(
      searchRoot,
      rootDirName,
      IMAGE_SEARCH_MAX_DEPTH,
    );
    if (found) {
      imageRootCache.set(rootDirName, found);
      return found;
    }
  }

  imageRootCache.set(rootDirName, null);
  return null;
}

function normalizeHyperlinkToAbsolutePath(
  hyperlink: string,
  fileName: string,
): string | null {
  const trimmed = hyperlink.trim();
  if (!trimmed) {
    return null;
  }

  if (hyperlinkPathCache.has(trimmed)) {
    return hyperlinkPathCache.get(trimmed) ?? null;
  }

  if (/^https?:\/\//i.test(trimmed)) {
    hyperlinkPathCache.set(trimmed, trimmed);
    return trimmed;
  }

  if (/^file:\/\//i.test(trimmed)) {
    try {
      const resolved = fileURLToPath(new URL(trimmed));
      hyperlinkPathCache.set(trimmed, resolved);
      return resolved;
    } catch {
      hyperlinkPathCache.set(trimmed, null);
      return null;
    }
  }

  if (path.isAbsolute(trimmed) || /^[a-zA-Z]:[\\/]/.test(trimmed)) {
    hyperlinkPathCache.set(trimmed, trimmed);
    return trimmed;
  }

  const segments = normalizeRelativePathSegments(trimmed);
  if (segments.length === 0) {
    hyperlinkPathCache.set(trimmed, null);
    return null;
  }

  const normalizedRelativePath = path.join(...segments);
  for (const root of getImageSearchRoots()) {
    const candidate = path.resolve(root, normalizedRelativePath);
    if (isPathExistingFile(candidate)) {
      hyperlinkPathCache.set(trimmed, candidate);
      return candidate;
    }
  }

  const expectedRootName = `${fileName.replace(/\.[^.]+$/, "")}-FILE`;
  const rootNameFromLink =
    segments.find((segment) => segment.endsWith("-FILE")) ?? expectedRootName;

  const matchedRootDir = locateImageRootDirectory(rootNameFromLink);
  if (matchedRootDir) {
    const rootIndex = segments.findIndex((item) => item === rootNameFromLink);
    const tail = rootIndex >= 0 ? segments.slice(rootIndex + 1) : segments;
    const candidate = path.join(matchedRootDir, ...tail);
    if (isPathExistingFile(candidate)) {
      hyperlinkPathCache.set(trimmed, candidate);
      return candidate;
    }
  }

  hyperlinkPathCache.set(trimmed, null);
  return null;
}

function toImageSrcFromHyperlink(
  hyperlink: string,
  fileName: string,
): string | null {
  const trimmed = hyperlink.trim();
  if (!trimmed) {
    return null;
  }

  if (/^https?:\/\//i.test(trimmed)) {
    try {
      const url = new URL(trimmed);
      const ext = getImageExtFromPathLike(url.pathname);
      if (!ext) {
        return null;
      }
      return trimmed;
    } catch {
      return null;
    }
  }

  const absolutePath = normalizeHyperlinkToAbsolutePath(trimmed, fileName);
  if (!absolutePath) {
    return null;
  }

  const ext = getImageExtFromPathLike(absolutePath);
  if (!ext) {
    return null;
  }

  if (!isPathExistingFile(absolutePath)) {
    return null;
  }

  return `/api/images/local?path=${encodeURIComponent(absolutePath)}`;
}

function extractHyperlinkFromFormula(formula: string): string | null {
  const match = /HYPERLINK\s*\(\s*"([^"]+)"/i.exec(formula);
  if (match?.[1]) {
    return match[1].trim();
  }
  return null;
}

function extractHyperlinkFromCell(cell: ExcelJS.Cell): string | null {
  const directHyperlink = (cell as { hyperlink?: unknown }).hyperlink;
  if (
    typeof directHyperlink === "string" &&
    directHyperlink.trim().length > 0
  ) {
    return directHyperlink.trim();
  }

  const rawValue = cell.value;
  if (rawValue && typeof rawValue === "object") {
    if ("hyperlink" in rawValue && typeof rawValue.hyperlink === "string") {
      const hyperlink = rawValue.hyperlink.trim();
      if (hyperlink.length > 0) {
        return hyperlink;
      }
    }

    if ("formula" in rawValue && typeof rawValue.formula === "string") {
      const fromFormula = extractHyperlinkFromFormula(rawValue.formula);
      if (fromFormula) {
        return fromFormula;
      }
    }
  }

  const text = normalizeCellText(rawValue);
  if (/^(file:\/\/|https?:\/\/)/i.test(text)) {
    return text;
  }
  if (path.isAbsolute(text) || /^[a-zA-Z]:[\\/]/.test(text)) {
    return text;
  }

  return null;
}

function buildRows(
  worksheet: ExcelJS.Worksheet,
  columns: ColumnRuntimeMeta[],
  headerRowIndex: number,
  fileId: string,
  fileName: string,
): ParsedRow[] {
  const rows: ParsedRow[] = [];

  for (
    let rowIndex = headerRowIndex + 1;
    rowIndex <= worksheet.rowCount;
    rowIndex += 1
  ) {
    const row = worksheet.getRow(rowIndex);
    const values: Record<string, ParsedCell> = {};
    let hasData = false;

    for (const column of columns) {
      const cell = row.getCell(column.colNumber);
      const rawValue = cell.value;
      const textValue = normalizeCellText(rawValue);
      const hyperlink = extractHyperlinkFromCell(cell);
      const imageSrc = hyperlink
        ? toImageSrcFromHyperlink(hyperlink, fileName)
        : null;

      if (imageSrc) {
        values[column.key] = textValue
          ? {
              type: "image",
              src: imageSrc,
              srcList: [imageSrc],
              value: textValue,
            }
          : {
              type: "image",
              src: imageSrc,
              srcList: [imageSrc],
            };
        hasData = true;
        continue;
      }

      if (textValue) {
        hasData = true;
      }

      values[column.key] = {
        type: "text",
        value: textValue,
      };
    }

    if (!hasData) {
      continue;
    }

    rows.push({
      rowId: `${fileId}_${rowIndex}`,
      values,
    });
  }

  return rows;
}

function getLevelOptions(rows: ParsedRow[], columnKey?: string): string[] {
  if (!columnKey) {
    return [];
  }

  const unique = new Set<string>();
  for (const row of rows) {
    const value = row.values[columnKey]?.value?.trim();
    if (value) {
      unique.add(value);
    }
  }
  return Array.from(unique);
}

function validateRequiredColumns(columns: ColumnRuntimeMeta[]): void {
  const hasLevel1 = columns.some((column) =>
    matchesHeader(column.title, LEVEL1_ALIASES),
  );
  const hasLevel2 = columns.some((column) =>
    matchesHeader(column.title, LEVEL2_ALIASES),
  );

  const missing: string[] = [];
  if (!hasLevel1) {
    missing.push("level1");
  }
  if (!hasLevel2) {
    missing.push("level2");
  }

  if (missing.length > 0) {
    throw new Error(`缺少必需列: ${missing.join("、")}`);
  }
}

export async function parseWorkbook(
  buffer: Buffer,
  fileName: string,
  fileId: string,
): Promise<ParsedWorkbook> {
  const workbook = new ExcelJSRuntime.Workbook();
  type ExcelLoadInput = Parameters<ExcelJS.Workbook["xlsx"]["load"]>[0];
  await workbook.xlsx.load(buffer as unknown as ExcelLoadInput);

  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    throw new Error("Excel 中没有可解析的工作表");
  }

  const headerRowIndex = detectHeaderRow(worksheet);
  const headerRow = worksheet.getRow(headerRowIndex);
  const columns = buildColumns(headerRow);
  if (columns.length === 0) {
    throw new Error("未检测到有效表头");
  }

  validateRequiredColumns(columns);

  const rows = buildRows(worksheet, columns, headerRowIndex, fileId, fileName);

  const level1Key = columns.find((column) =>
    matchesHeader(column.title, LEVEL1_ALIASES),
  )?.key;
  const level2Key = columns.find((column) =>
    matchesHeader(column.title, LEVEL2_ALIASES),
  )?.key;

  return {
    fileId,
    fileName,
    columns: columns.map(({ colNumber: _colNumber, ...column }) => column),
    rows,
    level1Options: getLevelOptions(rows, level1Key),
    level2Options: getLevelOptions(rows, level2Key),
  };
}
