import Papa from "papaparse";
import ExcelJS from "exceljs";
import type { ParsedFile, Row } from "./types";

/** Mirrors validator_v2.py's normalise_key: trims whitespace from a header, leaves falsy values untouched. */
function normaliseKey(col: string | null | undefined): string {
  return col ? col.trim() : (col ?? "");
}

/**
 * Mirrors validator_v2.py's safe_str: coerces a cell value to a trimmed string,
 * stripping a spurious trailing ".0" the way pandas produces for integer-valued
 * Excel cells read as floats, so IDs compare identically regardless of source format.
 */
function safeStr(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (value instanceof Date) {
    // Matches pandas' default string form for a midnight Timestamp read via dtype=str.
    const y = value.getFullYear();
    const m = String(value.getMonth() + 1).padStart(2, "0");
    const d = String(value.getDate()).padStart(2, "0");
    const hh = String(value.getHours()).padStart(2, "0");
    const mm = String(value.getMinutes()).padStart(2, "0");
    const ss = String(value.getSeconds()).padStart(2, "0");
    return `${y}-${m}-${d} ${hh}:${mm}:${ss}`;
  }
  if (typeof value === "object" && value !== null && "result" in value) {
    // ExcelJS formula cell: { formula, result }
    return safeStr((value as { result: unknown }).result);
  }
  const s = String(value);
  if (s.endsWith(".0") && s.length > 2) {
    const core = s.slice(0, -2);
    if (/^-?\d+$/.test(core)) return core;
  }
  return s;
}

function stripBom(text: string): string {
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}

function parseCsvText(text: string): ParsedFile {
  const result = Papa.parse<Record<string, string>>(stripBom(text), {
    header: true,
    skipEmptyLines: true,
  });
  const fieldnames = (result.meta.fields ?? []).map(normaliseKey);
  const rows: Row[] = result.data.map((raw) => {
    const row: Row = {};
    for (const [key, value] of Object.entries(raw)) {
      row[normaliseKey(key)] = safeStr(value);
    }
    return row;
  });
  return { rows, fieldnames };
}

async function parseXlsxBuffer(buffer: ArrayBuffer): Promise<ParsedFile> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const worksheet = workbook.worksheets[0];
  if (!worksheet) {
    throw new Error("The workbook has no sheets.");
  }

  const headerRow = worksheet.getRow(1);
  const fieldnames: string[] = [];
  headerRow.eachCell({ includeEmpty: false }, (cell) => {
    fieldnames.push(normaliseKey(safeStr(cell.value)));
  });

  const rows: Row[] = [];
  for (let r = 2; r <= worksheet.rowCount; r++) {
    const excelRow = worksheet.getRow(r);
    if (excelRow.cellCount === 0) continue;
    const row: Row = {};
    let hasValue = false;
    fieldnames.forEach((field, idx) => {
      const cell = excelRow.getCell(idx + 1);
      const value = safeStr(cell.value);
      if (value) hasValue = true;
      row[field] = value;
    });
    if (hasValue) rows.push(row);
  }
  return { rows, fieldnames };
}

/** Reads a migration file (CSV or XLSX) and returns normalised rows + column order. Mirrors load_input_bytes. */
export async function loadInputFile(file: File): Promise<ParsedFile> {
  const suffix = file.name.toLowerCase().split(".").pop();
  if (suffix === "xlsx" || suffix === "xls") {
    const buffer = await file.arrayBuffer();
    return parseXlsxBuffer(buffer);
  }
  const text = await file.text();
  const parsed = parseCsvText(text);
  if (parsed.fieldnames.length === 0) {
    throw new Error("The file appears to be empty or has no header row.");
  }
  return parsed;
}

/** Same as loadInputFile but for already-read bytes (used for cross-service reference files). */
export async function loadInputBytes(bytes: ArrayBuffer, filename: string): Promise<ParsedFile> {
  const suffix = filename.toLowerCase().split(".").pop();
  if (suffix === "xlsx" || suffix === "xls") {
    return parseXlsxBuffer(bytes);
  }
  const text = stripBom(new TextDecoder("utf-8").decode(bytes));
  const parsed = parseCsvText(text);
  if (parsed.fieldnames.length === 0) {
    throw new Error("The file appears to be empty or has no header row.");
  }
  return parsed;
}

export function getChildName(row: Row): string {
  const first = (row.Child_First_Name ?? "").trim();
  const last = (row.Child_Last_Name ?? "").trim();
  const name = `${first} ${last}`.trim();
  return name || "Unknown";
}
