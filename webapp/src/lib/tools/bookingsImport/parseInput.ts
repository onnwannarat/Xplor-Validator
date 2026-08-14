import Papa from "papaparse";
import ExcelJS from "exceljs";
import { RAW_COLS } from "./constants";

export type RawRow = Record<string, string>;

function safeStr(value: unknown): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "object" && value !== null && "result" in value) {
    return safeStr((value as { result: unknown }).result);
  }
  return String(value);
}

function toRawColsOnly(row: Record<string, string>): RawRow {
  const out: RawRow = {};
  for (const col of RAW_COLS) out[col] = row[col] ?? "";
  return out;
}

async function parseCsv(file: File): Promise<RawRow[]> {
  const text = (await file.text()).replace(/^﻿/, "");
  const result = Papa.parse<Record<string, string>>(text, { header: true, skipEmptyLines: true });
  return result.data.map(toRawColsOnly);
}

async function parseXlsx(file: File): Promise<RawRow[]> {
  const buffer = await file.arrayBuffer();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);
  const ws = wb.worksheets[0];
  if (!ws) return [];

  const headers: string[] = [];
  ws.getRow(1).eachCell({ includeEmpty: false }, (cell) => headers.push(safeStr(cell.value).trim()));

  const rows: RawRow[] = [];
  for (let r = 2; r <= ws.rowCount; r++) {
    const excelRow = ws.getRow(r);
    if (excelRow.cellCount === 0) continue;
    const row: Record<string, string> = {};
    let hasValue = false;
    headers.forEach((h, idx) => {
      const v = safeStr(excelRow.getCell(idx + 1).value);
      if (v) hasValue = true;
      row[h] = v;
    });
    if (hasValue) rows.push(toRawColsOnly(row));
  }
  return rows;
}

/** Reads a booking export (CSV or XLSX), returning only the recognised source columns (missing ones filled with ""). */
export async function readBookingFile(file: File): Promise<RawRow[]> {
  const ext = file.name.toLowerCase().split(".").pop();
  return ext === "xlsx" || ext === "xls" ? parseXlsx(file) : parseCsv(file);
}
