import ExcelJS from "exceljs";
import type { CentreGroup, ResolvedRow } from "./process";
import { resolveRows } from "./process";

function sanitiseFilename(name: string): string {
  return name.replace(/[\\/:*?"<>|]/g, "_");
}

/** Loads a fresh copy of the template workbook so each output's colour palette / styles are independent. */
async function loadTemplate(templateBytes: ArrayBuffer): Promise<{ workbook: ExcelJS.Workbook; worksheet: ExcelJS.Worksheet }> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(templateBytes.slice(0));
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error("Template workbook has no active sheet.");
  return { workbook, worksheet };
}

/** Captures row 5's per-column style so it can be reapplied to every data row (matches the template's grey header/border look). */
function captureRowFiveStyles(worksheet: ExcelJS.Worksheet): Partial<ExcelJS.Style>[] {
  const styles: Partial<ExcelJS.Style>[] = [];
  for (let col = 1; col <= worksheet.columnCount; col++) {
    styles.push(JSON.parse(JSON.stringify(worksheet.getCell(5, col).style)));
  }
  return styles;
}

function writeResolvedRow(worksheet: ExcelJS.Worksheet, excelRow: number, styles: Partial<ExcelJS.Style>[], data: ResolvedRow): void {
  const values = [data.centreName, data.firstName, data.lastName, data.credit, data.owing];
  values.forEach((value, idx) => {
    const cell = worksheet.getCell(excelRow, idx + 1);
    cell.value = value ?? null;
    if (styles[idx]) cell.style = JSON.parse(JSON.stringify(styles[idx]));
  });
}

/** Creates one styled output workbook for a single centre. Port of write_output. */
export async function writeCentreOutput(group: CentreGroup, templateBytes: ArrayBuffer): Promise<{ filename: string; buffer: ArrayBuffer }> {
  const { workbook, worksheet } = await loadTemplate(templateBytes);
  const styles = captureRowFiveStyles(worksheet);
  const resolved = resolveRows(group.centreName, group.rows);

  resolved.forEach((row, idx) => writeResolvedRow(worksheet, 5 + idx, styles, row));

  const buffer = (await workbook.xlsx.writeBuffer()) as unknown as ArrayBuffer;
  return { filename: `${sanitiseFilename(group.centreName)}_Balance_Import.xlsx`, buffer };
}

/** Creates one consolidated workbook containing every centre's rows in sequence. Port of write_consolidated_output. */
export async function writeConsolidatedOutput(groups: CentreGroup[], templateBytes: ArrayBuffer): Promise<ArrayBuffer> {
  const { workbook, worksheet } = await loadTemplate(templateBytes);
  const styles = captureRowFiveStyles(worksheet);

  let excelRow = 5;
  for (const group of groups) {
    const resolved = resolveRows(group.centreName, group.rows);
    for (const row of resolved) {
      writeResolvedRow(worksheet, excelRow, styles, row);
      excelRow += 1;
    }
  }

  return workbook.xlsx.writeBuffer() as unknown as Promise<ArrayBuffer>;
}

/** Creates a report of every row whose (centre, first, last) — case-insensitive — appears more than once. Port of write_duplicate_report. Returns null if there are no duplicates. */
export async function writeDuplicateReport(groups: CentreGroup[], templateBytes: ArrayBuffer): Promise<ArrayBuffer | null> {
  const resolvedAll: ResolvedRow[] = [];
  for (const group of groups) resolvedAll.push(...resolveRows(group.centreName, group.rows));

  const keyOf = (r: ResolvedRow) => `${r.centreName.trim().toLowerCase()}|${r.firstName.trim().toLowerCase()}|${r.lastName.trim().toLowerCase()}`;
  const counts = new Map<string, number>();
  for (const r of resolvedAll) counts.set(keyOf(r), (counts.get(keyOf(r)) ?? 0) + 1);

  const duplicates = resolvedAll.filter((r) => (counts.get(keyOf(r)) ?? 0) > 1);
  if (duplicates.length === 0) return null;

  const { workbook, worksheet } = await loadTemplate(templateBytes);
  const styles = captureRowFiveStyles(worksheet);
  duplicates.forEach((row, idx) => writeResolvedRow(worksheet, 5 + idx, styles, row));

  return workbook.xlsx.writeBuffer() as unknown as Promise<ArrayBuffer>;
}
