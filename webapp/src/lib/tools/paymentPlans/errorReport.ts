import ExcelJS from "exceljs";
import { ERROR_PRIORITY, noteForError, type ErrorBuckets } from "./constants";

const C_WEEKEND_HEADER = "C0392B";
const C_WEEKEND_ROW = "FADBD8";
const C_ERROR_HEADER = "E67E22";
const C_ERROR_ROW = "FDEBD0";
const C_WARN_HEADER = "F1C40F";
const C_WARN_ROW = "FEF9E7";
const C_COL_HEADER = "1A5276";
const C_WHITE = "FFFFFF";
const C_LIGHT_GRAY = "F2F2F2";

function argb(hex: string): string {
  return `FF${hex}`;
}
function fill(hex: string): ExcelJS.Fill {
  return { type: "pattern", pattern: "solid", fgColor: { argb: argb(hex) } };
}
const thinBorder: ExcelJS.Border = { style: "thin", color: { argb: argb("CCCCCC") } };
const border: Partial<ExcelJS.Borders> = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };
const center: Partial<ExcelJS.Alignment> = { horizontal: "center", vertical: "middle", wrapText: true };
const left: Partial<ExcelJS.Alignment> = { horizontal: "left", vertical: "middle", wrapText: true };

function headerColourFor(severity: "weekend" | "error" | "warn"): string {
  return severity === "weekend" ? C_WEEKEND_HEADER : severity === "error" ? C_ERROR_HEADER : C_WARN_HEADER;
}
function rowColourFor(severity: "weekend" | "error" | "warn"): string {
  return severity === "weekend" ? C_WEEKEND_ROW : severity === "error" ? C_ERROR_ROW : C_WARN_ROW;
}

/** Builds the 3-sheet colour-coded error report workbook. Port of write_error_report. */
export async function buildPaymentPlanErrorReport(errors: ErrorBuckets, sourceFilename: string): Promise<ArrayBuffer> {
  const workbook = new ExcelJS.Workbook();

  // ── Sheet 1: Summary ────────────────────────────────────────────────────
  const wsSum = workbook.addWorksheet("Summary", { views: [{ showGridLines: false }] });
  wsSum.mergeCells("A1:D1");
  wsSum.getCell("A1").value = "Payment Plan Import — Error Report";
  wsSum.getCell("A1").font = { bold: true, size: 16, color: { argb: argb(C_WHITE) } };
  wsSum.getCell("A1").fill = fill(C_COL_HEADER);
  wsSum.getCell("A1").alignment = center;
  wsSum.getRow(1).height = 30;

  wsSum.mergeCells("A2:D2");
  wsSum.getCell("A2").value = `Generated: ${new Date().toLocaleString("en-AU")}  |  Source: ${sourceFilename}`;
  wsSum.getCell("A2").font = { italic: true, size: 9, color: { argb: argb("666666") } };
  wsSum.getCell("A2").alignment = left;
  wsSum.getRow(2).height = 16;

  wsSum.addRow([]);
  const sumHeaderRow = wsSum.addRow(["Error Category", "Onboarding Tool Error Key", "Count", "Severity"]);
  sumHeaderRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: argb(C_WHITE) } };
    cell.fill = fill(C_COL_HEADER);
    cell.alignment = center;
    cell.border = border;
  });
  sumHeaderRow.height = 20;

  let total = 0;
  for (const { key, toolKey, label, severity } of ERROR_PRIORITY) {
    const n = errors[key].length;
    total += n;
    const row = wsSum.addRow([label, toolKey, n, severity.toUpperCase()]);
    const rowFill = fill(rowColourFor(severity));
    row.eachCell((cell, colNumber) => {
      cell.fill = rowFill;
      cell.alignment = colNumber <= 2 ? left : center;
      cell.border = border;
      cell.font = { size: 10 };
    });
    row.getCell(3).font = n > 0 ? { bold: true, color: { argb: argb("C0392B") } } : { color: { argb: argb("2ECC71") } };
    row.height = 18;
  }

  const totalRow = wsSum.addRow(["TOTAL ERRORS", "", total, ""]);
  totalRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: argb(C_WHITE) }, size: 11 };
    cell.fill = fill(total > 0 ? "C0392B" : "1E8449");
    cell.alignment = center;
    cell.border = border;
  });
  totalRow.height = 22;

  wsSum.getColumn(1).width = 36;
  wsSum.getColumn(2).width = 38;
  wsSum.getColumn(3).width = 10;
  wsSum.getColumn(4).width = 14;

  // ── Sheet 2: Error Detail ───────────────────────────────────────────────
  const wsDet = workbook.addWorksheet("Error Detail", { views: [{ showGridLines: false, state: "frozen", ySplit: 2 }] });
  wsDet.mergeCells("A1:I1");
  wsDet.getCell("A1").value = "Error Detail — All Issues Found";
  wsDet.getCell("A1").font = { bold: true, size: 13, color: { argb: argb(C_WHITE) } };
  wsDet.getCell("A1").fill = fill(C_COL_HEADER);
  wsDet.getCell("A1").alignment = center;
  wsDet.getRow(1).height = 24;

  const detCols = ["#", "Error Category", "Tool Error Key", "Row #", "Parent Legacy ID", "Child Legacy ID", "Parent Name", "Service Name", "Note / Detail"];
  const detHeaderRow = wsDet.addRow(detCols);
  detHeaderRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: argb(C_WHITE) } };
    cell.fill = fill(C_COL_HEADER);
    cell.alignment = center;
    cell.border = border;
  });
  detHeaderRow.height = 20;

  let seq = 0;
  for (const { key, toolKey, label, severity } of ERROR_PRIORITY) {
    const items = errors[key];
    if (items.length === 0) continue;

    const dividerRow = wsDet.addRow(["", `── ${label.toUpperCase()}  (${items.length} rows) ──`, "", "", "", "", "", "", ""]);
    wsDet.mergeCells(`B${dividerRow.number}:I${dividerRow.number}`);
    dividerRow.eachCell((cell) => {
      cell.fill = fill(headerColourFor(severity));
      cell.font = { bold: true, color: { argb: argb(C_WHITE) } };
      cell.alignment = left;
    });
    dividerRow.height = 18;

    const rowFillColour = fill(rowColourFor(severity));
    const altFill = fill(C_LIGHT_GRAY);

    items.forEach((e, i) => {
      seq += 1;
      const note = noteForError(key, e);
      const row = wsDet.addRow([seq, label, toolKey, e.row, e.parent_id, e.child_id, e.parent_name, e.service, note]);
      const rf = i % 2 === 0 ? rowFillColour : altFill;
      row.eachCell((cell, colNumber) => {
        cell.fill = rf;
        cell.border = border;
        cell.alignment = colNumber === 1 || colNumber === 4 ? center : left;
        cell.font = { size: 10 };
      });
      row.height = 16;
    });
  }

  wsDet.getColumn(1).width = 5;
  wsDet.getColumn(2).width = 30;
  wsDet.getColumn(3).width = 34;
  wsDet.getColumn(4).width = 7;
  wsDet.getColumn(5).width = 16;
  wsDet.getColumn(6).width = 16;
  wsDet.getColumn(7).width = 28;
  wsDet.getColumn(8).width = 40;
  wsDet.getColumn(9).width = 55;

  // ── Sheet 3: Weekend Errors ─────────────────────────────────────────────
  const wsWk = workbook.addWorksheet("Weekend Errors", { views: [{ showGridLines: false, state: "frozen", ySplit: 2 }] });
  wsWk.mergeCells("A1:G1");
  wsWk.getCell("A1").value = "*** WEEKEND ERRORS — Please verify before importing ***";
  wsWk.getCell("A1").font = { bold: true, size: 13, color: { argb: argb(C_WHITE) } };
  wsWk.getCell("A1").fill = fill(C_WEEKEND_HEADER);
  wsWk.getCell("A1").alignment = center;
  wsWk.getRow(1).height = 24;

  const wkCols = ["Row #", "Weekday", "Start Date", "Parent Legacy ID", "Child Legacy ID", "Parent Name", "Service Name"];
  const wkHeaderRow = wsWk.addRow(wkCols);
  wkHeaderRow.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: argb(C_WHITE) } };
    cell.fill = fill(C_WEEKEND_HEADER);
    cell.alignment = center;
    cell.border = border;
  });
  wkHeaderRow.height = 20;

  const weekendItems = errors.weekend;
  if (weekendItems.length > 0) {
    weekendItems.forEach((e, i) => {
      const row = wsWk.addRow([e.row, e.weekday, e.date, e.parent_id, e.child_id, e.parent_name, e.service]);
      const rf = fill(i % 2 === 0 ? C_WEEKEND_ROW : "F5B7B1");
      row.eachCell((cell, colNumber) => {
        cell.fill = rf;
        cell.border = border;
        cell.alignment = colNumber <= 3 ? center : left;
        cell.font = { size: 10 };
      });
      row.getCell(2).font = { bold: true, color: { argb: argb("C0392B") } };
      row.height = 16;
    });
  } else {
    wsWk.mergeCells("A3:G3");
    wsWk.getCell("A3").value = "✓ No Weekend Errors found";
    wsWk.getCell("A3").font = { bold: true, size: 11, color: { argb: argb("1E8449") } };
    wsWk.getCell("A3").alignment = center;
    wsWk.getCell("A3").fill = fill("D5F5E3");
  }

  wsWk.getColumn(1).width = 7;
  wsWk.getColumn(2).width = 10;
  wsWk.getColumn(3).width = 14;
  wsWk.getColumn(4).width = 16;
  wsWk.getColumn(5).width = 16;
  wsWk.getColumn(6).width = 30;
  wsWk.getColumn(7).width = 45;

  return workbook.xlsx.writeBuffer() as unknown as Promise<ArrayBuffer>;
}
