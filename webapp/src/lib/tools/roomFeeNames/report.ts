import ExcelJS from "exceljs";
import type { NameMismatch } from "./logic";

function writeSheet(
  workbook: ExcelJS.Workbook,
  title: string,
  rows: NameMismatch[],
  nameColumnLabel: string,
): void {
  const worksheet = workbook.addWorksheet(title);
  const headers = ["QK Service ID", "Xplor Service ID", "Service Name", nameColumnLabel, "Possible Match in Xplor"];
  worksheet.addRow(headers);
  for (const row of rows) {
    worksheet.addRow([row.qkServiceId, row.xplorServiceId, row.serviceName, row.nameInQk, row.possibleMatch]);
  }

  const headerRow = worksheet.getRow(1);
  headers.forEach((_, idx) => {
    const cell = headerRow.getCell(idx + 1);
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4472C4" } };
    cell.alignment = { horizontal: "center" };
  });

  worksheet.columns.forEach((column) => {
    let maxLen = 10;
    column.eachCell?.({ includeEmpty: false }, (cell) => {
      maxLen = Math.max(maxLen, cell.value ? String(cell.value).length : 0);
    });
    column.width = Math.min(maxLen + 4, 60);
  });
}

/** Builds the two-sheet mismatch report workbook. Port of _write_excel. */
export async function buildNameMismatchReport(
  feeMismatches: NameMismatch[],
  roomMismatches: NameMismatch[],
): Promise<ArrayBuffer> {
  const workbook = new ExcelJS.Workbook();
  writeSheet(workbook, "Fee Mismatches", feeMismatches, "Fee Name in QK");
  writeSheet(workbook, "Room Mismatches", roomMismatches, "Room Name in QK");
  return workbook.xlsx.writeBuffer() as unknown as Promise<ArrayBuffer>;
}
