import ExcelJS from "exceljs";
import {
  CLIENT_ISSUE_TAGS,
  COLOUR_ERROR,
  COLOUR_FIXED,
  COLOUR_HEADER,
  COLOUR_SUMMARY,
  COLOUR_WARNING,
  REPORT_FIELDNAMES,
} from "../constants";
import type { IssueRecorder } from "../issueRecorder";
import type { ServiceMapping } from "../serviceMapping";
import type { Issue, RowEntry } from "../types";
import { sanitiseFilename } from "./splitCsv";

const FILL: Record<string, string> = {
  ERROR: COLOUR_ERROR,
  WARNING: COLOUR_WARNING,
  FIXED: COLOUR_FIXED,
};

function argb(hex: string): string {
  return `FF${hex}`;
}

/** Sort key: alphabetical by service name, 'Unknown' always last. Port of _svc_sort_key. */
function svcSortKey(svcId: string, serviceMap: ServiceMapping): [number, string] {
  if (svcId === "Unknown") return [1, ""];
  const name = serviceMap.isLoaded ? serviceMap.getNameByXplor(svcId) : svcId;
  return [0, name.toLowerCase()];
}

function sortServiceIds(ids: string[], serviceMap: ServiceMapping): string[] {
  return [...ids].sort((a, b) => {
    const [ga, na] = svcSortKey(a, serviceMap);
    const [gb, nb] = svcSortKey(b, serviceMap);
    if (ga !== gb) return ga - gb;
    return na < nb ? -1 : na > nb ? 1 : 0;
  });
}

/**
 * Replaces input row numbers with output row numbers inside "rows X, Y" description
 * text, without touching row numbers embedded in other values. Port of _remap_row_numbers_in_text.
 */
function remapRowNumbersInText(text: string, rowNumMap: Map<number, number>): string {
  return text.replace(/\brows\s+([\d,\s]+)/g, (_match, numsPart: string) => {
    const remapped = numsPart.split(/,\s*/).map((token) => {
      const old = parseInt(token.trim(), 10);
      if (Number.isNaN(old)) return token.trim();
      return String(rowNumMap.get(old) ?? old);
    });
    return "rows " + remapped.join(", ");
  });
}

function applyHeaderStyle(worksheet: ExcelJS.Worksheet, rowNum: number, numCols: number): void {
  const row = worksheet.getRow(rowNum);
  for (let col = 1; col <= numCols; col++) {
    const cell = row.getCell(col);
    cell.font = { bold: true, color: { argb: "FFFFFFFF" }, size: 10 };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: argb(COLOUR_HEADER) } };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  }
}

function autoSizeColumns(worksheet: ExcelJS.Worksheet, minWidth = 12, maxWidth = 60): void {
  worksheet.columns.forEach((column) => {
    let maxLen = 0;
    column.eachCell?.({ includeEmpty: false }, (cell) => {
      const len = cell.value ? String(cell.value).length : 0;
      maxLen = Math.max(maxLen, len);
    });
    column.width = Math.max(minWidth, Math.min(maxLen + 4, maxWidth));
  });
}

function remapIssueRows(issues: Issue[], rowNumMap: Map<number, number> | undefined): Issue[] {
  if (!rowNumMap) return issues;
  return issues.map((issue) => ({
    ...issue,
    Row: rowNumMap.get(issue.Row) ?? issue.Row,
    Issue_Description: remapRowNumbersInText(issue.Issue_Description, rowNumMap),
  }));
}

function buildRowToService(allRows: RowEntry[]): Map<number, string> {
  const map = new Map<number, string>();
  for (const entry of allRows) {
    map.set(entry.rowNum, (entry.row.ServiceID ?? "").trim());
  }
  return map;
}

function writeIssueSheet(worksheet: ExcelJS.Worksheet, issues: Issue[]): void {
  worksheet.addRow([...REPORT_FIELDNAMES]);
  for (const issue of issues) {
    worksheet.addRow(REPORT_FIELDNAMES.map((f) => issue[f]));
  }
  applyHeaderStyle(worksheet, 1, REPORT_FIELDNAMES.length);
  worksheet.views = [{ state: "frozen", ySplit: 1 }];
  worksheet.getRow(1).height = 30;

  const sevColIdx = REPORT_FIELDNAMES.indexOf("Severity_Level") + 1;
  const descColIdx = REPORT_FIELDNAMES.indexOf("Issue_Description") + 1;

  for (let r = 2; r <= worksheet.rowCount; r++) {
    const row = worksheet.getRow(r);
    const severity = String(row.getCell(sevColIdx).value ?? "");
    const fillColour = FILL[severity];
    if (fillColour) {
      for (let c = 1; c <= REPORT_FIELDNAMES.length; c++) {
        row.getCell(c).fill = { type: "pattern", pattern: "solid", fgColor: { argb: argb(fillColour) } };
      }
    }
    row.getCell(descColIdx).alignment = { wrapText: true };
  }
  autoSizeColumns(worksheet);
}

function writeSummarySheet(
  worksheet: ExcelJS.Worksheet,
  header: string[],
  rows: (string | number)[][],
): void {
  worksheet.addRow(header);
  for (const row of rows) worksheet.addRow(row);
  applyHeaderStyle(worksheet, 1, header.length);

  const summaryFill: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: argb(COLOUR_SUMMARY) } };
  for (let r = 2; r < worksheet.rowCount; r++) {
    for (let c = 1; c <= header.length; c++) {
      worksheet.getRow(r).getCell(c).fill = summaryFill;
    }
  }
  const lastRow = worksheet.getRow(worksheet.rowCount);
  for (let c = 1; c <= header.length; c++) {
    lastRow.getCell(c).font = { bold: true };
  }
  worksheet.views = [{ state: "frozen", ySplit: 1 }];
  autoSizeColumns(worksheet);
}

/**
 * Builds the main colour-coded audit workbook — one tab per service plus a Summary
 * tab. Port of write_excel_report.
 */
export async function buildExcelReport(
  recorder: IssueRecorder,
  allRows: RowEntry[],
  serviceMap: ServiceMapping,
  rowNumMap: Map<number, number> | undefined,
): Promise<ArrayBuffer> {
  const rowToService = buildRowToService(allRows);
  const issues = remapIssueRows(recorder.issues, rowNumMap);

  const grouped = new Map<string, Issue[]>();
  for (const issue of recorder.issues) {
    const svcId = rowToService.get(issue.Row) ?? "Unknown";
    if (!grouped.has(svcId)) grouped.set(svcId, []);
  }
  // Group the *remapped* issues (display data) using the original row's service — order must line up.
  recorder.issues.forEach((original, idx) => {
    const svcId = rowToService.get(original.Row) ?? "Unknown";
    grouped.get(svcId)!.push(issues[idx]);
  });

  const serviceIds = sortServiceIds([...grouped.keys()], serviceMap);

  const workbook = new ExcelJS.Workbook();

  let totalErrors = 0;
  let totalWarnings = 0;
  let totalFixed = 0;
  const summaryRows: (string | number)[][] = [];
  for (const svcId of serviceIds) {
    const svcIssues = grouped.get(svcId) ?? [];
    const errors = svcIssues.filter((i) => i.Severity_Level === "ERROR").length;
    const warnings = svcIssues.filter((i) => i.Severity_Level === "WARNING").length;
    const fixed = svcIssues.filter((i) => i.Severity_Level === "FIXED").length;
    const svcName = serviceMap.isLoaded ? serviceMap.getNameByXplor(svcId) : svcId;
    totalErrors += errors;
    totalWarnings += warnings;
    totalFixed += fixed;
    summaryRows.push([svcId, svcName, errors, warnings, fixed, errors + warnings + fixed]);
  }
  summaryRows.push(["TOTAL", "", totalErrors, totalWarnings, totalFixed, totalErrors + totalWarnings + totalFixed]);

  writeSummarySheet(
    workbook.addWorksheet("Summary"),
    ["Service ID", "Service Name", "Errors", "Warnings", "Fixed", "Total Issues"],
    summaryRows,
  );

  for (const svcId of serviceIds) {
    const svcName = serviceMap.isLoaded ? serviceMap.getNameByXplor(svcId) : svcId;
    const sheetName = sanitiseFilename(svcName || svcId || "Unknown").slice(0, 31);
    writeIssueSheet(workbook.addWorksheet(sheetName), grouped.get(svcId) ?? []);
  }

  return workbook.xlsx.writeBuffer() as unknown as Promise<ArrayBuffer>;
}

/**
 * Builds the client-facing workbook, filtered to duplicate-parent-email and
 * redundant-EC issues only. Port of write_client_excel_report.
 */
export async function buildClientExcelReport(
  recorder: IssueRecorder,
  allRows: RowEntry[],
  serviceMap: ServiceMapping,
  rowNumMap: Map<number, number> | undefined,
): Promise<ArrayBuffer> {
  const rowToService = buildRowToService(allRows);
  const clientIssuesOriginal = recorder.issues.filter((i) => i._tag && CLIENT_ISSUE_TAGS.has(i._tag));
  const clientIssues = remapIssueRows(clientIssuesOriginal, rowNumMap);

  const grouped = new Map<string, Issue[]>();
  clientIssuesOriginal.forEach((original, idx) => {
    const svcId = rowToService.get(original.Row) ?? "Unknown";
    if (!grouped.has(svcId)) grouped.set(svcId, []);
    grouped.get(svcId)!.push(clientIssues[idx]);
  });

  const serviceIds = sortServiceIds([...grouped.keys()], serviceMap);

  const workbook = new ExcelJS.Workbook();

  let totalDup = 0;
  let totalEc = 0;
  const summaryRows: (string | number)[][] = [];
  for (const svcId of serviceIds) {
    const svcIssues = grouped.get(svcId) ?? [];
    const dupCount = svcIssues.filter((i) => i._tag === "duplicate_parent_email").length;
    const ecCount = svcIssues.filter((i) => i._tag === "redundant_ec").length;
    const svcName = serviceMap.isLoaded ? serviceMap.getNameByXplor(svcId) : svcId;
    totalDup += dupCount;
    totalEc += ecCount;
    summaryRows.push([svcName, dupCount, ecCount, dupCount + ecCount]);
  }
  summaryRows.push(["TOTAL", totalDup, totalEc, totalDup + totalEc]);

  writeSummarySheet(
    workbook.addWorksheet("Summary"),
    ["Service Name", "Duplicate Parent Emails", "Redundant Emergency Contacts", "Total Issues"],
    summaryRows,
  );

  for (const svcId of serviceIds) {
    const svcName = serviceMap.isLoaded ? serviceMap.getNameByXplor(svcId) : svcId;
    const sheetName = sanitiseFilename(svcName || svcId || "Unknown").slice(0, 31);
    writeIssueSheet(workbook.addWorksheet(sheetName), grouped.get(svcId) ?? []);
  }

  return workbook.xlsx.writeBuffer() as unknown as Promise<ArrayBuffer>;
}
