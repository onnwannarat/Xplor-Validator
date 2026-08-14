import ExcelJS from "exceljs";
import { DAY_LABELS, GROUP_COLOURS, HEADER_FILL, RAW_COLS, REMOVED_FILL, TEMPLATE_COLUMNS } from "./constants";
import { dayVal, daySummary } from "./helpers";
import type { DupeReportRow } from "./duplicates";
import type { ConflictRow } from "./recurringOverlaps";
import type { TemplateRow } from "./transform";

function argb(hex: string): string {
  return `FF${hex}`;
}
const thinBorder: ExcelJS.Border = { style: "thin", color: { argb: argb("BFBFBF") } };
const border: Partial<ExcelJS.Borders> = { top: thinBorder, bottom: thinBorder, left: thinBorder, right: thinBorder };

function styleHeader(worksheet: ExcelJS.Worksheet): void {
  const row = worksheet.getRow(1);
  row.eachCell((cell) => {
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: argb(HEADER_FILL) } };
    cell.font = { color: { argb: "FFFFFFFF" }, bold: true };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cell.border = border;
  });
  row.height = 30;
}

function writeStyledRow(worksheet: ExcelJS.Worksheet, values: (string | number)[], fill?: string, bold?: boolean): void {
  const row = worksheet.addRow(values);
  row.eachCell((cell) => {
    if (fill) cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: argb(fill) } };
    if (bold) cell.font = { bold: true };
    cell.border = border;
  });
}

function autoWidth(worksheet: ExcelJS.Worksheet, maxWidth = 40): void {
  worksheet.columns.forEach((column) => {
    let best = 8;
    column.eachCell?.({ includeEmpty: false }, (cell) => {
      if (cell.value !== null && cell.value !== undefined) best = Math.max(best, String(cell.value).length);
    });
    column.width = Math.min(best + 3, maxWidth);
  });
}

/** Adds an "About This Report" worksheet with label/description rows. Pass ["", ""] for a blank spacer row. */
function addNotesSheet(workbook: ExcelJS.Workbook, rows: [string, string][]): ExcelJS.Worksheet {
  const worksheet = workbook.addWorksheet("About This Report");
  worksheet.getColumn(1).width = 30;
  worksheet.getColumn(2).width = 90;
  appendNotesRows(worksheet, rows);
  return worksheet;
}

function appendNotesRows(worksheet: ExcelJS.Worksheet, rows: [string, string][]): void {
  for (const [label, text] of rows) {
    const row = worksheet.addRow([label, text]);
    row.getCell(1).font = { bold: true, color: { argb: argb("1F497D") } };
    row.getCell(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: argb("F2F2F2") } };
    row.getCell(2).fill = { type: "pattern", pattern: "solid", fgColor: { argb: argb("F2F2F2") } };
    row.getCell(2).alignment = { wrapText: true, vertical: "top" };
    row.height = Math.max(15, Math.min(15 * (Math.floor(text.length / 90) + 1), 60));
  }
}

/**
 * Builds the "Duplicate Booking Patterns" sheet + "About This Report" notes sheet.
 * Port of detect_duplicates_and_report's Excel-writing half.
 */
export function buildDuplicateReportWorkbook(dupeReportRows: DupeReportRow[], nDupeGroups: number): ExcelJS.Workbook {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Duplicate Booking Patterns");

  const header = [...RAW_COLS, "Booked Days Summary", "Duplicate Group", "Status"];
  worksheet.addRow(header);
  styleHeader(worksheet);

  if (dupeReportRows.length === 0) {
    worksheet.addRow(["No duplicate booking patterns found.", ...Array(header.length - 1).fill("")]);
  } else {
    let prevGroup: number | null = null;
    for (const { row, groupNum, isFirst } of dupeReportRows) {
      if (prevGroup !== null && groupNum !== prevGroup) {
        worksheet.addRow(Array(header.length).fill(""));
      }
      const fill = GROUP_COLOURS[(groupNum - 1) % GROUP_COLOURS.length];
      const status = isFirst ? "KEEP (first)" : "REMOVE (duplicate)";
      const values = RAW_COLS.map((c) => (c in DAY_LABELS ? dayVal(row[c]) : row[c] || ""));
      values.push(daySummary(row), `Group ${groupNum}`, status);
      writeStyledRow(worksheet, values, fill, isFirst);
      prevGroup = groupNum;
    }
  }

  autoWidth(worksheet);
  worksheet.views = [{ state: "frozen", ySplit: 1 }];

  const nDupeRows = dupeReportRows.length;
  addNotesSheet(workbook, [
    ["DUPLICATE BOOKING PATTERNS REPORT", ""],
    ["", ""],
    [
      "What is a duplicate?",
      "A booking row where ALL of the following fields are identical to another row: Service Legacy ID, Child Legacy ID, Child Name, Start Date, End Date, Fee Name, Room Name, Frequency, and Day Pattern (Monday–Sunday, Weeks 1 & 2).",
    ],
    ["", ""],
    ["HOW TO READ THIS REPORT", ""],
    ["Colour coding", "Each colour group represents one set of duplicate rows. All rows sharing the same colour are exact duplicates of each other."],
    ["Bold row → KEEP (first)", "The first occurrence of the duplicate. This row IS included in the import file."],
    ["Normal rows → REMOVE (duplicate)", "All subsequent occurrences. These rows are NOT included in the import file."],
    ["Booked Days Summary", "Human-readable list of booked days, e.g. 'Mon1, Wed1, Fri2'. Week 1 = standard week; Week 2 = second week of a fortnightly cycle."],
    ["Duplicate Group", "Sequential group number. All rows with the same number are duplicates of each other."],
    ["Status", "KEEP (first) = included in import file.  |  REMOVE (duplicate) = excluded from import file."],
    ["", ""],
    ["ACTION REQUIRED", ""],
    ["If duplicates look correct", "No action needed — the script has already excluded the extra rows from all import files."],
    ["If a row should NOT be a duplicate", "The two rows may have genuinely different intended bookings (e.g. different dates). Correct the source data and re-run the script."],
    ["", ""],
    ["Total duplicate groups found", String(nDupeGroups)],
    ["Total duplicate rows removed", String(Math.max(0, nDupeRows - nDupeGroups))],
  ]);

  return workbook;
}

/**
 * Appends a "Recurring Schedule Overlaps" sheet to the duplicate-report workbook,
 * plus a summary section on "About This Report". Port of _add_schedule_overlap_sheet.
 */
export function appendScheduleOverlapSheet(workbook: ExcelJS.Workbook, conflicts: ConflictRow[]): void {
  const worksheet = workbook.addWorksheet("Recurring Schedule Overlaps");
  const reportCols = [...TEMPLATE_COLUMNS, "ConflictReason"];
  worksheet.addRow(reportCols);
  styleHeader(worksheet);

  if (conflicts.length === 0) {
    worksheet.addRow(["No recurring schedule overlaps found.", ...Array(reportCols.length - 1).fill("")]);
  } else {
    const sorted = [...conflicts].sort((a, b) => {
      const sa = a.row.ServiceID ?? "";
      const sb = b.row.ServiceID ?? "";
      if (sa !== sb) return sa < sb ? -1 : 1;
      const ca = a.row.Child_Legacy_Id ?? "";
      const cb = b.row.Child_Legacy_Id ?? "";
      return ca < cb ? -1 : ca > cb ? 1 : 0;
    });

    let groupNum = 0;
    let prevKey: string | null = null;
    const groupKeys = new Set<string>();
    for (const { row, conflictReason } of sorted) {
      const key = `${row.ServiceID ?? ""}|${row.Child_Legacy_Id ?? ""}`;
      if (key !== prevKey) {
        if (prevKey !== null) worksheet.addRow(Array(reportCols.length).fill(""));
        groupNum += 1;
        prevKey = key;
        groupKeys.add(key);
      }
      const fill = GROUP_COLOURS[(groupNum - 1) % GROUP_COLOURS.length];
      const values = TEMPLATE_COLUMNS.map((c) => row[c] ?? "");
      values.push(conflictReason);
      writeStyledRow(worksheet, values, fill);
    }
    autoWidth(worksheet);
  }
  worksheet.views = [{ state: "frozen", ySplit: 1 }];

  const notesWs = workbook.getWorksheet("About This Report");
  if (notesWs) {
    const nGroups = new Set(conflicts.map((c) => `${c.row.ServiceID ?? ""}|${c.row.Child_Legacy_Id ?? ""}`)).size;
    appendNotesRows(notesWs, [
      ["", ""],
      ["RECURRING SCHEDULE OVERLAPS", ""],
      ["", ""],
      [
        "What is a recurring schedule overlap?",
        "Two (or more) recurring booking rows for the same child and service whose date ranges overlap AND which share at least one common booked weekday. This means the child would be double-booked in the same slot.",
      ],
      [
        "Action taken",
        "Within each conflicting group, the row with the furthest (latest) EndDate is kept and included in the Recurring bookings_import file. The other row(s) in the group have been removed and appear in the 'Recurring Schedule Overlaps' sheet of this workbook.",
      ],
      ["If removal looks incorrect", "Verify the date ranges and day patterns in the source data. Correct the source and re-run the script."],
      ["", ""],
      ["Total conflict groups found", String(nGroups)],
      ["Total rows removed", String(conflicts.length)],
    ]);
  }
}

/** Builds the "Removed Casual Bookings" report workbook. Port of the removed-overlap-report block in main(). */
export function buildRemovedOverlapReportWorkbook(removedRows: { row: TemplateRow; overlapReason: string }[]): ExcelJS.Workbook {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Removed Casual Bookings");
  const reportCols = [...TEMPLATE_COLUMNS, "OverlapReason"];
  worksheet.addRow(reportCols);
  styleHeader(worksheet);

  for (const { row, overlapReason } of removedRows) {
    const values = TEMPLATE_COLUMNS.map((c) => row[c] ?? "");
    values.push(overlapReason);
    writeStyledRow(worksheet, values, REMOVED_FILL);
  }
  autoWidth(worksheet);
  worksheet.views = [{ state: "frozen", ySplit: 1 }];

  addNotesSheet(workbook, [
    ["REMOVED CASUAL BOOKINGS REPORT", ""],
    ["", ""],
    [
      "Why were these rows removed?",
      "Each casual booking in this report conflicts with an existing recurring booking for the same child. The casual booking falls on a day that is already covered by the recurring pattern (same room, same day-of-week, within the recurring date range). Uploading it would place two bookings in the same room slot.",
    ],
    ["", ""],
    ["OVERLAP CONDITIONS", ""],
    [
      "Conditions 1–5 must all be true",
      "1. Same child (Child_Legacy_Id)  2. Same service (ServiceID)  3. Same room (ImportedRoom)  4. Casual date falls within the recurring booking's Start–End Date range  5. Casual date's weekday is a booked day in the recurring pattern (for Fortnightly: checked against the correct Week 1 or Week 2 column; for Weekly and other frequencies: checked against Week 1 columns only)",
    ],
    ["", ""],
    ["OVERLAP SUB-TYPES", ""],
    ["Same fee (direct double-charge)", "The casual and recurring bookings have the same fee. Uploading both would charge the parent twice at the same rate."],
    [
      "Same room, different fee",
      "The casual and recurring bookings share the same room and day but use different fee names. The slot is already occupied by the recurring booking; the casual fee would charge the parent an additional (usually higher) amount. The OverlapReason column identifies which sub-type applies to each row.",
    ],
    ["", ""],
    ["UPLOAD ORDER", ""],
    ["Step 1", "Upload ALL files in Output/Recurring/ first."],
    ["Step 2", "Upload ALL files in Output/Casual/ after Step 1 is complete."],
    ["Note", "The casual bookings in this report have already been excluded from the Output/Casual/ files — no further action is needed before uploading."],
    ["", ""],
    ["COLUMN GUIDE", ""],
    [
      "OverlapReason",
      "Describes which recurring booking caused the conflict: the casual date, day-of-week, week number (for Fortnightly), the recurring period (Start – End Date), booked days, Room, and Fee. For different-fee overlaps, both the casual and recurring fee names are shown.",
    ],
    ["", ""],
    ["ACTION REQUIRED", ""],
    ["If removal looks correct", "No action needed. The row is already excluded from the Casual import files."],
    [
      "If removal looks incorrect",
      "The casual booking may be a genuinely different session. Verify with the service and, if needed, re-add it manually after the recurring import is complete.",
    ],
    ["", ""],
    ["Total casual bookings removed", String(removedRows.length)],
  ]);

  return workbook;
}
