import { detectCasualOverlaps } from "./casualOverlaps";
import { detectDuplicates } from "./duplicates";
import { appendScheduleOverlapSheet, buildDuplicateReportWorkbook, buildRemovedOverlapReportWorkbook } from "./excelReports";
import { readBookingFile } from "./parseInput";
import { detectRecurringScheduleOverlaps } from "./recurringOverlaps";
import { saveByService, type ServiceCsvFile } from "./splitCsv";
import { transformRows, type TemplateRow } from "./transform";
import type { ServiceMapping } from "@/lib/validator/serviceMapping";

function todayStamp(): string {
  const d = new Date();
  return `${d.getFullYear()}${String(d.getMonth() + 1).padStart(2, "0")}${String(d.getDate()).padStart(2, "0")}`;
}

function sortByServiceThenChild(rows: TemplateRow[]): TemplateRow[] {
  return [...rows].sort((a, b) => {
    const sa = a.ServiceID ?? "";
    const sb = b.ServiceID ?? "";
    if (sa !== sb) return sa < sb ? -1 : 1;
    const ca = a.Child_Legacy_Id ?? "";
    const cb = b.Child_Legacy_Id ?? "";
    return ca < cb ? -1 : ca > cb ? 1 : 0;
  });
}

export interface BookingsImportResult {
  nInputFiles: number;
  nRawRows: number;
  nDupeRows: number;
  nDupeGroups: number;
  nSchedConflictRows: number;
  nSchedConflictGroups: number;
  nRecurring: number;
  nCasual: number;
  nCasualRemoved: number;
  unmappedIds: Set<string>;
  recurringFiles: ServiceCsvFile[];
  casualFiles: ServiceCsvFile[];
  duplicateReportBuffer: ArrayBuffer;
  duplicateReportFilename: string;
  removedOverlapReportBuffer: ArrayBuffer | null;
  removedOverlapReportFilename: string | null;
}

/**
 * Orchestrates the full bookings-import pipeline, entirely client-side. Port of main()
 * in prepare_bookings_import.py (the CLI-only `_main_cli` fallback is dropped — it has
 * no browser equivalent and duplicates this same logic).
 */
export async function runBookingsImport(
  files: File[],
  serviceMap: ServiceMapping,
  defaultEndDate = "",
): Promise<BookingsImportResult> {
  const xplorIdToName = serviceMap.getXplorIdToNameMap();

  const rawFrames = await Promise.all(files.map(readBookingFile));
  const rawRowsAll = rawFrames.flat();
  if (rawRowsAll.length === 0) {
    throw new Error("No readable rows found in the uploaded files.");
  }

  const { cleanRows: rawClean, dupeReportRows, nDupeRows, nDupeGroups } = detectDuplicates(rawRowsAll);

  const { rows: templateRows, unmappedIds } = transformRows(rawClean, serviceMap, defaultEndDate);

  const isCasual = (row: TemplateRow) => (row.WeekType ?? "").toUpperCase() === "CASUAL";
  let recurring = templateRows.filter((r) => !isCasual(r));
  let casual = templateRows.filter(isCasual);

  const { cleanRows: recurringClean, conflicts } = detectRecurringScheduleOverlaps(recurring);
  recurring = recurringClean;

  const dupWorkbook = buildDuplicateReportWorkbook(dupeReportRows, nDupeGroups);
  appendScheduleOverlapSheet(dupWorkbook, conflicts);
  const duplicateReportBuffer = (await dupWorkbook.xlsx.writeBuffer()) as unknown as ArrayBuffer;

  let removedCount = 0;
  let removedRows: { row: TemplateRow; overlapReason: string }[] = [];
  if (recurring.length > 0 && casual.length > 0) {
    const overlapResults = detectCasualOverlaps(casual, recurring);
    removedRows = casual
      .map((row, idx) => ({ row, ...overlapResults[idx] }))
      .filter((r) => r.isOverlap)
      .map(({ row, reason }) => ({ row, overlapReason: reason }));
    removedCount = removedRows.length;
    casual = casual.filter((_, idx) => !overlapResults[idx].isOverlap);
  }

  recurring = sortByServiceThenChild(recurring);
  casual = sortByServiceThenChild(casual);

  const today = todayStamp();
  const recurringFiles = saveByService(recurring, xplorIdToName, `bookings_import_${today}`);
  const casualFiles = saveByService(casual, xplorIdToName, `casualbookings_import_${today}`);

  let removedOverlapReportBuffer: ArrayBuffer | null = null;
  let removedOverlapReportFilename: string | null = null;
  if (removedRows.length > 0) {
    const removedSorted = [...removedRows].sort((a, b) => {
      const sa = a.row.ServiceID ?? "";
      const sb = b.row.ServiceID ?? "";
      if (sa !== sb) return sa < sb ? -1 : 1;
      const ca = a.row.Child_Legacy_Id ?? "";
      const cb = b.row.Child_Legacy_Id ?? "";
      return ca < cb ? -1 : ca > cb ? 1 : 0;
    });
    const wb = buildRemovedOverlapReportWorkbook(removedSorted);
    removedOverlapReportBuffer = (await wb.xlsx.writeBuffer()) as unknown as ArrayBuffer;
    removedOverlapReportFilename = `removed_overlap_report_${today}.xlsx`;
  }

  const nSchedConflictGroups = new Set(conflicts.map((c) => `${c.row.ServiceID ?? ""}|${c.row.Child_Legacy_Id ?? ""}`)).size;

  return {
    nInputFiles: files.length,
    nRawRows: rawRowsAll.length,
    nDupeRows,
    nDupeGroups,
    nSchedConflictRows: conflicts.length,
    nSchedConflictGroups,
    nRecurring: recurring.length,
    nCasual: casual.length,
    nCasualRemoved: removedCount,
    unmappedIds,
    recurringFiles,
    casualFiles,
    duplicateReportBuffer,
    duplicateReportFilename: `duplicate_bookings_report_${today}.xlsx`,
    removedOverlapReportBuffer,
    removedOverlapReportFilename,
  };
}
