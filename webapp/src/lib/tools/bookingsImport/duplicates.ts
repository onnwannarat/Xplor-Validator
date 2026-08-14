import { DAY_LABELS, DUPE_KEY } from "./constants";
import type { RawRow } from "./parseInput";

function toComparisonRow(row: RawRow): RawRow {
  const cmp: RawRow = {};
  for (const [k, v] of Object.entries(row)) cmp[k] = typeof v === "string" ? v.trim() : v;
  for (const dayCol of Object.keys(DAY_LABELS)) {
    const n = Number(cmp[dayCol]);
    cmp[dayCol] = Number.isFinite(n) ? String(Math.trunc(n)) : "0";
  }
  return cmp;
}

function buildKey(cmpRow: RawRow): string {
  return DUPE_KEY.map((field) => cmpRow[field] ?? "").join("");
}

export interface DupeReportRow {
  row: RawRow; // trimmed/day-int comparison values, for display
  groupNum: number;
  isFirst: boolean;
}

export interface DuplicateDetectionResult {
  cleanRows: RawRow[]; // original raw rows, deduped (first occurrence of each exact-duplicate group kept)
  dupeReportRows: DupeReportRow[]; // sorted by group number, for the report sheet
  nDupeRows: number;
  nDupeGroups: number;
}

/**
 * Detects exact-duplicate booking rows (matching every DUPE_KEY field, with day
 * columns compared as booked/not-booked) and returns the de-duplicated rows plus
 * the data needed to render the duplicate report. Port of detect_duplicates_and_report
 * (computation only — Excel rendering lives in excelReports.ts).
 */
export function detectDuplicates(rawRows: RawRow[]): DuplicateDetectionResult {
  const cmpRows = rawRows.map(toComparisonRow);
  const keys = cmpRows.map(buildKey);

  const totalCounts = new Map<string, number>();
  for (const k of keys) totalCounts.set(k, (totalCounts.get(k) ?? 0) + 1);
  const isDupe = keys.map((k) => (totalCounts.get(k) ?? 0) > 1);

  const dupeRank: number[] = [];
  const seenCounts = new Map<string, number>();
  for (const k of keys) {
    const count = seenCounts.get(k) ?? 0;
    dupeRank.push(count);
    seenCounts.set(k, count + 1);
  }

  const keepMask = keys.map((_, i) => !isDupe[i] || dupeRank[i] === 0);
  const cleanRows = rawRows.filter((_, i) => keepMask[i]);

  const groupNumForKey = new Map<string, number>();
  let nextGroup = 1;
  const dupeReportRows: DupeReportRow[] = [];
  for (let i = 0; i < rawRows.length; i++) {
    if (!isDupe[i]) continue;
    const k = keys[i];
    if (!groupNumForKey.has(k)) {
      groupNumForKey.set(k, nextGroup);
      nextGroup += 1;
    }
    dupeReportRows.push({ row: cmpRows[i], groupNum: groupNumForKey.get(k)!, isFirst: dupeRank[i] === 0 });
  }
  dupeReportRows.sort((a, b) => a.groupNum - b.groupNum); // stable sort — preserves within-group original order

  return {
    cleanRows,
    dupeReportRows,
    nDupeRows: dupeReportRows.length,
    nDupeGroups: nextGroup - 1,
  };
}
