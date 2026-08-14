import Papa from "papaparse";
import { ColumnKey, DEFAULT_COLUMNS, ErrorBuckets, VALID_CYCLES, WEEKEND_DAYS, WEEKDAY_MAP } from "./constants";
import { formatDmy, parsePaymentPlanDate, weekdayIndexMondayZero, weekdayName, type ParsedDate } from "./dateParsing";

export interface Row {
  data: Record<string, string>;
  rowNum: number;
}

export interface ProcessStats {
  total: number;
  date_fixed: number;
  weekday_fixed: number;
  spaces_fixed: number;
}

export interface ProcessResult {
  processedRows: Row[];
  fieldnames: string[];
  errors: ErrorBuckets;
  stats: ProcessStats;
}

function emptyBuckets(): ErrorBuckets {
  return {
    weekend: [],
    missing_date: [],
    missing_weekday: [],
    missing_parent: [],
    missing_service_id: [],
    invalid_cycle: [],
    manual_not_monday: [],
    negative_limit: [],
    negative_fixed: [],
    both_amounts: [],
    unparseable_date: [],
    unknown_weekday: [],
  };
}

function entry(row: Record<string, string>, col: Record<ColumnKey, string>, key: ColumnKey): string {
  const colName = col[key];
  return colName ? (row[colName] ?? "").trim() : "";
}

function isNumeric(val: string): boolean {
  if (val.trim() === "") return false;
  return Number.isFinite(Number(val));
}

/** Port of process_csv — validates every row and applies the same in-place auto-fixes. */
export async function processCsv(file: File, col: Record<ColumnKey, string>): Promise<ProcessResult> {
  const text = (await file.text()).replace(/^﻿/, "");
  const parsed = Papa.parse<Record<string, string>>(text, { header: true, skipEmptyLines: true });
  const fieldnames = parsed.meta.fields ?? [];
  const rawRows = parsed.data;

  const errors = emptyBuckets();
  const stats: ProcessStats = { total: rawRows.length, date_fixed: 0, weekday_fixed: 0, spaces_fixed: 0 };
  const processed: Row[] = [];

  rawRows.forEach((rawRow, idx) => {
    const rowNum = idx + 2;
    const row: Record<string, string> = {};

    for (const [key, value] of Object.entries(rawRow)) {
      const stripped = (value ?? "").trim();
      if (stripped !== value) stats.spaces_fixed += 1;
      row[key] = stripped;
    }

    const g = (key: ColumnKey) => entry(row, col, key);

    const parentFn = g("parent_fn");
    const parentLn = g("parent_ln");
    const parentName = `${parentFn} ${parentLn}`.trim();
    const service = g("service");
    const parentId = g("parent_id");
    const childId = g("child_id");
    const ctx = { row: rowNum, parent_id: parentId, child_id: childId, parent_name: parentName, service };

    // 1. Fix & validate Start Date
    const rawDate = g("date");
    let parsedDt: ParsedDate | null = null;
    if (rawDate) {
      parsedDt = parsePaymentPlanDate(rawDate);
      if (parsedDt) {
        const fixed = formatDmy(parsedDt);
        if (fixed !== rawDate) stats.date_fixed += 1;
        const colName = col.date;
        if (colName) row[colName] = fixed;
      } else {
        errors.unparseable_date.push({ ...ctx, value: rawDate });
      }
    } else {
      errors.missing_date.push({ ...ctx, weekday: g("weekday") || "(empty)" });
    }

    // 2. Fix & validate Weekday
    const rawWd = g("weekday");
    const wdFixed = WEEKDAY_MAP[rawWd.toLowerCase()] ?? "";
    const colNameWd = col.weekday;
    if (wdFixed) {
      if (wdFixed !== rawWd) stats.weekday_fixed += 1;
      if (colNameWd) row[colNameWd] = wdFixed;
    } else if (rawWd) {
      errors.unknown_weekday.push({ ...ctx, value: rawWd });
    } else {
      errors.missing_weekday.push({ ...ctx, date: rawDate || "(empty)" });
    }

    // 3. Weekend check
    if (WEEKEND_DAYS.has(wdFixed)) {
      const dateStr = row[col.date] ?? "";
      errors.weekend.push({ ...ctx, weekday: wdFixed, date: dateStr });
    }

    // 4. Parent name
    if (!parentFn || !parentLn) {
      errors.missing_parent.push({ ...ctx, first_name: parentFn || "(empty)", last_name: parentLn || "(empty)" });
    }

    // 6. Billing Cycle validation
    const cycle = g("cycle");
    if (cycle && !VALID_CYCLES.has(cycle.toLowerCase())) {
      errors.invalid_cycle.push({ ...ctx, value: cycle });
    }

    // 7. Manual plan -> start date must be Monday
    const manualVal = g("manual").toLowerCase();
    const isManual = manualVal === "yes" || manualVal === "1" || manualVal === "true";
    if (isManual && parsedDt && weekdayIndexMondayZero(parsedDt) !== 0) {
      errors.manual_not_monday.push({ ...ctx, date: formatDmy(parsedDt), day: weekdayName(parsedDt) });
    }

    // 8. Limit / Fixed Amount checks
    const limitRaw = g("limit");
    const fixedRaw = g("fixed_amount");

    if (limitRaw && isNumeric(limitRaw) && Number(limitRaw) < 0) {
      errors.negative_limit.push({ ...ctx, value: limitRaw });
    }
    if (fixedRaw && isNumeric(fixedRaw) && Number(fixedRaw) < 0) {
      errors.negative_fixed.push({ ...ctx, value: fixedRaw });
    }

    const limitNum = limitRaw ? Number(limitRaw) : 0;
    const fixedNum = fixedRaw ? Number(fixedRaw) : 0;
    if (Number.isFinite(limitNum) && Number.isFinite(fixedNum) && limitNum > 0 && fixedNum > 0) {
      errors.both_amounts.push({ ...ctx, limit: limitRaw, fixed: fixedRaw });
    }

    processed.push({ data: row, rowNum });
  });

  return { processedRows: processed, fieldnames, errors, stats };
}

export { DEFAULT_COLUMNS };
