import { DEFAULT_END_DATE } from "./constants";

export interface SimpleDate {
  year: number;
  month: number; // 1-12
  day: number;
}

function isValidYmd(year: number, month: number, day: number): boolean {
  if (month < 1 || month > 12 || day < 1 || day > 31) return false;
  const d = new Date(Date.UTC(year, month - 1, day));
  return d.getUTCFullYear() === year && d.getUTCMonth() === month - 1 && d.getUTCDate() === day;
}

const DMY_TIME_12H = /^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+\d{1,2}:\d{2}(:\d{2})?\s*(AM|PM)$/i;
const DMY_TIME_24H = /^(\d{1,2})\/(\d{1,2})\/(\d{4})\s+\d{1,2}:\d{2}(:\d{2})?$/;
const DMY_ONLY = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/;
const YMD_TIME = /^(\d{4})-(\d{1,2})-(\d{1,2})\s+\d{1,2}:\d{2}(:\d{2})?$/;
const YMD_ONLY = /^(\d{4})-(\d{1,2})-(\d{1,2})$/;

/**
 * Parses a single date string using the same ordered strict-format fallback as
 * prepare_bookings_import.py's parse_date, then a best-effort flexible fallback
 * (day-first for ambiguous numeric forms, native parsing for named-month forms)
 * mirroring pandas' final `dayfirst=True` pass.
 */
export function parseBookingDate(raw: string): SimpleDate | null {
  const trimmed = raw.trim();
  if (!trimmed) return null;

  for (const pattern of [DMY_TIME_12H, DMY_TIME_24H, DMY_ONLY]) {
    const m = trimmed.match(pattern);
    if (m) {
      const day = Number(m[1]);
      const month = Number(m[2]);
      const year = Number(m[3]);
      if (isValidYmd(year, month, day)) return { year, month, day };
    }
  }
  for (const pattern of [YMD_TIME, YMD_ONLY]) {
    const m = trimmed.match(pattern);
    if (m) {
      const year = Number(m[1]);
      const month = Number(m[2]);
      const day = Number(m[3]);
      if (isValidYmd(year, month, day)) return { year, month, day };
    }
  }

  // Flexible fallback for anything the strict formats missed (e.g. "6-8-2024" or named months).
  const dmySeparators = trimmed.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})/);
  if (dmySeparators) {
    const day = Number(dmySeparators[1]);
    const month = Number(dmySeparators[2]);
    const year = Number(dmySeparators[3]);
    if (isValidYmd(year, month, day)) return { year, month, day };
  }
  const native = new Date(trimmed);
  if (!Number.isNaN(native.getTime())) {
    return { year: native.getFullYear(), month: native.getMonth() + 1, day: native.getDate() };
  }
  return null;
}

export function formatDmy(date: SimpleDate): string {
  return `${String(date.day).padStart(2, "0")}/${String(date.month).padStart(2, "0")}/${date.year}`;
}

/**
 * Returns the fallback EndDate for a row whose EndDate is blank: the supplied
 * override if given, else 31/12 of the StartDate's year, else DEFAULT_END_DATE.
 * Port of _year_end_date.
 */
export function yearEndDate(startDate: string, override = ""): string {
  const trimmedOverride = override.trim();
  if (trimmedOverride) return trimmedOverride;
  const parsed = parseBookingDate(startDate);
  if (parsed) return `31/12/${parsed.year}`;
  return DEFAULT_END_DATE;
}

/** Parses a DD/MM/YYYY string (as produced by parseBookingDate + formatDmy) back into a comparable Date, or null. */
export function parseDmyStrict(value: string): Date | null {
  const m = value.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  const day = Number(m[1]);
  const month = Number(m[2]);
  const year = Number(m[3]);
  if (!isValidYmd(year, month, day)) return null;
  return new Date(Date.UTC(year, month - 1, day));
}

/** 0 = Monday ... 6 = Sunday, matching Python's date.weekday(). */
export function weekdayIndexMondayZero(date: Date): number {
  return (date.getUTCDay() + 6) % 7;
}
