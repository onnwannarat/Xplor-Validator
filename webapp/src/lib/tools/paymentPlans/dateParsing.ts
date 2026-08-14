/**
 * Flexible date parser matching payment_plan_checker.py's DATE_FORMATS fallback
 * list exactly, including order — the order matters for ambiguous dates like
 * "03/04/2024" (DD/MM/YYYY formats are tried before MM/DD/YYYY, so day-first
 * wins for plain ambiguous dates, matching Australian convention).
 */

export interface ParsedDate {
  date: Date; // UTC-midnight-based components; only Y/M/D/H/Min/Sec are meaningful
  year: number;
  month: number; // 1-12
  day: number;
  hour: number;
  minute: number;
  second: number;
}

function isValidYmd(year: number, month: number, day: number): boolean {
  if (month < 1 || month > 12 || day < 1 || day > 31) return false;
  const d = new Date(Date.UTC(year, month - 1, day));
  return d.getUTCFullYear() === year && d.getUTCMonth() === month - 1 && d.getUTCDate() === day;
}

function to24Hour(hour12: number, meridiem: string): number {
  const isPM = meridiem.toUpperCase() === "PM";
  if (hour12 === 12) return isPM ? 12 : 0;
  return isPM ? hour12 + 12 : hour12;
}

function build(year: number, month: number, day: number, hour = 0, minute = 0, second = 0): ParsedDate | null {
  if (!isValidYmd(year, month, day)) return null;
  if (hour > 23 || minute > 59 || second > 59) return null;
  return { date: new Date(Date.UTC(year, month - 1, day, hour, minute, second)), year, month, day, hour, minute, second };
}

const TIME_12H_SEC = /^(\d{1,2}):(\d{2}):(\d{2})\s*(AM|PM)$/i;
const TIME_12H = /^(\d{1,2}):(\d{2})\s*(AM|PM)$/i;
const TIME_24H_SEC = /^(\d{1,2}):(\d{2}):(\d{2})$/;
const TIME_24H = /^(\d{1,2}):(\d{2})$/;

type Order = "dmy" | "ymd" | "mdy";

function parseDateAndOptionalTime(raw: string, order: Order, twoDigitYear: boolean): ParsedDate | null {
  const parts = raw.split(/\s+/);
  const datePart = parts[0];
  const timePart = parts.slice(1).join(" ");

  let y: number, m: number, d: number;
  if (order === "ymd") {
    const match = datePart.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (!match) return null;
    [y, m, d] = [Number(match[1]), Number(match[2]), Number(match[3])];
  } else {
    const sep = twoDigitYear ? /^(\d{1,2})\/(\d{1,2})\/(\d{2})$/ : /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/;
    const match = datePart.match(sep);
    if (!match) return null;
    const a = Number(match[1]);
    const b = Number(match[2]);
    let yr = Number(match[3]);
    if (twoDigitYear) yr = yr < 69 ? 2000 + yr : 1900 + yr;
    if (order === "dmy") {
      d = a;
      m = b;
    } else {
      m = a;
      d = b;
    }
    y = yr;
  }

  if (!timePart) return build(y, m, d);

  let match = timePart.match(TIME_12H_SEC);
  if (match) return build(y, m, d, to24Hour(Number(match[1]), match[4]), Number(match[2]), Number(match[3]));

  match = timePart.match(TIME_12H);
  if (match) return build(y, m, d, to24Hour(Number(match[1]), match[3]), Number(match[2]));

  match = timePart.match(TIME_24H_SEC);
  if (match) return build(y, m, d, Number(match[1]), Number(match[2]), Number(match[3]));

  match = timePart.match(TIME_24H);
  if (match) return build(y, m, d, Number(match[1]), Number(match[2]));

  return null;
}

// Mirrors DATE_FORMATS order exactly.
const FORMAT_ATTEMPTS: [Order, boolean][] = [
  ["dmy", false], // covers all 4 %d/%m/%Y (+time) variants and the bare date, tried before ymd/mdy
  ["ymd", false], // covers all %Y-%m-%d (+time) variants
  ["mdy", false], // covers all %m/%d/%Y (+time) variants
  ["dmy", true], // %d/%m/%y
];

/** Parses a date string using the same ordered fallback as payment_plan_checker.py's parse_date. */
export function parsePaymentPlanDate(raw: string): ParsedDate | null {
  const trimmed = raw.trim();
  if (!trimmed) return null;
  for (const [order, twoDigitYear] of FORMAT_ATTEMPTS) {
    const result = parseDateAndOptionalTime(trimmed, order, twoDigitYear);
    if (result) return result;
  }
  return null;
}

/** Formats a parsed date as DD/MM/YYYY, matching the tool's output format. */
export function formatDmy(parsed: ParsedDate): string {
  const dd = String(parsed.day).padStart(2, "0");
  const mm = String(parsed.month).padStart(2, "0");
  return `${dd}/${mm}/${parsed.year}`;
}

/** 0 = Monday ... 6 = Sunday, matching Python's date.weekday(). */
export function weekdayIndexMondayZero(parsed: ParsedDate): number {
  const jsDay = parsed.date.getUTCDay(); // 0 = Sunday ... 6 = Saturday
  return (jsDay + 6) % 7;
}

const WEEKDAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"];

export function weekdayName(parsed: ParsedDate): string {
  return WEEKDAY_NAMES[weekdayIndexMondayZero(parsed)];
}
