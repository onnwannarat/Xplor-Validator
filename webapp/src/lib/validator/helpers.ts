import {
  CRN_PATTERN,
  DATE_PATTERN_DMY,
  DATE_PATTERN_ISO,
  EMAIL_PATTERN,
  PHONE_PATTERN,
} from "./constants";
import type { Row } from "./types";

export function isBlank(value: string | null | undefined): boolean {
  return value === null || value === undefined || value.trim() === "";
}

/** Returns true for a valid date in YYYY-MM-DD or D/MM/YYYY format. Rejects impossible dates (e.g. 29/02/2023). */
export function isValidDate(value: string): boolean {
  const v = value.trim();
  if (DATE_PATTERN_ISO.test(v)) {
    const [y, m, d] = v.split("-").map(Number);
    return isRealCalendarDate(y, m, d);
  }
  if (DATE_PATTERN_DMY.test(v)) {
    const [d, m, y] = v.split("/").map(Number);
    return isRealCalendarDate(y, m, d);
  }
  return false;
}

function isRealCalendarDate(year: number, month: number, day: number): boolean {
  const date = new Date(Date.UTC(year, month - 1, day));
  return (
    date.getUTCFullYear() === year && date.getUTCMonth() === month - 1 && date.getUTCDate() === day
  );
}

/** Parses YYYY-MM-DD or D/MM/YYYY into a UTC-midnight Date. Caller must check isValidDate first. */
export function parseDateValue(value: string): Date {
  const v = value.trim();
  if (DATE_PATTERN_ISO.test(v)) {
    const [y, m, d] = v.split("-").map(Number);
    return new Date(Date.UTC(y, m - 1, d));
  }
  const [d, m, y] = v.split("/").map(Number);
  return new Date(Date.UTC(y, m - 1, d));
}

export function today(): Date {
  const now = new Date();
  return new Date(Date.UTC(now.getFullYear(), now.getMonth(), now.getDate()));
}

export function isValidCrn(value: string): boolean {
  return CRN_PATTERN.test(value.trim());
}

/** Accepts standard AU mobile/landline formats and bare 9-digit numbers starting with '4'. */
export function isValidPhone(value: string): boolean {
  const cleaned = value.replace(/[\s\-()]/g, "");
  if (/^4\d{8}$/.test(cleaned)) return true;
  return PHONE_PATTERN.test(cleaned);
}

export function isValidEmail(value: string): boolean {
  return EMAIL_PATTERN.test(value.trim());
}

export function normaliseKey(col: string): string {
  return col ? col.trim() : col;
}

export function getChildNameFromRow(row: Row): string {
  const first = (row.Child_First_Name ?? "").trim();
  const last = (row.Child_Last_Name ?? "").trim();
  const name = `${first} ${last}`.trim();
  return name || "Unknown";
}

export function normalisePhoneForMatch(value: string): string {
  return value.replace(/[\s\-()]/g, "").toLowerCase();
}
