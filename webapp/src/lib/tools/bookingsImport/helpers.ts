import { DAY_LABELS } from "./constants";

/** Returns 1 if the day cell is booked, "" otherwise — mirrors _day_val's float-tolerant check ("1", "1.0", 1 all count). */
export function dayVal(value: unknown): 1 | "" {
  if (value === null || value === undefined || value === "") return "";
  const n = Number(value);
  return Number.isFinite(n) && Math.trunc(n) === 1 ? 1 : "";
}

/** Normalises a room name for comparison: strips non-ASCII/punctuation, lowercases. Port of _normalize_room. */
export function normalizeRoom(s: string): string {
  const noNonAscii = [...s].map((ch) => (ch.charCodeAt(0) <= 0x7f ? ch : " ")).join("");
  const alnumOnly = noNonAscii.replace(/[^a-zA-Z0-9 ]/g, " ");
  return alnumOnly.replace(/\s+/g, " ").trim().toLowerCase();
}

/** Human-readable booked-days string, e.g. "Mon1, Wed1, Fri2". Port of _day_summary. */
export function daySummary(row: Record<string, unknown>): string {
  const labels = Object.entries(DAY_LABELS)
    .filter(([col]) => row[col] === 1)
    .map(([, label]) => label);
  return labels.length > 0 ? labels.join(", ") : "None";
}

/** Converts a service name to a safe filename component. Port of sanitize_filename. */
export function sanitizeFilename(name: string): string {
  let safe = name.replace(/[^\w\s'-]/g, "");
  safe = safe.trim().replace(/[\s']+/g, "_");
  safe = safe.replace(/_+/g, "_");
  return safe.replace(/^_+|_+$/g, "");
}
