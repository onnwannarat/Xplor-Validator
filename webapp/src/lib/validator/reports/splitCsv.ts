import type { ServiceMapping } from "../serviceMapping";
import type { RowEntry } from "../types";

/** Removes characters not permitted in filenames across Windows and macOS. */
export function sanitiseFilename(name: string): string {
  return name.replace(/[\\/:*?"<>|]/g, "_").trim();
}

function csvEscape(value: string): string {
  if (/[",\n\r]/.test(value)) {
    return `"${value.replace(/"/g, '""')}"`;
  }
  return value;
}

function rowsToCsv(fieldnames: string[], rows: Record<string, string>[]): string {
  const header = fieldnames.map(csvEscape).join(",");
  const lines = rows.map((row) => fieldnames.map((f) => csvEscape(row[f] ?? "")).join(","));
  // BOM prefix matches the Python tool's utf-8-sig encoding, for correct Excel rendering of special characters.
  return "﻿" + [header, ...lines].join("\r\n");
}

export interface SplitCsvResult {
  files: { filename: string; content: string }[];
  rowNumMap: Map<number, number>;
}

/**
 * Groups rows by Xplor ServiceID and produces one CSV per service, named
 * "{Service_Name}_families_import.csv" (falling back to the ID). Also returns a
 * map from original row number to the row's position within its service's file
 * (row 2 = first data row), needed so report row references stay correct.
 * Port of write_split_csvs.
 */
export function buildSplitCsvs(
  allRows: RowEntry[],
  fieldnames: string[],
  serviceMap: ServiceMapping,
): SplitCsvResult {
  const serviceEntries = new Map<string, RowEntry[]>();
  for (const entry of allRows) {
    const svcId = (entry.row.ServiceID ?? "").trim() || "Unknown";
    if (!serviceEntries.has(svcId)) serviceEntries.set(svcId, []);
    serviceEntries.get(svcId)!.push(entry);
  }

  const files: { filename: string; content: string }[] = [];
  const rowNumMap = new Map<number, number>();

  for (const [svcId, entries] of serviceEntries) {
    let svcName = entries.length > 0 ? (entries[0].row.Service_Name ?? "").trim() : "";
    if (!svcName && serviceMap.isLoaded) {
      const candidate = serviceMap.getNameByXplor(svcId);
      if (candidate !== svcId) svcName = candidate;
    }
    const label = sanitiseFilename(svcName || svcId || "Unknown");
    const filename = `${label}_families_import.csv`;

    entries.forEach((entry, idx) => {
      rowNumMap.set(entry.rowNum, idx + 2);
    });

    files.push({ filename, content: rowsToCsv(fieldnames, entries.map((e) => e.row)) });
  }

  return { files, rowNumMap };
}
