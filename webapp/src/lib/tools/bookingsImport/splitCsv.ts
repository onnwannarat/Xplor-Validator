import { TEMPLATE_COLUMNS } from "./constants";
import { sanitizeFilename } from "./helpers";
import type { TemplateRow } from "./transform";

function csvEscape(value: string): string {
  return /[",\n\r]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
}

function rowsToCsv(rows: TemplateRow[]): string {
  const header = [...TEMPLATE_COLUMNS].map(csvEscape).join(",");
  const lines = rows.map((row) => TEMPLATE_COLUMNS.map((c) => csvEscape(row[c] ?? "")).join(","));
  return "﻿" + [header, ...lines].join("\r\n");
}

export interface ServiceCsvFile {
  filename: string;
  content: string;
  rows: number;
}

/** Splits rows by ServiceID (ascending) and writes one CSV per service. Port of save_by_service. */
export function saveByService(rows: TemplateRow[], xplorIdToName: Map<string, string>, fileSuffix: string): ServiceCsvFile[] {
  if (rows.length === 0) return [];

  const groups = new Map<string, TemplateRow[]>();
  for (const row of rows) {
    const xplorId = (row.ServiceID ?? "").trim();
    if (!groups.has(xplorId)) groups.set(xplorId, []);
    groups.get(xplorId)!.push(row);
  }

  const sortedIds = [...groups.keys()].sort();
  return sortedIds.map((xplorId) => {
    const groupRows = groups.get(xplorId)!;
    const filename = xplorId
      ? `${sanitizeFilename(xplorIdToName.get(xplorId) ?? xplorId)}_${fileSuffix}.csv`
      : `UNMAPPED_${fileSuffix}.csv`;
    return { filename, content: rowsToCsv(groupRows), rows: groupRows.length };
  });
}
