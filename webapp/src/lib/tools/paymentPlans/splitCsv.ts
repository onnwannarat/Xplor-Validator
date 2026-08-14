import type { ServiceMapping } from "@/lib/validator/serviceMapping";
import { ColumnKey, ErrorBuckets, TEMPLATE_COLUMNS } from "./constants";
import type { Row } from "./processCsv";

function sanitiseFilename(name: string): string {
  return name.replace(/[\\/:*?"<>|]/g, "_").trim();
}

function entry(row: Record<string, string>, col: Record<ColumnKey, string>, key: ColumnKey): string {
  const colName = col[key];
  return colName ? (row[colName] ?? "").trim() : "";
}

function csvEscape(value: string): string {
  return /[",\n\r]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
}

function rowsToCsv(rows: string[][]): string {
  return "﻿" + rows.map((r) => r.map(csvEscape).join(",")).join("\r\n");
}

export interface SplitCsvResult {
  files: { serviceName: string; filename: string; content: string }[];
  rowMap: Map<number, number>;
}

/**
 * Groups processed rows by service (Service ID column if present, else Service
 * Name column), resolves each group to an Xplor service via serviceMap (by
 * QKServiceID) or serviceNameMap (by matched QikKids service name), and writes
 * one CSV per resolved service using the 15-column import template. Unresolved
 * groups are skipped entirely. Port of write_split_csvs.
 */
export function buildPaymentPlanSplitCsvs(
  rows: Row[],
  col: Record<ColumnKey, string>,
  serviceMap: ServiceMapping,
  serviceNameMap: Map<string, [string, string]>,
): SplitCsvResult {
  const groups = new Map<string, Row[]>();
  for (const row of rows) {
    const sid = entry(row.data, col, "service_id");
    const key = sid || entry(row.data, col, "service");
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key)!.push(row);
  }

  const files: SplitCsvResult["files"] = [];
  const rowMap = new Map<number, number>();

  for (const [key, groupRows] of groups) {
    const [xplorIdDirect, nameDirect] = serviceMap.lookupByQk(key);
    let xplorId = xplorIdDirect;
    let serviceName = nameDirect;
    if (xplorId === null) {
      const viaName = serviceNameMap.get(key.toLowerCase());
      if (viaName) [xplorId, serviceName] = viaName;
    }
    if (xplorId === null) continue; // not in serviceIDs.csv — skip

    const dest = `${sanitiseFilename(serviceName ?? "")}_payment_plan_import.csv`;
    const templateRows: string[][] = [];

    groupRows.forEach((row, idx) => {
      const fileRowNum = idx + 2; // row 1 = header
      rowMap.set(row.rowNum, fileRowNum);
      const g = (k: ColumnKey) => entry(row.data, col, k);
      const childId = g("child_id");
      templateRows.push([
        xplorId!,
        serviceName ?? "",
        g("parent_id"),
        g("parent_fn"),
        g("parent_ln"),
        childId ? `${xplorId}_${childId}` : "",
        g("child_fn"),
        g("child_ln"),
        g("date"),
        g("weekday"),
        g("manual"),
        g("cycle"),
        g("limit"),
        g("fixed_amount"),
        g("gateway"),
      ]);
    });

    files.push({ serviceName: serviceName ?? "", filename: dest, content: rowsToCsv([[...TEMPLATE_COLUMNS], ...templateRows]) });
  }

  return { files, rowMap };
}

/** Replaces original-file row numbers with service-file row numbers in-place. Port of _translate_error_rows. */
export function translateErrorRows(errors: ErrorBuckets, rowMap: Map<number, number>): void {
  for (const items of Object.values(errors)) {
    for (const e of items) {
      const mapped = rowMap.get(e.row);
      if (mapped !== undefined) e.row = mapped;
    }
  }
}
