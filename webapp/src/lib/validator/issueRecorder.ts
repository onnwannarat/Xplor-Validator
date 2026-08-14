import { REPORT_FIELDNAMES } from "./constants";
import type { Issue, Severity } from "./types";

/**
 * Collects validation issues and transformation actions across all rows.
 * Port of validator_v2.py's IssueRecorder.
 *
 * Severity levels:
 *   ERROR   — must be resolved before importing
 *   WARNING — review recommended before importing
 *   FIXED   — issue was detected and automatically corrected by the tool
 */
export class IssueRecorder {
  issues: Issue[] = [];

  add(
    rowNum: number,
    childName: string,
    field: string,
    description: string,
    severity: Severity,
    options: { action?: string; tag?: string; meta?: Record<string, unknown> } = {},
  ): void {
    const entry: Issue = {
      Row: rowNum,
      Child_Name: childName,
      Field: field,
      Issue_Description: description,
      Severity_Level: severity,
      Action_Taken: options.action ?? "",
      _tag: options.tag ?? "",
    };
    if (options.meta) {
      for (const [key, value] of Object.entries(options.meta)) {
        entry[`_${key}`] = value;
      }
    }
    this.issues.push(entry);
  }

  errorCount(): number {
    return this.issues.filter((i) => i.Severity_Level === "ERROR").length;
  }

  warningCount(): number {
    return this.issues.filter((i) => i.Severity_Level === "WARNING").length;
  }

  fixedCount(): number {
    return this.issues.filter((i) => i.Severity_Level === "FIXED").length;
  }

  toCsvText(): string {
    const header = REPORT_FIELDNAMES.join(",");
    const escape = (v: unknown) => {
      const s = String(v ?? "");
      return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
    };
    const lines = this.issues.map((issue) =>
      REPORT_FIELDNAMES.map((field) => escape(issue[field])).join(","),
    );
    return [header, ...lines].join("\r\n");
  }
}
