export type Row = Record<string, string>;

export interface RowEntry {
  rowNum: number;
  childName: string;
  row: Row;
}

export type Severity = "ERROR" | "WARNING" | "FIXED";

export interface Issue {
  Row: number;
  Child_Name: string;
  Field: string;
  Issue_Description: string;
  Severity_Level: Severity;
  Action_Taken: string;
  _tag?: string;
  // Arbitrary structured metadata attached by specialist checks (cross-service /
  // intra-file duplicate parent reports read these via `_<key>` lookups).
  [metaKey: `_${string}`]: unknown;
}

export interface ParentProfile {
  first_name: string;
  last_name: string;
  dob: string;
  contacts: Set<string>;
  emails: Set<string>;
  source_file: string;
  service_id: string;
  legacy_id: string;
  parent_crn: string;
}

export interface ParsedFile {
  rows: Row[];
  fieldnames: string[];
}

export interface ValidationResult {
  recorder: import("./issueRecorder").IssueRecorder;
  rows: RowEntry[];
  fieldnames: string[];
}

export interface RunOptions {
  includeWaitlist: boolean;
  existingFiles?: { bytes: ArrayBuffer; filename: string }[];
}
