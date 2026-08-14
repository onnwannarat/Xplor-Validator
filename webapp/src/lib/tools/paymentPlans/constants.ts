export type ColumnKey =
  | "service_id"
  | "service"
  | "parent_id"
  | "parent_fn"
  | "parent_ln"
  | "child_id"
  | "child_fn"
  | "child_ln"
  | "date"
  | "weekday"
  | "manual"
  | "cycle"
  | "limit"
  | "fixed_amount"
  | "gateway";

export const DEFAULT_COLUMNS: Record<ColumnKey, string> = {
  service_id: "Service ID",
  service: "Service Name",
  parent_id: "Parent Legacy ID",
  parent_fn: "Primary Guardian First Name",
  parent_ln: "Primary Guardian Last Name",
  child_id: "Child Legacy ID",
  child_fn: "Child First Name",
  child_ln: "Child Last Name",
  date: "Start Date",
  weekday: "Weekday",
  manual: "Manual",
  cycle: "Cycle",
  limit: "Limit",
  fixed_amount: "Fixed Amount",
  gateway: "Gateway Reference",
};

export const COLUMN_LABELS: Record<ColumnKey, string> = {
  service_id: "Service ID  *required",
  service: "Service Name",
  parent_id: "Parent Legacy ID",
  parent_fn: "Parent First Name",
  parent_ln: "Parent Last Name",
  child_id: "Child Legacy ID",
  child_fn: "Child First Name",
  child_ln: "Child Last Name",
  date: "Start Date",
  weekday: "Weekday",
  manual: "Manual (Yes/No)",
  cycle: "Billing Cycle",
  limit: "Direct Debit Limit",
  fixed_amount: "Fixed Amount",
  gateway: "Gateway Reference",
};

export const WEEKDAY_MAP: Record<string, string> = {
  monday: "Mon",
  tuesday: "Tue",
  wednesday: "Wed",
  thursday: "Thu",
  friday: "Fri",
  saturday: "Sat",
  sunday: "Sun",
  mon: "Mon",
  tue: "Tue",
  wed: "Wed",
  thu: "Thu",
  fri: "Fri",
  sat: "Sat",
  sun: "Sun",
};

export const VALID_CYCLES = new Set(["weekly", "fortnightly", "monthly"]);
export const WEEKEND_DAYS = new Set(["Sat", "Sun"]);

// Output CSV template — must match payment_plan_onboarding_tools.csv
export const TEMPLATE_COLUMNS = [
  "Service_ID",
  "Service_Name",
  "Parent_Legacy_Id",
  "Parent_First_Name",
  "Parent_Last_Name",
  "Child_Legacy_Id",
  "Child_First_Name",
  "Child_Last_Name",
  "Direct_Debit_Start_Date",
  "Direct_Debit_Day",
  "Manual",
  "Billing_Cycle",
  "Direct_Debit_Limit",
  "Fixed_Amount",
  "Gateway_Reference",
] as const;

export type ErrorCategory =
  | "weekend"
  | "missing_date"
  | "missing_weekday"
  | "missing_parent"
  | "missing_service_id"
  | "invalid_cycle"
  | "manual_not_monday"
  | "negative_limit"
  | "negative_fixed"
  | "both_amounts"
  | "unparseable_date"
  | "unknown_weekday";

export interface ErrorEntry {
  row: number;
  parent_id: string;
  child_id: string;
  parent_name: string;
  service: string;
  [extra: string]: string | number;
}

export type ErrorBuckets = Record<ErrorCategory, ErrorEntry[]>;

// (category key, Onboarding Tool error key, display label, severity)
export const ERROR_PRIORITY: { key: ErrorCategory; toolKey: string; label: string; severity: "weekend" | "error" | "warn" }[] = [
  { key: "weekend", toolKey: "ERROR_INVALID_PAYMENT_DAY", label: "WEEKEND — Sat / Sun", severity: "weekend" },
  { key: "missing_date", toolKey: "ERROR_MISSING_BOOKING_START_DATE", label: "Missing Start Date", severity: "error" },
  { key: "missing_weekday", toolKey: "ERROR_MISSING_PAYMENT_DAY", label: "Missing Weekday", severity: "error" },
  { key: "missing_parent", toolKey: "ERROR_MISSING_GUARDIAN", label: "Missing Parent Name", severity: "error" },
  { key: "missing_service_id", toolKey: "ERROR_MISSING_SERVICE_ID", label: "Missing Service ID", severity: "error" },
  { key: "invalid_cycle", toolKey: "ERROR_INVALID_FREQUENCY", label: "Invalid Billing Cycle", severity: "error" },
  { key: "manual_not_monday", toolKey: "MANUAL_PLAN_NOT_MONDAY", label: "Manual Plan — Not Monday", severity: "error" },
  { key: "negative_limit", toolKey: "ERROR_NEGATIVE_DIRECT_DEBIT_LIMIT", label: "Negative Direct Debit Limit", severity: "error" },
  { key: "negative_fixed", toolKey: "ERROR_NEGATIVE_FIXED_LIMIT", label: "Negative Fixed Amount", severity: "error" },
  { key: "both_amounts", toolKey: "ERROR_ONLY_ONE_AMOUNT_ALLOWED", label: "Both Limit + Fixed Amount Set", severity: "error" },
  { key: "unparseable_date", toolKey: "UNPARSEABLE_DATE", label: "Unparseable Date Value", severity: "warn" },
  { key: "unknown_weekday", toolKey: "UNKNOWN_WEEKDAY", label: "Unknown Weekday Value", severity: "warn" },
];

export function noteForError(key: ErrorCategory, e: ErrorEntry): string {
  switch (key) {
    case "weekend":
      return `Falls on a ${e.weekday === "Sun" ? "Sunday" : "Saturday"} (${e.date ?? ""}) — please verify before importing`;
    case "missing_date":
      return "Start Date is empty — cannot import";
    case "missing_weekday":
      return "Weekday is empty — cannot import";
    case "missing_parent":
      return `First: "${e.first_name ?? ""}"  Last: "${e.last_name ?? ""}"`;
    case "missing_service_id":
      return "Service ID is empty (required field)";
    case "invalid_cycle":
      return `Value: "${e.value ?? ""}" — must be Weekly / Fortnightly / Monthly`;
    case "manual_not_monday":
      return `Starts on ${e.day ?? ""} (${e.date ?? ""}) — Manual plans must start on Monday`;
    case "negative_limit":
      return `Value = ${e.value ?? ""}  (must be >= 0)`;
    case "negative_fixed":
      return `Value = ${e.value ?? ""}  (must be >= 0)`;
    case "both_amounts":
      return `Limit=${e.limit ?? ""}  AND  Fixed=${e.fixed ?? ""}  (only one allowed)`;
    case "unparseable_date":
      return `Cannot parse: "${e.value ?? ""}"`;
    case "unknown_weekday":
      return `Unrecognised value: "${e.value ?? ""}"`;
  }
}
