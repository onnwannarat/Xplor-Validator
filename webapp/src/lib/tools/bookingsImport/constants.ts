export const DEFAULT_END_DATE = "31/12/2026";

// Final column order must match booking_onboarding_tools.csv exactly
export const TEMPLATE_COLUMNS = [
  "ServiceID",
  "Service_Name",
  "Child_Legacy_Id",
  "Child_First_Name",
  "Child_Last_Name",
  "StartDate",
  "EndDate",
  "ImportedFee",
  "ImportedRoom",
  "WeekType",
  "MON1",
  "TUE1",
  "WED1",
  "THU1",
  "FRI1",
  "SAT1",
  "SUN1",
  "MON2",
  "TUE2",
  "WED2",
  "THU2",
  "FRI2",
  "SAT2",
  "SUN2",
  "QKCreatedDate",
  "QKCreatedVia",
] as const;

// Source column -> Template column mapping
export const COLUMN_MAP: Record<string, string> = {
  "Service Legacy ID": "ServiceID",
  "Service Name": "Service_Name",
  "Child Legacy ID": "Child_Legacy_Id",
  "Child First Name": "Child_First_Name",
  "Child Last Name": "Child_Last_Name",
  "Start Date": "StartDate",
  "End Date": "EndDate",
  "Fee Name": "ImportedFee",
  "Room Name": "ImportedRoom",
  Frequency: "WeekType",
  Monday1: "MON1",
  Tuesday1: "TUE1",
  Wednesday1: "WED1",
  Thursday1: "THU1",
  Friday1: "FRI1",
  Saturday1: "SAT1",
  Sunday1: "SUN1",
  Monday2: "MON2",
  Tuesday2: "TUE2",
  Wednesday2: "WED2",
  Thursday2: "THU2",
  Friday2: "FRI2",
  Saturday2: "SAT2",
  Sunday2: "SUN2",
  "Created Date": "QKCreatedDate",
  "Created via": "QKCreatedVia",
};

export const RAW_COLS = Object.keys(COLUMN_MAP);

// Weekday number (Monday=0) -> day column prefix
export const WEEKDAY_TO_COL: Record<number, string> = {
  0: "MON",
  1: "TUE",
  2: "WED",
  3: "THU",
  4: "FRI",
  5: "SAT",
  6: "SUN",
};

export const DAY_COLS = [
  "MON1",
  "TUE1",
  "WED1",
  "THU1",
  "FRI1",
  "SAT1",
  "SUN1",
  "MON2",
  "TUE2",
  "WED2",
  "THU2",
  "FRI2",
  "SAT2",
  "SUN2",
] as const;

// Source day-column -> short label (used in duplicate report)
export const DAY_LABELS: Record<string, string> = {
  Monday1: "Mon1",
  Tuesday1: "Tue1",
  Wednesday1: "Wed1",
  Thursday1: "Thu1",
  Friday1: "Fri1",
  Saturday1: "Sat1",
  Sunday1: "Sun1",
  Monday2: "Mon2",
  Tuesday2: "Tue2",
  Wednesday2: "Wed2",
  Thursday2: "Thu2",
  Friday2: "Fri2",
  Saturday2: "Sat2",
  Sunday2: "Sun2",
};

// Key that identifies an exact-duplicate booking row
export const DUPE_KEY = [
  "Service Legacy ID",
  "Child Legacy ID",
  "Child First Name",
  "Child Last Name",
  "Start Date",
  "End Date",
  "Fee Name",
  "Room Name",
  "Frequency",
  "Monday1",
  "Tuesday1",
  "Wednesday1",
  "Thursday1",
  "Friday1",
  "Saturday1",
  "Sunday1",
  "Monday2",
  "Tuesday2",
  "Wednesday2",
  "Thursday2",
  "Friday2",
  "Saturday2",
  "Sunday2",
];

export const HEADER_FILL = "2F75B6";
export const WARN_FILL = "FFEB9C";
export const REMOVED_FILL = "F4CCCC";
export const GROUP_COLOURS = [
  "FFF2CC",
  "FCE4D6",
  "DDEBF7",
  "E2EFDA",
  "F4CCCC",
  "D9D2E9",
  "D0E0E3",
  "FFE599",
  "CFE2F3",
  "EAD1DC",
];
