// Ported 1:1 from validator_v2.py (config section, lines 49-263).
// Business rules owned by the Onboarding team — do not "simplify" without checking
// against the Python source, which remains the source of truth until parity is confirmed.

export const MANDATORY_CHILD_FIELDS = [
  "ServiceID",
  "Child_Legacy_Id",
  "Child_First_Name",
  "Child_Last_Name",
  "DOB",
  "Status",
  "Child_CRN",
  "Room_Name",
  "Enrolment_Start_Date",
] as const;

// Waitlist records are exempt from these fields — enrolment date is not yet confirmed.
export const MANDATORY_ACTIVE_ONLY_FIELDS = new Set(["Enrolment_Start_Date"]);

export const MANDATORY_PARENT1_FIELDS = [
  "Parent1_CRN",
  "Parent1_Legacy_Account_ID",
  "Parent1_DOB",
  "Parent1_Email",
] as const;

export const MANDATORY_PARENT2_FIELDS = [
  "Parent2_CRN",
  "Parent2_Legacy_Account_ID",
  "Parent2_DOB",
  "Parent2_Email",
] as const;

export const EMERGENCY_CONTACT_LEGACY_ID_FIELDS: [string, string][] = [
  ["EmergencyContact1_LegacyID", "EmergencyContact1_First_Name"],
  ["EmergencyContact2_LegacyID", "EmergencyContact2_First_Name"],
  ["EmergencyContact3_LegacyID", "EmergencyContact3_First_Name"],
  ["EmergencyContact4_LegacyID", "EmergencyContact4_First_Name"],
  ["EmergencyContact5_LegacyID", "EmergencyContact5_First_Name"],
];

export const PAIRED_NAME_FIELDS: [string, string][] = [
  ["Parent1_First_Name", "Parent1_Last_Name"],
  ["Parent2_First_Name", "Parent2_Last_Name"],
  ["EmergencyContact1_First_Name", "EmergencyContact1_Last_Name"],
  ["EmergencyContact2_First_Name", "EmergencyContact2_Last_Name"],
  ["EmergencyContact3_First_Name", "EmergencyContact3_Last_Name"],
  ["EmergencyContact4_First_Name", "EmergencyContact4_Last_Name"],
  ["EmergencyContact5_First_Name", "EmergencyContact5_Last_Name"],
];

// Compared case-insensitively
export const VALID_STATUSES = new Set(["active", "inactive", "waitlist"]);

export const VALID_GENDERS = new Set(["Male", "Female"]);

export const STANDARD_GENDER_IDENTITIES = new Set([
  "Male",
  "Female",
  "Non-Binary",
  "Trans Female",
  "Trans Male",
]);

// The Migration Tool auto-converts Yes/No and True/False to 1/0
export const VALID_BOOLEAN_VALUES = new Set(["0", "1", "yes", "no", "true", "false"]);

export const VALID_AU_STATES = new Set(["NSW", "VIC", "QLD", "SA", "WA", "TAS", "ACT", "NT"]);

export const DATE_FIELDS = [
  "DOB",
  "Enrolment_Start_Date",
  "Medicare_Expiry_Date",
  "Parent1_DOB",
  "Parent2_DOB",
];

export const EMAIL_FIELDS = [
  "Parent1_Email",
  "Parent1_Work_Email",
  "Parent2_Email",
  "Parent2_Work_Email",
  "EmergencyContact1_Email",
  "EmergencyContact2_Email",
  "EmergencyContact3_Email",
  "EmergencyContact4_Email",
  "EmergencyContact5_Email",
];

export const PHONE_FIELDS = [
  "Parent1_Contact_Mobile",
  "Parent1_Contact_Home",
  "Parent1_Work_Phone",
  "Parent2_Contact_Mobile",
  "Parent2_Contact_Home",
  "Parent2_Work_Phone",
  "Medical_Practitioner_Phone",
  "EmergencyContact1_Contact_Number",
  "EmergencyContact2_Contact_Number",
  "EmergencyContact3_Contact_Number",
  "EmergencyContact4_Contact_Number",
  "EmergencyContact5_Contact_Number",
];

export const CRN_FIELDS = ["Child_CRN", "Parent1_CRN", "Parent2_CRN", "Enrolment_Parent_CRN"];

export const BOOLEAN_FIELDS = [
  "Consents_Photos_Videos",
  "Epipen/Anipen",
  "EmergencyContact1_Emergency_Contact",
  "EmergencyContact1_Medical_Nominee",
  "EmergencyContact1_Collection_Nominee",
  "EmergencyContact1_Excursion_Nominee",
  "EmergencyContact2_Emergency_Contact",
  "EmergencyContact2_Medical_Nominee",
  "EmergencyContact2_Collection_Nominee",
  "EmergencyContact2_Excursion_Nominee",
  "EmergencyContact3_Emergency_Contact",
  "EmergencyContact3_Medical_Nominee",
  "EmergencyContact3_Collection_Nominee",
  "EmergencyContact3_Excursion_Nominee",
  "EmergencyContact4_Emergency_Contact",
  "EmergencyContact4_Medical_Nominee",
  "EmergencyContact4_Collection_Nominee",
  "EmergencyContact4_Excursion_Nominee",
  "EmergencyContact5_Emergency_Contact",
  "EmergencyContact5_Medical_Nominee",
  "EmergencyContact5_Collection_Nominee",
  "EmergencyContact5_Excursion_Nominee",
];

export const STATE_FIELDS = [
  "State",
  "Parent1_State",
  "Parent1_Work_State",
  "Parent2_State",
  "Parent2_Work_State",
  "EmergencyContact1_State",
  "EmergencyContact2_State",
  "EmergencyContact3_State",
  "EmergencyContact4_State",
  "EmergencyContact5_State",
];

export const POSTCODE_FIELDS = [
  "PostCode",
  "Parent1_Post_Code",
  "Parent1_Work_Postcode",
  "Parent2_Post_Code",
  "Parent2_Work_Postcode",
  "EmergencyContact1_Postcode",
  "EmergencyContact2_Postcode",
  "EmergencyContact3_Postcode",
  "EmergencyContact4_Postcode",
  "EmergencyContact5_Postcode",
];

export const FIELD_LENGTH_LIMITS: Record<string, number> = {
  Child_First_Name: 100,
  Child_Middle_Name: 100,
  Child_Last_Name: 100,
  Gender: 100,
  Gender_Identity: 100,
  School: 200,
  Class: 255,
  Address: 255,
  Suburb: 255,
  Country: 255,
  State: 255,
  PostCode: 30,
  Religion: 100,
  Language: 255,
  Cultural_Background: 255,
  Cultural_Requirements: 255,
  Indigenous_Status: 255,
  Medicare_Number: 255,
  Medicare_Expiry_Date: 255,
  Ambulance_Cover_Number: 255,
  Health_Care_Centre: 255,
  Medical_Practitioner_Name: 255,
  Medical_Practitioner_Phone: 255,
  Medical_Practitioner_Address: 255,
  Child_CRN: 255,
  Child_Legacy_Id: 255,
  Parent1_CRN: 255,
  Parent1_Legacy_Account_ID: 100,
  Parent1_Title: 100,
  Parent1_First_Name: 100,
  Parent1_Middle_Name: 100,
  Parent1_Last_Name: 100,
  Parent1_Email: 200,
  Parent1_Contact_Home: 20,
  Parent1_Gender: 50,
  Parent1_Address_2: 100,
  Parent1_Suburb: 255,
  Parent1_State: 255,
  Parent1_Post_Code: 255,
  Parent1_Indigenous_Status: 255,
  Parent1_Language: 255,
  Parent1_Cultural_Background: 255,
  Parent1_Work_Email: 255,
  Parent1_Work_Phone: 255,
  Parent1_Work_Address: 255,
  Parent1_Work_Suburb: 255,
  Parent1_Work_Postcode: 255,
  Parent1_Work_Country: 255,
  Parent1_Work_State: 255,
  Parent1_Country: 255,
  Parent2_CRN: 255,
  Parent2_Legacy_Account_ID: 100,
  Parent2_Email: 200,
  Parent2_Work_Email: 255,
  EmergencyContact1_Email: 60,
  EmergencyContact2_Email: 60,
  EmergencyContact3_Email: 60,
  EmergencyContact4_Email: 60,
  EmergencyContact5_Email: 60,
  EmergencyContact1_Contact_Number: 30,
  EmergencyContact2_Contact_Number: 30,
  EmergencyContact3_Contact_Number: 30,
  EmergencyContact4_Contact_Number: 30,
  EmergencyContact5_Contact_Number: 30,
  EmergencyContact1_First_Name: 100,
  EmergencyContact1_Last_Name: 100,
  EmergencyContact1_Postcode: 20,
  EmergencyContact2_First_Name: 100,
  EmergencyContact2_Last_Name: 100,
  EmergencyContact2_Postcode: 20,
  EmergencyContact3_First_Name: 100,
  EmergencyContact3_Last_Name: 100,
  EmergencyContact3_Postcode: 20,
  EmergencyContact4_First_Name: 100,
  EmergencyContact4_Last_Name: 100,
  EmergencyContact4_Postcode: 20,
  EmergencyContact5_First_Name: 100,
  EmergencyContact5_Last_Name: 100,
  EmergencyContact5_Postcode: 20,
};

export const STATE_NORMALISATION_MAP: Record<string, string> = {
  "new south wales": "NSW",
  "new south wale": "NSW",
  nsw: "NSW",
  "n.s.w": "NSW",
  "n.s.w.": "NSW",
  victoria: "VIC",
  vic: "VIC",
  v: "VIC",
  queensland: "QLD",
  qld: "QLD",
  queesland: "QLD",
  queenslnd: "QLD",
  "south australia": "SA",
  sa: "SA",
  "sth australia": "SA",
  "south aust": "SA",
  "s.a": "SA",
  "s.a.": "SA",
  "western australia": "WA",
  wa: "WA",
  "west australia": "WA",
  "w.a": "WA",
  "w.a.": "WA",
  tasmania: "TAS",
  tas: "TAS",
  tassie: "TAS",
  "australian capital territory": "ACT",
  act: "ACT",
  "a.c.t": "ACT",
  "a.c.t.": "ACT",
  canberra: "ACT",
  "northern territory": "NT",
  nt: "NT",
  "n.t": "NT",
  "n.t.": "NT",
};

// Emergency contact email fields in priority order (EC1 takes precedence)
export const EC_EMAIL_FIELDS = [
  "EmergencyContact1_Email",
  "EmergencyContact2_Email",
  "EmergencyContact3_Email",
  "EmergencyContact4_Email",
  "EmergencyContact5_Email",
];

// Parent phone fields used for cross-service duplicate matching
export const PARENT_PHONE_FIELDS_BY_PREFIX: Record<string, string[]> = {
  Parent1: ["Parent1_Contact_Mobile", "Parent1_Contact_Home", "Parent1_Work_Phone"],
  Parent2: ["Parent2_Contact_Mobile", "Parent2_Contact_Home", "Parent2_Work_Phone"],
};

// Parent email fields used for cross-service duplicate matching
export const PARENT_EMAIL_FIELDS_BY_PREFIX: Record<string, string[]> = {
  Parent1: ["Parent1_Email", "Parent1_Work_Email"],
  Parent2: ["Parent2_Email", "Parent2_Work_Email"],
};

// ── Regex patterns ──────────────────────────────────────────────────────────

export const CRN_PATTERN = /^\d{9}[A-Za-z]$/;
export const PHONE_PATTERN =
  /^(\+?61[\s-]?|0)([2378]\d[\s-]?\d{4}[\s-]?\d{4}|4\d{2}[\s-]?\d{3}[\s-]?\d{3})$/;
export const EMAIL_PATTERN = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;
export const DATE_PATTERN_ISO = /^\d{4}-\d{2}-\d{2}$/;
export const DATE_PATTERN_DMY = /^\d{1,2}\/\d{1,2}\/\d{4}$/;
export const POSTCODE_PATTERN = /^\d{4}$/;

// ── Report styling ──────────────────────────────────────────────────────────

export const COLOUR_ERROR = "FFE4E6"; // Pastel blush pink
export const COLOUR_WARNING = "FFF8DC"; // Pastel cornsilk yellow
export const COLOUR_FIXED = "E6F4EA"; // Pastel mint green
export const COLOUR_HEADER = "4472C4"; // Xplor cornflower blue
export const COLOUR_SUMMARY = "EBF3FB"; // Very light sky blue
export const COLOUR_DUPLICATE = "FFF0E6"; // Pastel peach

export const CLIENT_ISSUE_TAGS = new Set(["duplicate_parent_email", "redundant_ec"]);
export const CLIENT_TAG_LABELS: Record<string, string> = {
  duplicate_parent_email: "Duplicate Parent Emails",
  redundant_ec: "Redundant Emergency Contacts",
};

export const REPORT_FIELDNAMES = [
  "Row",
  "Child_Name",
  "Field",
  "Issue_Description",
  "Severity_Level",
  "Action_Taken",
] as const;

export const DUPLICATE_PARENTS_REPORT_FIELDNAMES = [
  "Service_Name",
  "Parent_Legacy_ID",
  "Parent_Name",
  "Matched_On",
  "Parent_CRN",
  "Parent_Slot",
  "Linked_Child",
] as const;
