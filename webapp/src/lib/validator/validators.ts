import {
  BOOLEAN_FIELDS,
  CRN_FIELDS,
  DATE_FIELDS,
  EMAIL_FIELDS,
  FIELD_LENGTH_LIMITS,
  MANDATORY_ACTIVE_ONLY_FIELDS,
  MANDATORY_CHILD_FIELDS,
  MANDATORY_PARENT1_FIELDS,
  MANDATORY_PARENT2_FIELDS,
  EMERGENCY_CONTACT_LEGACY_ID_FIELDS,
  PAIRED_NAME_FIELDS,
  PHONE_FIELDS,
  POSTCODE_PATTERN,
  POSTCODE_FIELDS,
  STANDARD_GENDER_IDENTITIES,
  STATE_FIELDS,
  VALID_AU_STATES,
  VALID_BOOLEAN_VALUES,
  VALID_GENDERS,
  VALID_STATUSES,
} from "./constants";
import { isBlank, isValidCrn, isValidDate, isValidEmail, isValidPhone, parseDateValue, today } from "./helpers";
import type { IssueRecorder } from "./issueRecorder";
import type { Row } from "./types";

type ValidatorFn = (
  row: Row,
  rowNum: number,
  childName: string,
  recorder: IssueRecorder,
  headers: Set<string>,
) => void;

export const validateMandatoryChildFields: ValidatorFn = (row, rowNum, childName, recorder, headers) => {
  const isWaitlist = (row.Status ?? "").trim().toLowerCase() === "waitlist";
  for (const field of MANDATORY_CHILD_FIELDS) {
    if (!headers.has(field)) continue;
    if (isWaitlist && MANDATORY_ACTIVE_ONLY_FIELDS.has(field)) continue;
    if (isBlank(row[field])) {
      recorder.add(rowNum, childName, field, `Mandatory field '${field}' is missing or empty.`, "ERROR");
    }
  }
};

export const validateMandatoryParentFields: ValidatorFn = (row, rowNum, childName, recorder, headers) => {
  const blockActive = (indicatorFields: string[]) =>
    indicatorFields.some((f) => headers.has(f) && !isBlank(row[f]));

  if (blockActive(["Parent1_First_Name", "Parent1_Last_Name", "Parent1_Email", "Parent1_CRN"])) {
    for (const field of MANDATORY_PARENT1_FIELDS) {
      if (headers.has(field) && isBlank(row[field])) {
        recorder.add(rowNum, childName, field, `Mandatory Parent 1 field '${field}' is missing or empty.`, "ERROR");
      }
    }
  }

  if (blockActive(["Parent2_First_Name", "Parent2_Last_Name", "Parent2_Email", "Parent2_CRN"])) {
    for (const field of MANDATORY_PARENT2_FIELDS) {
      if (headers.has(field) && isBlank(row[field])) {
        recorder.add(rowNum, childName, field, `Mandatory Parent 2 field '${field}' is missing or empty.`, "ERROR");
      }
    }
  }
};

export const validateEmergencyContactLegacyIds: ValidatorFn = (row, rowNum, childName, recorder, headers) => {
  for (const [legacyIdField, indicatorField] of EMERGENCY_CONTACT_LEGACY_ID_FIELDS) {
    if (!headers.has(indicatorField)) continue;
    if (!isBlank(row[indicatorField])) {
      if (!headers.has(legacyIdField) || isBlank(row[legacyIdField])) {
        recorder.add(
          rowNum,
          childName,
          legacyIdField,
          `Emergency contact '${indicatorField}' is populated but '${legacyIdField}' is missing. A Legacy ID is required.`,
          "ERROR",
        );
      }
    }
  }
};

export const validateDateFields: ValidatorFn = (row, rowNum, childName, recorder, headers) => {
  for (const field of DATE_FIELDS) {
    if (!headers.has(field)) continue;
    const value = (row[field] ?? "").trim();
    if (isBlank(value)) continue;
    if (!isValidDate(value)) {
      recorder.add(
        rowNum,
        childName,
        field,
        `Date field '${field}' contains an invalid value '${value}'. Expected format: YYYY-MM-DD or DD/MM/YYYY.`,
        "ERROR",
      );
    }
  }
};

export const validateCrnFormat: ValidatorFn = (row, rowNum, childName, recorder, headers) => {
  for (const field of CRN_FIELDS) {
    if (!headers.has(field)) continue;
    const value = (row[field] ?? "").trim();
    if (isBlank(value)) continue;
    if (!isValidCrn(value)) {
      recorder.add(
        rowNum,
        childName,
        field,
        `CRN field '${field}' contains an invalid value '${value}'. Expected format: 9 digits + 1 letter (e.g. 123456789A).`,
        "ERROR",
      );
    }
  }
};

export const validatePhoneFields: ValidatorFn = (row, rowNum, childName, recorder, headers) => {
  for (const field of PHONE_FIELDS) {
    if (!headers.has(field)) continue;
    const value = (row[field] ?? "").trim();
    if (isBlank(value)) continue;
    if (!isValidPhone(value)) {
      recorder.add(
        rowNum,
        childName,
        field,
        `Phone field '${field}' does not match a recognised Australian format: '${value}'. Expected: 04xx xxx xxx (mobile) or 0x xxxx xxxx (landline).`,
        "WARNING",
      );
    }
  }
};

export const validateEmailFields: ValidatorFn = (row, rowNum, childName, recorder, headers) => {
  for (const field of EMAIL_FIELDS) {
    if (!headers.has(field)) continue;
    const value = (row[field] ?? "").trim();
    if (isBlank(value)) continue;
    if (!isValidEmail(value)) {
      recorder.add(rowNum, childName, field, `Email field '${field}' contains an invalid address: '${value}'.`, "ERROR");
    }
  }
};

export const validateStatusField: ValidatorFn = (row, rowNum, childName, recorder, headers) => {
  const field = "Status";
  if (!headers.has(field)) return;
  const value = (row[field] ?? "").trim();
  if (isBlank(value)) return;
  if (!VALID_STATUSES.has(value.toLowerCase())) {
    recorder.add(
      rowNum,
      childName,
      field,
      `Status contains an invalid value: '${value}'. Accepted values: Active, Inactive, Waitlist (case-insensitive).`,
      "ERROR",
    );
  }
};

export const validateGenderFields: ValidatorFn = (row, rowNum, childName, recorder, headers) => {
  if (headers.has("Gender")) {
    const value = (row.Gender ?? "").trim();
    if (!isBlank(value) && !VALID_GENDERS.has(value)) {
      recorder.add(
        rowNum,
        childName,
        "Gender",
        `Gender contains an unexpected value: '${value}'. Accepted values: Male, Female.`,
        "WARNING",
      );
    }
  }
  if (headers.has("Gender_Identity")) {
    const value = (row.Gender_Identity ?? "").trim();
    if (!isBlank(value) && !STANDARD_GENDER_IDENTITIES.has(value)) {
      recorder.add(
        rowNum,
        childName,
        "Gender_Identity",
        `Gender Identity contains a non-standard value: '${value}'. Standard values: Male, Female, Non-Binary, Trans Female, Trans Male. Free-text entries are permitted but please verify this is intentional.`,
        "WARNING",
      );
    }
  }
};

export const validateBooleanFields: ValidatorFn = (row, rowNum, childName, recorder, headers) => {
  for (const field of BOOLEAN_FIELDS) {
    if (!headers.has(field)) continue;
    const value = (row[field] ?? "").trim();
    if (isBlank(value)) continue;
    if (!VALID_BOOLEAN_VALUES.has(value.toLowerCase())) {
      recorder.add(
        rowNum,
        childName,
        field,
        `Boolean field '${field}' contains an invalid value: '${value}'. Accepted values: 0, 1, Yes, No, True, False.`,
        "ERROR",
      );
    }
  }
};

export const validateStateFields: ValidatorFn = (row, rowNum, childName, recorder, headers) => {
  for (const field of STATE_FIELDS) {
    if (!headers.has(field)) continue;
    const value = (row[field] ?? "").trim();
    if (isBlank(value)) continue;
    if (!VALID_AU_STATES.has(value.toUpperCase())) {
      recorder.add(
        rowNum,
        childName,
        field,
        `State field '${field}' still contains a non-standard value: '${value}' after attempted normalisation. Expected: NSW, VIC, QLD, SA, WA, TAS, ACT, NT.`,
        "ERROR",
      );
    }
  }
};

export const validatePostcodeFields: ValidatorFn = (row, rowNum, childName, recorder, headers) => {
  for (const field of POSTCODE_FIELDS) {
    if (!headers.has(field)) continue;
    const value = (row[field] ?? "").trim();
    if (isBlank(value)) continue;
    if (!POSTCODE_PATTERN.test(value)) {
      recorder.add(
        rowNum,
        childName,
        field,
        `Postcode field '${field}' contains an invalid value: '${value}'. Expected a 4-digit Australian postcode (e.g. 2000).`,
        "WARNING",
      );
    }
  }
};

export const validateFieldLengths: ValidatorFn = (row, rowNum, childName, recorder, headers) => {
  for (const [field, limit] of Object.entries(FIELD_LENGTH_LIMITS)) {
    if (!headers.has(field)) continue;
    const value = row[field] ?? "";
    if (value && value.length > limit) {
      recorder.add(
        rowNum,
        childName,
        field,
        `Field '${field}' exceeds the maximum of ${limit} characters (current length: ${value.length}).`,
        "ERROR",
      );
    }
  }
};

export const validateServiceId: ValidatorFn = (row, rowNum, childName, recorder, headers) => {
  const field = "ServiceID";
  if (!headers.has(field)) return;
  const value = (row[field] ?? "").trim();
  if (isBlank(value)) return;
  const n = Number(value);
  if (!Number.isInteger(n) || n <= 0 || !/^-?\d+$/.test(value)) {
    recorder.add(
      rowNum,
      childName,
      field,
      `Service ID '${value}' is not a valid positive integer. Please verify the correct Xplor Service ID.`,
      "ERROR",
    );
  }
};

export const validatePairedNameFields: ValidatorFn = (row, rowNum, childName, recorder, headers) => {
  for (const [firstField, lastField] of PAIRED_NAME_FIELDS) {
    if (!headers.has(firstField) || !headers.has(lastField)) continue;
    const first = (row[firstField] ?? "").trim();
    const last = (row[lastField] ?? "").trim();
    if (first && !last) {
      recorder.add(
        rowNum,
        childName,
        lastField,
        `'${firstField}' is populated ('${first}') but '${lastField}' is missing. Both must be provided together.`,
        "ERROR",
      );
    } else if (last && !first) {
      recorder.add(
        rowNum,
        childName,
        firstField,
        `'${lastField}' is populated ('${last}') but '${firstField}' is missing. Both must be provided together.`,
        "ERROR",
      );
    }
  }
};

export const validateWaitlistLogic: ValidatorFn = (row, rowNum, childName, recorder, headers) => {
  if ((row.Status ?? "").trim().toLowerCase() !== "waitlist") return;

  const hasParent = ["Parent1_Email", "Parent1_CRN", "Parent1_Legacy_Account_ID"]
    .filter((f) => headers.has(f))
    .some((f) => !isBlank(row[f]));
  if (!hasParent) {
    recorder.add(
      rowNum,
      childName,
      "Status / Parent1_Email",
      "Child has 'Waitlist' status but no Parent 1 guardian information is provided. Please supply guardian details or confirm whether this record should remain on the waitlist.",
      "WARNING",
    );
  }

  const enrolmentStr = (row.Enrolment_Start_Date ?? "").trim();
  if (!isBlank(enrolmentStr) && isValidDate(enrolmentStr)) {
    if (parseDateValue(enrolmentStr) < today()) {
      recorder.add(
        rowNum,
        childName,
        "Status / Enrolment_Start_Date",
        `Child has 'Waitlist' status but the Enrolment_Start_Date (${enrolmentStr}) has already passed. The status should likely be updated to 'Active' before import.`,
        "ERROR",
      );
    }
  }
};

/** Safety-net check — should rarely trigger since transformCrnChildParentEquality already clears it. */
export const validateCrnChildParentEquality: ValidatorFn = (row, rowNum, childName, recorder, headers) => {
  const childCrn = (row.Child_CRN ?? "").trim();
  if (isBlank(childCrn)) return;
  for (const parentCrnField of ["Parent1_CRN", "Parent2_CRN"]) {
    if (!headers.has(parentCrnField)) continue;
    const parentCrn = (row[parentCrnField] ?? "").trim();
    if (parentCrn && childCrn === parentCrn) {
      recorder.add(
        rowNum,
        childName,
        `Child_CRN / ${parentCrnField}`,
        `Child CRN '${childCrn}' is identical to ${parentCrnField}. A child's CRN and parent's CRN must differ.`,
        "ERROR",
      );
    }
  }
};

export const validateFutureDob: ValidatorFn = (row, rowNum, childName, recorder) => {
  const dobStr = (row.DOB ?? "").trim();
  if (isBlank(dobStr) || !isValidDate(dobStr)) return;
  if (parseDateValue(dobStr) > today()) {
    recorder.add(
      rowNum,
      childName,
      "DOB",
      `Child's date of birth (${dobStr}) is set to a future date. Please verify this is correct.`,
      "ERROR",
    );
  }
};

export const validateEnrolmentDateNotBeforeDob: ValidatorFn = (row, rowNum, childName, recorder) => {
  const dobStr = (row.DOB ?? "").trim();
  const enrolmentStr = (row.Enrolment_Start_Date ?? "").trim();
  if (isBlank(dobStr) || isBlank(enrolmentStr)) return;
  if (!isValidDate(dobStr) || !isValidDate(enrolmentStr)) return;
  if (parseDateValue(enrolmentStr) < parseDateValue(dobStr)) {
    recorder.add(
      rowNum,
      childName,
      "Enrolment_Start_Date",
      `Enrolment_Start_Date (${enrolmentStr}) is earlier than the child's date of birth (${dobStr}).`,
      "ERROR",
    );
  }
};

export const validateMedicareNumber: ValidatorFn = (row, rowNum, childName, recorder, headers) => {
  if (!headers.has("Medicare_Number")) return;
  const value = (row.Medicare_Number ?? "").trim();
  if (!isBlank(value) && !/^\d+$/.test(value)) {
    recorder.add(rowNum, childName, "Medicare_Number", `Medicare Number '${value}' contains non-numeric characters.`, "WARNING");
  }
};

export const validateConsentsPhotosVideos: ValidatorFn = (row, rowNum, childName, recorder, headers) => {
  const field = "Consents_Photos_Videos";
  if (!headers.has(field)) return;
  if (isBlank(row[field])) {
    recorder.add(
      rowNum,
      childName,
      field,
      `'${field}' is blank. Please confirm consent for photos/videos (1 = yes, 0 = no).`,
      "WARNING",
    );
  }
};

/** Runs all 22 per-row validators in the exact order documented in the audit. */
export const ALL_VALIDATORS: ValidatorFn[] = [
  validateMandatoryChildFields,
  validateMandatoryParentFields,
  validateEmergencyContactLegacyIds,
  validateDateFields,
  validateCrnFormat,
  validatePhoneFields,
  validateEmailFields,
  validateStatusField,
  validateGenderFields,
  validateBooleanFields,
  validateStateFields,
  validatePostcodeFields,
  validateFieldLengths,
  validateServiceId,
  validatePairedNameFields,
  validateWaitlistLogic,
  validateCrnChildParentEquality,
  validateFutureDob,
  validateEnrolmentDateNotBeforeDob,
  validateMedicareNumber,
  validateConsentsPhotosVideos,
  // validateEcParentEmailRedundancy (transformEcParentEmailRedundancy in transforms.ts) runs last — wired in pipeline.ts
];
