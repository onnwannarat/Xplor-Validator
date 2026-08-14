import {
  BOOLEAN_FIELDS,
  EC_EMAIL_FIELDS,
  PHONE_FIELDS,
  STATE_FIELDS,
  STATE_NORMALISATION_MAP,
  VALID_AU_STATES,
} from "./constants";
import { isBlank } from "./helpers";
import type { IssueRecorder } from "./issueRecorder";
import type { ServiceMapping } from "./serviceMapping";
import type { Row, RowEntry } from "./types";

/**
 * Resolves ServiceID/Xplor_Service_ID and corrects Service_Name.
 * Port of validator_v2.py's transform_service_id.
 */
export function transformServiceId(
  row: Row,
  rowNum: number,
  childName: string,
  recorder: IssueRecorder,
  serviceMap: ServiceMapping,
): void {
  if (!serviceMap.isLoaded) return;

  const field = "ServiceID";
  const value = (row[field] ?? "").trim();
  if (isBlank(value)) return; // Mandatory field check will catch this

  let resolvedXplorId: string | null = null;
  let canonicalName: string | null = null;

  // Step 1: already a known Xplor Service ID?
  if (serviceMap.isValidXplorId(value)) {
    resolvedXplorId = value;
    canonicalName = serviceMap.getNameByXplor(value);
  }

  // Step 2: QK legacy ID lookup (numeric "182" or full "4258-182")
  if (resolvedXplorId === null) {
    const [xplorId, svcName] = serviceMap.lookupByQk(value);
    if (xplorId !== null) {
      const original = value;
      row[field] = xplorId;
      resolvedXplorId = xplorId;
      canonicalName = svcName;
      recorder.add(
        rowNum,
        childName,
        field,
        `QK Service ID '${original}' mapped to Xplor Service ID '${xplorId}' (${svcName}).`,
        "FIXED",
        { action: `Service ID updated: ${original} → ${xplorId} (${svcName})` },
      );
    }
  }

  // Step 3: name-based lookup
  if (resolvedXplorId === null) {
    const [xplorByName, svcName] = serviceMap.lookupByName(value);
    if (xplorByName !== null) {
      const original = value;
      row[field] = xplorByName;
      resolvedXplorId = xplorByName;
      canonicalName = svcName;
      recorder.add(
        rowNum,
        childName,
        field,
        `Service name '${original}' resolved to Xplor Service ID '${xplorByName}' (${svcName}).`,
        "FIXED",
        { action: `Service ID updated from name: '${original}' → '${xplorByName}' (${svcName})` },
      );
    }
  }

  // Step 4: not found anywhere
  if (resolvedXplorId === null) {
    recorder.add(
      rowNum,
      childName,
      field,
      `Service ID '${value}' was not found in the service mapping file (checked as Xplor Service ID, QK Service ID, and Service Name). Please verify the correct Xplor Service ID.`,
      "WARNING",
    );
    return;
  }

  // Service_Name column correction
  const nameField = "Service_Name";
  if (nameField in row && canonicalName) {
    const currentName = (row[nameField] ?? "").trim();
    if (!isBlank(currentName) && currentName !== canonicalName) {
      row[nameField] = canonicalName;
      recorder.add(
        rowNum,
        childName,
        nameField,
        `Service_Name '${currentName}' does not match the canonical name '${canonicalName}' for Xplor Service ID '${resolvedXplorId}'.`,
        "FIXED",
        { action: `Service_Name corrected: '${currentName}' → '${canonicalName}'` },
      );
    }
  }
}

/** Returns [normalisedValue, wasChanged]. Port of normalise_state_value. */
export function normaliseStateValue(raw: string, fallback = ""): [string, boolean] {
  const stripped = raw.trim();
  if (isBlank(stripped)) return [stripped, false];

  if (VALID_AU_STATES.has(stripped.toUpperCase())) {
    return [stripped.toUpperCase(), stripped.toUpperCase() !== stripped];
  }

  const mapped = STATE_NORMALISATION_MAP[stripped.toLowerCase()];
  if (mapped) return [mapped, true];

  if (fallback) return [fallback, true];

  return [stripped, false];
}

/** Builds the per-service modal valid state, used as a fallback for unrecognisable values. */
export function buildServiceStateFallbacks(allRows: RowEntry[]): Record<string, string> {
  const serviceStateCounts = new Map<string, Map<string, number>>();

  for (const entry of allRows) {
    const svcId = (entry.row.ServiceID ?? "").trim();
    const childState = (entry.row.State ?? "").trim();
    if (!isBlank(svcId) && !isBlank(childState) && VALID_AU_STATES.has(childState.toUpperCase())) {
      const counts = serviceStateCounts.get(svcId) ?? new Map<string, number>();
      counts.set(childState.toUpperCase(), (counts.get(childState.toUpperCase()) ?? 0) + 1);
      serviceStateCounts.set(svcId, counts);
    }
  }

  const fallbacks: Record<string, string> = {};
  for (const [svcId, counts] of serviceStateCounts) {
    let best = "";
    let bestCount = -1;
    for (const [state, count] of counts) {
      if (count > bestCount) {
        best = state;
        bestCount = count;
      }
    }
    if (best) fallbacks[svcId] = best;
  }
  return fallbacks;
}

export function transformStates(
  row: Row,
  rowNum: number,
  childName: string,
  recorder: IssueRecorder,
  serviceFallbacks: Record<string, string>,
): void {
  const svcId = (row.ServiceID ?? "").trim();
  const fallback = serviceFallbacks[svcId] ?? "";

  for (const field of STATE_FIELDS) {
    const original = (row[field] ?? "").trim();
    if (isBlank(original)) continue;

    const [normalised, changed] = normaliseStateValue(original, fallback);
    if (changed) {
      row[field] = normalised;
      let actionDetail = `State normalised: '${original}' → '${normalised}'`;
      if (normalised === fallback && !VALID_AU_STATES.has(original.toUpperCase())) {
        actionDetail += ` (using service modal state as fallback)`;
      }
      recorder.add(
        rowNum,
        childName,
        field,
        `State value '${original}' normalised to '${normalised}'.`,
        "FIXED",
        { action: actionDetail },
      );
    }
  }
}

/**
 * Removes duplicate emails within a row: Parent2 vs Parent1, then EC1-5 priority chain.
 * Port of transform_email_dedup.
 */
export function transformEmailDedup(
  row: Row,
  rowNum: number,
  childName: string,
  recorder: IssueRecorder,
  headers: Set<string>,
): void {
  const p1Email = (row.Parent1_Email ?? "").trim().toLowerCase();
  const p2Field = "Parent2_Email";
  if (headers.has(p2Field)) {
    const p2Email = (row[p2Field] ?? "").trim().toLowerCase();
    if (p2Email && p1Email && p2Email === p1Email) {
      const original = row[p2Field];
      row[p2Field] = "";
      recorder.add(
        rowNum,
        childName,
        p2Field,
        `Parent2_Email was identical to Parent1_Email ('${original}'). Parent2_Email has been cleared to avoid a duplicate account conflict.`,
        "FIXED",
        {
          action: `Duplicate email removed from Parent2_Email (same as Parent1_Email: '${original}')`,
          tag: "duplicate_parent_email",
        },
      );
    }
  }

  const seenEcEmails = new Set<string>();
  for (const field of EC_EMAIL_FIELDS) {
    if (!headers.has(field)) continue;
    const ecEmail = (row[field] ?? "").trim().toLowerCase();
    if (isBlank(ecEmail)) continue;

    if (seenEcEmails.has(ecEmail)) {
      const original = row[field];
      row[field] = "";
      recorder.add(
        rowNum,
        childName,
        field,
        `'${field}' email ('${original}') is a duplicate of an earlier emergency contact's email. The value has been cleared — only the first occurrence is retained.`,
        "FIXED",
        { action: `Duplicate email removed from '${field}': '${original}'` },
      );
    } else {
      seenEcEmails.add(ecEmail);
    }
  }
}

/** Clears Child_CRN if it matches Parent1/Parent2 CRN. Port of transform_crn_child_parent_equality. */
export function transformCrnChildParentEquality(
  row: Row,
  rowNum: number,
  childName: string,
  recorder: IssueRecorder,
  headers: Set<string>,
): void {
  const childCrn = (row.Child_CRN ?? "").trim();
  if (isBlank(childCrn)) return;
  for (const parentCrnField of ["Parent1_CRN", "Parent2_CRN"]) {
    if (!headers.has(parentCrnField)) continue;
    const parentCrn = (row[parentCrnField] ?? "").trim();
    if (parentCrn && childCrn === parentCrn) {
      row.Child_CRN = "";
      recorder.add(
        rowNum,
        childName,
        "Child_CRN",
        `Child_CRN '${childCrn}' was identical to ${parentCrnField}. Child_CRN has been removed — a child and parent cannot share the same CRN.`,
        "FIXED",
        { action: `Child_CRN cleared (was '${childCrn}', same as ${parentCrnField})` },
      );
      return;
    }
  }
}

const NAME_PAIRS: [string, string][] = [
  ["Child_First_Name", "Child_Last_Name"],
  ["Parent1_First_Name", "Parent1_Last_Name"],
  ["Parent2_First_Name", "Parent2_Last_Name"],
  ["EmergencyContact1_First_Name", "EmergencyContact1_Last_Name"],
  ["EmergencyContact2_First_Name", "EmergencyContact2_Last_Name"],
  ["EmergencyContact3_First_Name", "EmergencyContact3_Last_Name"],
  ["EmergencyContact4_First_Name", "EmergencyContact4_Last_Name"],
  ["EmergencyContact5_First_Name", "EmergencyContact5_Last_Name"],
];

/** Copies Last_Name into blank First_Name fields. Port of transform_blank_first_names. */
export function transformBlankFirstNames(
  row: Row,
  rowNum: number,
  childName: string,
  recorder: IssueRecorder,
  headers: Set<string>,
): void {
  for (const [firstField, lastField] of NAME_PAIRS) {
    if (!headers.has(firstField) || !headers.has(lastField)) continue;
    const firstVal = (row[firstField] ?? "").trim();
    const lastVal = (row[lastField] ?? "").trim();
    if (isBlank(firstVal) && !isBlank(lastVal)) {
      row[firstField] = lastVal;
      recorder.add(
        rowNum,
        childName,
        firstField,
        `'${firstField}' was blank while '${lastField}' was '${lastVal}'. First name has been set to the last name value.`,
        "FIXED",
        { action: `Copied '${lastVal}' from ${lastField} into ${firstField}` },
      );
    }
  }
}

/**
 * Prefixes Child_Legacy_Id with the Xplor Service ID; uniquifies Parent2 legacy ID
 * when Parent1 and Parent2 are the same physical person. Port of transform_legacy_ids.
 */
export function transformLegacyIds(
  row: Row,
  rowNum: number,
  childName: string,
  recorder: IssueRecorder,
  headers: Set<string>,
): void {
  const serviceId = (row.ServiceID ?? "").trim();
  if (!serviceId) return;

  const childLegacyField = "Child_Legacy_Id";
  if (headers.has(childLegacyField)) {
    const original = (row[childLegacyField] ?? "").trim();
    if (original && original !== "0") {
      const newVal = `${serviceId}_${original}`;
      row[childLegacyField] = newVal;
      recorder.add(
        rowNum,
        childName,
        childLegacyField,
        `Child Legacy ID prefixed with Xplor Service ID: '${original}' → '${newVal}'.`,
        "FIXED",
        { action: `Child Legacy ID prefixed: '${original}' → '${newVal}'` },
      );
    }
  }

  const p1First = (row.Parent1_First_Name ?? "").trim().toLowerCase();
  const p1Last = (row.Parent1_Last_Name ?? "").trim().toLowerCase();
  const p1Dob = (row.Parent1_DOB ?? "").trim();
  const p1Legacy = (row.Parent1_Legacy_Account_ID ?? "").trim();

  const p2First = (row.Parent2_First_Name ?? "").trim().toLowerCase();
  const p2Last = (row.Parent2_Last_Name ?? "").trim().toLowerCase();
  const p2Dob = (row.Parent2_DOB ?? "").trim();
  const p2Legacy = (row.Parent2_Legacy_Account_ID ?? "").trim();

  if (
    p1First &&
    p2First &&
    p1First === p2First &&
    p1Last === p2Last &&
    p1Dob &&
    p2Dob &&
    p1Dob === p2Dob &&
    p1Legacy &&
    p2Legacy &&
    p1Legacy === p2Legacy
  ) {
    const uniqueP2 = `${p2Legacy}_1`;
    row.Parent2_Legacy_Account_ID = uniqueP2;
    recorder.add(
      rowNum,
      childName,
      "Parent2_Legacy_Account_ID",
      `Parent1 and Parent2 appear to be the same person (name, DOB, and Legacy ID all match: '${p2Legacy}'). Parent2 Legacy ID made unique: '${p2Legacy}' → '${uniqueP2}'.`,
      "FIXED",
      { action: `Parent2 Legacy ID uniquified: '${p2Legacy}' → '${uniqueP2}'` },
    );
  }
}

/** Prepends a leading '0' to bare-digit AU phone numbers. Port of transform_phone_leading_zero. */
export function transformPhoneLeadingZero(
  row: Row,
  rowNum: number,
  childName: string,
  recorder: IssueRecorder,
  headers: Set<string>,
): void {
  for (const field of PHONE_FIELDS) {
    if (!headers.has(field)) continue;
    const original = (row[field] ?? "").trim();
    if (isBlank(original)) continue;
    if (original.startsWith("+")) continue;
    const stripped = original.replace(/[\s\-()]/g, "");
    if (/^\d+$/.test(stripped) && !stripped.startsWith("0")) {
      const newVal = "0" + original;
      row[field] = newVal;
      recorder.add(
        rowNum,
        childName,
        field,
        `Phone number '${original}' was missing a leading zero. Prepended '0': '${newVal}'.`,
        "FIXED",
        { action: `Leading zero added: '${original}' → '${newVal}'` },
      );
    }
  }
}

/** Defaults blank photo-consent fields to 'N'. Port of transform_consents_photos. */
export function transformConsentsPhotos(
  row: Row,
  rowNum: number,
  childName: string,
  recorder: IssueRecorder,
  headers: Set<string>,
): void {
  for (const field of ["Consents_Photos", "Consents_Photos_Videos"]) {
    if (!headers.has(field)) continue;
    if (isBlank(row[field] ?? "")) {
      row[field] = "N";
      recorder.add(
        rowNum,
        childName,
        field,
        `'${field}' was blank — defaulted to 'N' (no photo consent).`,
        "FIXED",
        { action: `'${field}' set to 'N' (was blank)` },
      );
    }
  }
}

/** Converts True/False (case-insensitive) EC boolean fields to 1/0. Port of transform_emergency_contact_booleans. */
export function transformEmergencyContactBooleans(
  row: Row,
  rowNum: number,
  childName: string,
  recorder: IssueRecorder,
  headers: Set<string>,
): void {
  for (const field of BOOLEAN_FIELDS) {
    if (!field.includes("EmergencyContact")) continue;
    if (!headers.has(field)) continue;
    const value = (row[field] ?? "").trim();
    let newValue: string | null = null;
    if (value.toLowerCase() === "true") newValue = "1";
    else if (value.toLowerCase() === "false") newValue = "0";
    else continue;

    row[field] = newValue;
    recorder.add(
      rowNum,
      childName,
      field,
      `'${field}' value '${value}' converted to '${newValue}'.`,
      "FIXED",
      { action: `'${field}' set to '${newValue}' (was '${value}')` },
    );
  }
}

/**
 * If any Emergency Contact email matches Parent 1/2's email, clears the EC email.
 * Port of validate_ec_parent_email_redundancy (semantically a transform, but the
 * Python source keeps it in the validator pass — see validators.ts).
 */
export function transformEcParentEmailRedundancy(
  row: Row,
  rowNum: number,
  childName: string,
  recorder: IssueRecorder,
  headers: Set<string>,
): void {
  const p1Email = (row.Parent1_Email ?? "").trim().toLowerCase();
  const p2Email = (row.Parent2_Email ?? "").trim().toLowerCase();
  const parentEmails = new Map<string, string>();
  if (p1Email) parentEmails.set(p1Email, "Parent1_Email");
  if (p2Email) parentEmails.set(p2Email, "Parent2_Email");
  if (parentEmails.size === 0) return;

  for (const ecField of EC_EMAIL_FIELDS) {
    if (!headers.has(ecField)) continue;
    const ecEmail = (row[ecField] ?? "").trim().toLowerCase();
    if (isBlank(ecEmail)) continue;
    const matchedParent = parentEmails.get(ecEmail);
    if (matchedParent) {
      const original = (row[ecField] ?? "").trim();
      row[ecField] = "";
      recorder.add(
        rowNum,
        childName,
        ecField,
        `Redundant Contact: '${ecField}' ('${original}') is identical to ${matchedParent}. A parent must not also be listed as an Emergency Contact — the email has been cleared.`,
        "FIXED",
        {
          action: `Duplicate email removed from '${ecField}': '${original}' (same as ${matchedParent})`,
          tag: "redundant_ec",
        },
      );
    }
  }
}
