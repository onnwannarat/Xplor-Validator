import {
  checkCrossServiceParentDuplicates,
  checkDuplicateParentEmails,
  checkDuplicates,
  checkEnrolmentParentCrnConsistency,
  checkIntraFileParentDuplicates,
  checkParentCrnEmailRegistry,
  loadExistingParentProfiles,
} from "./crossRow";
import { getChildName, loadInputFile } from "./parse";
import { IssueRecorder } from "./issueRecorder";
import type { ServiceMapping } from "./serviceMapping";
import {
  buildServiceStateFallbacks,
  transformBlankFirstNames,
  transformConsentsPhotos,
  transformCrnChildParentEquality,
  transformEcParentEmailRedundancy,
  transformEmailDedup,
  transformEmergencyContactBooleans,
  transformLegacyIds,
  transformPhoneLeadingZero,
  transformServiceId,
  transformStates,
} from "./transforms";
import { ALL_VALIDATORS } from "./validators";
import type { RowEntry, RunOptions, ValidationResult } from "./types";

/**
 * Runs the full client-side validation pipeline against one uploaded migration file.
 * Mirrors validator_v2.py's run_v2_from_bytes exactly (transform passes → per-row
 * validation → cross-row checks → intra/cross-service duplicate parent checks).
 * Nothing is written to disk — the caller drives report generation from the result.
 */
export async function runValidation(
  file: File,
  serviceMap: ServiceMapping,
  options: RunOptions,
): Promise<ValidationResult> {
  const recorder = new IssueRecorder();

  const { rows: rawRows, fieldnames } = await loadInputFile(file);
  const headers = new Set(fieldnames);

  let allRows: RowEntry[] = rawRows.map((row, idx) => ({
    rowNum: idx + 2, // row 1 is the header
    childName: getChildName(row),
    row,
  }));

  // Import scope filter — drop Waitlist rows when the user selects Active only
  if (!options.includeWaitlist) {
    allRows = allRows.filter((e) => (e.row.Status ?? "").trim().toLowerCase() !== "waitlist");
  }

  // Transformations
  for (const entry of allRows) {
    transformServiceId(entry.row, entry.rowNum, entry.childName, recorder, serviceMap);
  }

  // Filter to only rows whose service mapped to a known Xplor ID
  if (serviceMap.isLoaded) {
    allRows = allRows.filter((e) => serviceMap.isValidXplorId((e.row.ServiceID ?? "").trim()));
  }

  const serviceFallbacks = buildServiceStateFallbacks(allRows);
  for (const entry of allRows) {
    transformStates(entry.row, entry.rowNum, entry.childName, recorder, serviceFallbacks);
  }

  for (const entry of allRows) {
    transformEmailDedup(entry.row, entry.rowNum, entry.childName, recorder, headers);
  }

  for (const entry of allRows) {
    transformCrnChildParentEquality(entry.row, entry.rowNum, entry.childName, recorder, headers);
  }

  for (const entry of allRows) {
    transformBlankFirstNames(entry.row, entry.rowNum, entry.childName, recorder, headers);
  }

  for (const entry of allRows) {
    transformLegacyIds(entry.row, entry.rowNum, entry.childName, recorder, headers);
  }

  for (const entry of allRows) {
    transformPhoneLeadingZero(entry.row, entry.rowNum, entry.childName, recorder, headers);
  }

  for (const entry of allRows) {
    transformConsentsPhotos(entry.row, entry.rowNum, entry.childName, recorder, headers);
  }

  for (const entry of allRows) {
    transformEmergencyContactBooleans(entry.row, entry.rowNum, entry.childName, recorder, headers);
  }

  // Per-row validation, in the exact documented order
  for (const entry of allRows) {
    for (const validator of ALL_VALIDATORS) {
      validator(entry.row, entry.rowNum, entry.childName, recorder, headers);
    }
    // Runs last in the per-row pass, matching validate_ec_parent_email_redundancy's position.
    transformEcParentEmailRedundancy(entry.row, entry.rowNum, entry.childName, recorder, headers);
  }

  // Cross-row validation
  checkDuplicates(allRows, recorder, headers);
  checkDuplicateParentEmails(allRows, recorder, headers);
  checkParentCrnEmailRegistry(allRows, recorder, headers);
  checkEnrolmentParentCrnConsistency(allRows, recorder, headers);

  // Intra-file duplicate parent check (always runs)
  checkIntraFileParentDuplicates(allRows, recorder);

  // Cross-service duplicate parent check (only when existing files provided)
  if (options.existingFiles && options.existingFiles.length > 0) {
    const existingProfiles = await loadExistingParentProfiles(options.existingFiles);
    checkCrossServiceParentDuplicates(allRows, existingProfiles, recorder);
  }

  return { recorder, rows: allRows, fieldnames };
}
