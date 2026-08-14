import { PARENT_EMAIL_FIELDS_BY_PREFIX, PARENT_PHONE_FIELDS_BY_PREFIX } from "./constants";
import { isBlank, normalisePhoneForMatch } from "./helpers";
import { loadInputBytes } from "./parse";
import type { IssueRecorder } from "./issueRecorder";
import type { ParentProfile, Row, RowEntry } from "./types";

function stripLegacySuffix(legacyId: string): string {
  return legacyId ? legacyId.replace(/_\d+$/, "") : legacyId;
}

/** Best available unique identifier for a parent: CRN > Legacy ID > "first|last" name. */
function getParentIdentityKey(row: Row, prefix: string): string {
  const crn = (row[`${prefix}_CRN`] ?? "").trim();
  if (!isBlank(crn)) return crn.toLowerCase();
  const legacy = (row[`${prefix}_Legacy_Account_ID`] ?? "").trim();
  if (!isBlank(legacy)) return legacy.toLowerCase();
  const first = (row[`${prefix}_First_Name`] ?? "").trim().toLowerCase();
  const last = (row[`${prefix}_Last_Name`] ?? "").trim().toLowerCase();
  if (first || last) return `${first}|${last}`;
  return "";
}

/** Checks Child_Legacy_Id and Child_CRN for uniqueness across all rows. Port of check_duplicates. */
export function checkDuplicates(allRows: RowEntry[], recorder: IssueRecorder, headers: Set<string>): void {
  const fields = ["Child_Legacy_Id", "Child_CRN"] as const;
  const seen: Record<string, Map<string, [number, string][]>> = {
    Child_Legacy_Id: new Map(),
    Child_CRN: new Map(),
  };

  for (const entry of allRows) {
    for (const field of fields) {
      if (!headers.has(field)) continue;
      const value = (entry.row[field] ?? "").trim();
      if (!isBlank(value)) {
        const key = value.toLowerCase();
        const list = seen[field].get(key) ?? [];
        list.push([entry.rowNum, entry.childName]);
        seen[field].set(key, list);
      }
    }
  }

  for (const field of fields) {
    if (!headers.has(field)) continue;
    for (const [value, occurrences] of seen[field]) {
      if (occurrences.length > 1) {
        const rowNumsStr = occurrences.map(([r]) => r).join(", ");
        for (const [rowNum, childName] of occurrences) {
          recorder.add(
            rowNum,
            childName,
            field,
            `Duplicate value in '${field}': '${value}' appears in rows ${rowNumsStr}. Each record must have a unique value.`,
            "ERROR",
          );
        }
      }
    }
  }
}

/**
 * Parent 1 Email: same email + different identity = ERROR.
 * Parent 2 Email: any duplicate = WARNING.
 * Port of check_duplicate_parent_emails.
 */
export function checkDuplicateParentEmails(allRows: RowEntry[], recorder: IssueRecorder, headers: Set<string>): void {
  const prefixes: [string, boolean][] = [
    ["Parent1", true],
    ["Parent2", false],
  ];

  for (const [prefix, errorOnDiffIdentity] of prefixes) {
    const emailField = `${prefix}_Email`;
    if (!headers.has(emailField)) continue;

    const emailToIdentities = new Map<string, Set<string>>();
    const emailToRows = new Map<string, [number, string][]>();

    for (const entry of allRows) {
      const email = (entry.row[emailField] ?? "").trim().toLowerCase();
      if (isBlank(email)) continue;
      const identity = getParentIdentityKey(entry.row, prefix);
      if (!emailToIdentities.has(email)) emailToIdentities.set(email, new Set());
      emailToIdentities.get(email)!.add(identity);
      if (!emailToRows.has(email)) emailToRows.set(email, []);
      emailToRows.get(email)!.push([entry.rowNum, entry.childName]);
    }

    for (const [email, identities] of emailToIdentities) {
      const occurrences = emailToRows.get(email)!;
      if (occurrences.length <= 1) continue;

      if (errorOnDiffIdentity) {
        if (identities.size > 1) {
          const rowNumsStr = occurrences.map(([r]) => r).join(", ");
          for (const [rowNum, childName] of occurrences) {
            recorder.add(
              rowNum,
              childName,
              emailField,
              `Duplicate '${emailField}': '${email}' appears in rows ${rowNumsStr} linked to different parent identities. The same email cannot belong to two different parents.`,
              "ERROR",
            );
          }
        }
      } else {
        const rowNumsStr = occurrences.map(([r]) => r).join(", ");
        for (const [rowNum, childName] of occurrences) {
          recorder.add(
            rowNum,
            childName,
            emailField,
            `Duplicate '${emailField}': '${email}' appears in rows ${rowNumsStr}. Please verify this is intentional.`,
            "WARNING",
          );
        }
      }
    }
  }
}

/** Enforces '1 CRN = 1 Email'. Port of check_parent_crn_email_registry. */
export function checkParentCrnEmailRegistry(allRows: RowEntry[], recorder: IssueRecorder, headers: Set<string>): void {
  for (const prefix of ["Parent1", "Parent2"]) {
    const crnField = `${prefix}_CRN`;
    const emailField = `${prefix}_Email`;
    if (!headers.has(crnField) || !headers.has(emailField)) continue;

    const crnToEmails = new Map<string, Set<string>>();
    const crnToRows = new Map<string, [number, string][]>();

    for (const entry of allRows) {
      const crn = (entry.row[crnField] ?? "").trim().toLowerCase();
      const email = (entry.row[emailField] ?? "").trim().toLowerCase();
      if (isBlank(crn)) continue;
      if (!crnToRows.has(crn)) crnToRows.set(crn, []);
      crnToRows.get(crn)!.push([entry.rowNum, entry.childName]);
      if (!isBlank(email)) {
        if (!crnToEmails.has(crn)) crnToEmails.set(crn, new Set());
        crnToEmails.get(crn)!.add(email);
      }
    }

    for (const [crn, emails] of crnToEmails) {
      if (emails.size <= 1) continue;
      const emailsStr = [...emails].sort().join(", ");
      const rows = crnToRows.get(crn)!;
      const rowNumsStr = rows.map(([r]) => r).join(", ");
      for (const [rowNum, childName] of rows) {
        recorder.add(
          rowNum,
          childName,
          `${crnField} / ${emailField}`,
          `Parent CRN '${crn.toUpperCase()}' is linked to multiple email addresses across rows ${rowNumsStr}: ${emailsStr}. One CRN must map to exactly one email address.`,
          "ERROR",
        );
      }
    }
  }
}

/** Enrolment_Parent_CRN must match Parent1_CRN or Parent2_CRN in the same row. Port of check_enrolment_parent_crn_consistency. */
export function checkEnrolmentParentCrnConsistency(allRows: RowEntry[], recorder: IssueRecorder, headers: Set<string>): void {
  const field = "Enrolment_Parent_CRN";
  if (!headers.has(field)) return;
  for (const entry of allRows) {
    const enrolmentCrn = (entry.row[field] ?? "").trim();
    if (isBlank(enrolmentCrn)) continue;
    const p1 = headers.has("Parent1_CRN") ? (entry.row.Parent1_CRN ?? "").trim() : "";
    const p2 = headers.has("Parent2_CRN") ? (entry.row.Parent2_CRN ?? "").trim() : "";
    if (![p1, p2].filter(Boolean).includes(enrolmentCrn)) {
      recorder.add(
        entry.rowNum,
        entry.childName,
        field,
        `Enrolment_Parent_CRN '${enrolmentCrn}' does not match any listed guardian's CRN (Parent1: '${p1}', Parent2: '${p2}'). The CCS parent must be one of the child's listed guardians.`,
        "ERROR",
      );
    }
  }
}

/** Loads parent profiles from other already-imported service files, for cross-service duplicate detection. */
export async function loadExistingParentProfiles(
  files: { bytes: ArrayBuffer; filename: string }[],
): Promise<ParentProfile[]> {
  const allProfiles: ParentProfile[] = [];
  for (const { bytes, filename } of files) {
    try {
      const { rows } = await loadInputBytes(bytes, filename);
      allProfiles.push(...extractParentProfilesFromRows(rows, filename));
    } catch {
      // Mirrors the Python tool's behavior: skip unreadable reference files rather than aborting the run.
    }
  }
  return allProfiles;
}

function extractParentProfilesFromRows(rows: Row[], sourceFile: string): ParentProfile[] {
  const profiles: ParentProfile[] = [];
  for (const row of rows) {
    for (const prefix of ["Parent1", "Parent2"]) {
      const firstName = (row[`${prefix}_First_Name`] ?? "").trim().toLowerCase();
      const lastName = (row[`${prefix}_Last_Name`] ?? "").trim().toLowerCase();
      if (!firstName && !lastName) continue;

      const dob = (row[`${prefix}_DOB`] ?? "").trim().toLowerCase();

      const contacts = new Set<string>();
      for (const cf of PARENT_PHONE_FIELDS_BY_PREFIX[prefix] ?? []) {
        const v = (row[cf] ?? "").trim();
        if (v) contacts.add(normalisePhoneForMatch(v));
      }

      const emails = new Set<string>();
      for (const ef of PARENT_EMAIL_FIELDS_BY_PREFIX[prefix] ?? []) {
        const v = (row[ef] ?? "").trim().toLowerCase();
        if (v) emails.add(v);
      }

      profiles.push({
        first_name: firstName,
        last_name: lastName,
        dob,
        contacts,
        emails,
        source_file: sourceFile,
        service_id: (row.ServiceID ?? "").trim(),
        legacy_id: (row[`${prefix}_Legacy_Account_ID`] ?? "").trim().toLowerCase(),
        parent_crn: (row[`${prefix}_CRN`] ?? "").trim().toLowerCase(),
      });
    }
  }
  return profiles;
}

/**
 * Flags parents in the input file that match a profile from another already-imported
 * service. Port of check_cross_service_parent_duplicates.
 */
export function checkCrossServiceParentDuplicates(
  allRows: RowEntry[],
  existingProfiles: ParentProfile[],
  recorder: IssueRecorder,
): void {
  if (existingProfiles.length === 0) return;

  for (const entry of allRows) {
    const { row, rowNum, childName } = entry;

    for (const prefix of ["Parent1", "Parent2"]) {
      const firstName = (row[`${prefix}_First_Name`] ?? "").trim().toLowerCase();
      const lastName = (row[`${prefix}_Last_Name`] ?? "").trim().toLowerCase();
      if (!firstName && !lastName) continue;

      const dob = (row[`${prefix}_DOB`] ?? "").trim().toLowerCase();

      const contacts = new Set<string>();
      for (const cf of PARENT_PHONE_FIELDS_BY_PREFIX[prefix] ?? []) {
        const v = (row[cf] ?? "").trim();
        if (v) contacts.add(normalisePhoneForMatch(v));
      }

      const emails = new Set<string>();
      for (const ef of PARENT_EMAIL_FIELDS_BY_PREFIX[prefix] ?? []) {
        const v = (row[ef] ?? "").trim().toLowerCase();
        if (v) emails.add(v);
      }

      const inputLegacyId = (row[`${prefix}_Legacy_Account_ID`] ?? "").trim().toLowerCase();

      for (const profile of existingProfiles) {
        if (profile.first_name !== firstName || profile.last_name !== lastName) continue;

        if (inputLegacyId && profile.legacy_id && inputLegacyId === profile.legacy_id) continue;

        const matchedFields: string[] = [];
        if (dob && profile.dob && dob === profile.dob) {
          matchedFields.push(`DOB: ${(row[`${prefix}_DOB`] ?? "").trim()}`);
        }
        const sharedContacts = [...contacts].filter((c) => profile.contacts.has(c));
        if (sharedContacts.length > 0) matchedFields.push(`Contact: ${sharedContacts[0]}`);
        const sharedEmails = [...emails].filter((e) => profile.emails.has(e));
        if (sharedEmails.length > 0) matchedFields.push(`Email: ${sharedEmails[0]}`);

        if (matchedFields.length === 0) continue;

        const displayName = `${(row[`${prefix}_First_Name`] ?? "").trim()} ${(row[`${prefix}_Last_Name`] ?? "").trim()}`.trim();
        const svcLabel = profile.service_id ? `Service ID ${profile.service_id}` : "unknown service";
        const matchedStr = matchedFields.join(", ");

        recorder.add(
          rowNum,
          childName,
          `${prefix}_First_Name / ${prefix}_Last_Name`,
          `Potential duplicate parent: '${displayName}' matches an existing parent profile in '${profile.source_file}' (${svcLabel}). Matched on: ${matchedStr}. Action required: (1) Link this parent's children to the existing profile in the existing service. (2) Delete the newly created duplicate profile.`,
          "ERROR",
          {
            tag: "cross_service_duplicate_parent",
            meta: {
              parent_slot: prefix,
              parent_name: displayName,
              matched_on: matchedStr,
              duplicate_source: `Existing file: ${profile.source_file}`,
              duplicate_service_id: profile.service_id,
              duplicate_parent_crn: profile.parent_crn,
              duplicate_legacy_id: profile.legacy_id,
            },
          },
        );
        break; // Report once per parent slot — first match is sufficient
      }
    }
  }
}

interface IntraProfile {
  row: Row;
  prefix: string;
  rowNum: number;
  childName: string;
  parentSlot: string;
  display: string;
  firstName: string;
  lastName: string;
  dob: string;
  contacts: Set<string>;
  crn: string;
  legacyId: string;
  serviceId: string;
}

function buildIntraProfiles(allRows: RowEntry[]): IntraProfile[] {
  const profiles: IntraProfile[] = [];
  for (const entry of allRows) {
    const { row, rowNum, childName } = entry;
    const serviceId = (row.ServiceID ?? "").trim();

    for (const prefix of ["Parent1", "Parent2"]) {
      const firstName = (row[`${prefix}_First_Name`] ?? "").trim().toLowerCase();
      const lastName = (row[`${prefix}_Last_Name`] ?? "").trim().toLowerCase();
      if (!firstName && !lastName) continue;

      const dob = (row[`${prefix}_DOB`] ?? "").trim();

      const contacts = new Set<string>();
      for (const cf of PARENT_PHONE_FIELDS_BY_PREFIX[prefix] ?? []) {
        const v = (row[cf] ?? "").trim();
        if (v) contacts.add(normalisePhoneForMatch(v));
      }

      const crn = (row[`${prefix}_CRN`] ?? "").trim().toLowerCase();
      const legacyId = (row[`${prefix}_Legacy_Account_ID`] ?? "").trim().toLowerCase();
      const display = `${(row[`${prefix}_First_Name`] ?? "").trim()} ${(row[`${prefix}_Last_Name`] ?? "").trim()}`.trim();

      profiles.push({
        row,
        prefix,
        rowNum,
        childName,
        parentSlot: prefix,
        display,
        firstName,
        lastName,
        dob,
        contacts,
        crn,
        legacyId,
        serviceId,
      });
    }
  }
  return profiles;
}

/**
 * Detects duplicate parent profiles within the upload file (matching name + DOB +
 * contact but differing legacy IDs) and auto-links the later row to the first-created
 * profile's Legacy ID / CRN. Port of check_intra_file_parent_duplicates.
 */
export function checkIntraFileParentDuplicates(allRows: RowEntry[], recorder: IssueRecorder): void {
  const profiles = buildIntraProfiles(allRows);
  const reportedPairs = new Set<string>();

  for (let i = 0; i < profiles.length; i++) {
    const pa = profiles[i];
    for (let j = i + 1; j < profiles.length; j++) {
      const pb = profiles[j];

      if (pa.firstName !== pb.firstName || pa.lastName !== pb.lastName) continue;
      if (!pa.dob || !pb.dob || pa.dob !== pb.dob) continue;
      const sharedContacts = [...pa.contacts].filter((c) => pb.contacts.has(c));
      if (sharedContacts.length === 0) continue;
      if (pa.legacyId && pb.legacyId && stripLegacySuffix(pa.legacyId) === stripLegacySuffix(pb.legacyId)) continue;

      const pairKey = [
        `${pa.rowNum}|${pa.parentSlot}`,
        `${pb.rowNum}|${pb.parentSlot}`,
      ]
        .sort()
        .join("::");
      if (reportedPairs.has(pairKey)) continue;
      reportedPairs.add(pairKey);

      const sameService = pa.serviceId === pb.serviceId;
      const context = sameService ? `Service ${pa.serviceId}` : `Service ${pa.serviceId} and Service ${pb.serviceId}`;
      const matchedContact = sharedContacts[0];
      const matchedStr = `DOB: ${pa.dob}, Contact: ${matchedContact}`;

      const [first, later] = pa.rowNum <= pb.rowNum ? [pa, pb] : [pb, pa];

      const firstLegacyRaw = (first.row[`${first.prefix}_Legacy_Account_ID`] ?? "").trim();
      const firstCrnRaw = (first.row[`${first.prefix}_CRN`] ?? "").trim();
      const laterLegacyBefore = (later.row[`${later.prefix}_Legacy_Account_ID`] ?? "").trim();
      const laterCrnBefore = (later.row[`${later.prefix}_CRN`] ?? "").trim();

      if (firstLegacyRaw) later.row[`${later.prefix}_Legacy_Account_ID`] = firstLegacyRaw;
      if (firstCrnRaw) later.row[`${later.prefix}_CRN`] = firstCrnRaw;

      const msg =
        `Duplicate parent profile: '${pa.display}' (legacy '${pa.legacyId}') and ` +
        `'${pb.display}' (legacy '${pb.legacyId}') appear to be the same person in ${context}. ` +
        `Matched on: ${matchedStr}. Auto-linked: ${later.childName}'s ${later.parentSlot} profile ` +
        `(legacy '${laterLegacyBefore}', CRN '${laterCrnBefore}') was reassigned to ` +
        `${first.childName}'s ${first.parentSlot} profile (legacy '${firstLegacyRaw}', CRN '${firstCrnRaw}', ` +
        `first created at row ${first.rowNum}) so both children link to the same parent profile post-publishing.`;

      for (const [src, dst] of [
        [pa, pb],
        [pb, pa],
      ] as [IntraProfile, IntraProfile][]) {
        recorder.add(src.rowNum, src.childName, `${src.parentSlot}_Legacy_Account_ID`, msg, "FIXED", {
          action:
            src === later
              ? `Linked to first-created profile at row ${first.rowNum} (legacy '${firstLegacyRaw}', CRN '${firstCrnRaw}')`
              : `Kept as first-created profile; row ${later.rowNum} linked to it`,
          tag: "intra_file_duplicate_parent",
          meta: {
            parent_slot: src.parentSlot,
            parent_name: src.display,
            matched_on: matchedStr,
            duplicate_source: "Input file",
            duplicate_row_num: dst.rowNum,
            duplicate_service_id: dst.serviceId,
            duplicate_parent_crn: dst.crn,
            duplicate_legacy_id: dst.legacyId,
          },
        });
      }
    }
  }
}

export interface DupGroupMember {
  first_name: string;
  last_name: string;
  dob: string;
  contacts: Set<string>;
  legacy_id: string;
  crn: string;
  service_name: string;
  parent_name: string;
  parent_slot: string;
  linked_child: string;
}

/**
 * Union-Find clustering of parent profiles that are the same physical person but
 * carry differing base legacy IDs — feeds the duplicate-parents report.
 * Port of _collect_intra_file_dup_groups.
 */
export function collectIntraFileDupGroups(
  allRows: RowEntry[],
  getServiceName: (row: Row) => string,
): DupGroupMember[][] {
  const profiles: DupGroupMember[] = [];
  for (const entry of allRows) {
    const { row } = entry;
    const svcName = getServiceName(row);
    const childName = `${(row.Child_First_Name ?? "").trim()} ${(row.Child_Last_Name ?? "").trim()}`.trim();

    for (const prefix of ["Parent1", "Parent2"]) {
      const first = (row[`${prefix}_First_Name`] ?? "").trim().toLowerCase();
      const last = (row[`${prefix}_Last_Name`] ?? "").trim().toLowerCase();
      if (!first && !last) continue;

      const dob = (row[`${prefix}_DOB`] ?? "").trim();
      const contacts = new Set<string>();
      for (const cf of PARENT_PHONE_FIELDS_BY_PREFIX[prefix] ?? []) {
        const v = (row[cf] ?? "").trim();
        if (v) contacts.add(normalisePhoneForMatch(v));
      }

      const legacyId = (row[`${prefix}_Legacy_Account_ID`] ?? "").trim();
      const crn = (row[`${prefix}_CRN`] ?? "").trim();
      const parentName = `${(row[`${prefix}_First_Name`] ?? "").trim()} ${(row[`${prefix}_Last_Name`] ?? "").trim()}`.trim();

      profiles.push({
        first_name: first,
        last_name: last,
        dob,
        contacts,
        legacy_id: legacyId,
        crn,
        service_name: svcName,
        parent_name: parentName,
        parent_slot: prefix,
        linked_child: childName,
      });
    }
  }

  const n = profiles.length;
  const uf = Array.from({ length: n }, (_, i) => i);
  const find = (x: number): number => {
    while (uf[x] !== x) {
      uf[x] = uf[uf[x]];
      x = uf[x];
    }
    return x;
  };
  const union = (x: number, y: number) => {
    uf[find(x)] = find(y);
  };

  for (let i = 0; i < n; i++) {
    const pa = profiles[i];
    for (let j = i + 1; j < n; j++) {
      const pb = profiles[j];
      if (pa.first_name !== pb.first_name || pa.last_name !== pb.last_name) continue;
      if (!pa.dob || !pb.dob || pa.dob !== pb.dob) continue;
      if (![...pa.contacts].some((c) => pb.contacts.has(c))) continue;
      if (pa.legacy_id && pb.legacy_id && stripLegacySuffix(pa.legacy_id) === stripLegacySuffix(pb.legacy_id)) continue;
      union(i, j);
    }
  }

  const components = new Map<number, DupGroupMember[]>();
  for (let i = 0; i < n; i++) {
    const root = find(i);
    if (!components.has(root)) components.set(root, []);
    components.get(root)!.push(profiles[i]);
  }

  const result: DupGroupMember[][] = [];
  for (const members of components.values()) {
    if (members.length < 2) continue;
    const uniqueBases = new Set(members.filter((m) => m.legacy_id).map((m) => stripLegacySuffix(m.legacy_id)));
    if (uniqueBases.size < 2) continue;
    result.push(members);
  }
  return result;
}
