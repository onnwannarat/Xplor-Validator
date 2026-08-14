import ExcelJS from "exceljs";
import { DUPLICATE_PARENTS_REPORT_FIELDNAMES } from "../constants";
import { collectIntraFileDupGroups } from "../crossRow";
import type { IssueRecorder } from "../issueRecorder";
import type { ServiceMapping } from "../serviceMapping";
import type { Row, RowEntry } from "../types";

interface DetailRow {
  Service_Name: string;
  Parent_Legacy_ID: string;
  Parent_Name: string;
  Matched_On: string;
  Parent_CRN: string;
  Parent_Slot: string;
  Linked_Child: string;
}

function resolveServiceName(row: Row, serviceMap: ServiceMapping): string {
  const svcId = (row.ServiceID ?? "").trim();
  let svcName = (row.Service_Name ?? "").trim();
  if (!svcName && serviceMap.isLoaded) {
    svcName = serviceMap.getNameByXplor(svcId) || svcId;
  }
  return svcName;
}

/**
 * Builds the duplicate-parents workbook: one row per occurrence, merging intra-file
 * clusters (via Union-Find) with cross-service matches from the recorder.
 * Port of write_duplicate_parents_report. Returns [buffer, crossServiceCount, intraFileGroupCount].
 */
export async function buildDuplicateParentsReport(
  recorder: IssueRecorder,
  serviceMap: ServiceMapping,
  allRows: RowEntry[],
): Promise<{ buffer: ArrayBuffer; crossServiceCount: number; intraFileGroupCount: number }> {
  const groups = collectIntraFileDupGroups(allRows, (row) => resolveServiceName(row, serviceMap));
  const intraCount = groups.length;

  const detailRows: DetailRow[] = [];

  for (const members of groups) {
    const dob = members[0].dob;
    const contactFreq = new Map<string, number>();
    for (const m of members) {
      for (const c of m.contacts) {
        contactFreq.set(c, (contactFreq.get(c) ?? 0) + 1);
      }
    }
    const shared = [...contactFreq.entries()].filter(([, count]) => count >= 2).map(([c]) => c);
    const matchedOn = shared.length > 0 ? `DOB: ${dob}, Contact: ${shared[0]}` : `DOB: ${dob}`;

    const sorted = [...members].sort((a, b) => {
      if (a.service_name !== b.service_name) return a.service_name < b.service_name ? -1 : 1;
      return a.linked_child < b.linked_child ? -1 : a.linked_child > b.linked_child ? 1 : 0;
    });

    for (const m of sorted) {
      detailRows.push({
        Service_Name: m.service_name,
        Parent_Legacy_ID: m.legacy_id,
        Parent_Name: m.parent_name,
        Matched_On: matchedOn,
        Parent_CRN: m.crn,
        Parent_Slot: m.parent_slot,
        Linked_Child: m.linked_child,
      });
    }
  }

  const crossIssues = recorder.issues.filter((i) => i._tag === "cross_service_duplicate_parent");
  const crossCount = crossIssues.length;

  const rowLookup = new Map(allRows.map((e) => [e.rowNum, e]));

  for (const issue of crossIssues) {
    const prefix = String(issue._parent_slot ?? "");
    const entry = rowLookup.get(issue.Row);
    let svcName = "";
    let legacyId = "";
    let crn = "";
    let linked = "";
    if (entry) {
      const { row } = entry;
      svcName = resolveServiceName(row, serviceMap);
      legacyId = (row[`${prefix}_Legacy_Account_ID`] ?? "").trim();
      crn = (row[`${prefix}_CRN`] ?? "").trim();
      linked = `${(row.Child_First_Name ?? "").trim()} ${(row.Child_Last_Name ?? "").trim()}`.trim();
    } else {
      linked = issue.Child_Name;
    }

    detailRows.push({
      Service_Name: svcName,
      Parent_Legacy_ID: legacyId,
      Parent_Name: String(issue._parent_name ?? ""),
      Matched_On: String(issue._matched_on ?? ""),
      Parent_CRN: crn,
      Parent_Slot: prefix,
      Linked_Child: linked,
    });
  }

  detailRows.sort((a, b) => {
    if (a.Parent_Name !== b.Parent_Name) return a.Parent_Name < b.Parent_Name ? -1 : 1;
    if (a.Service_Name !== b.Service_Name) return a.Service_Name < b.Service_Name ? -1 : 1;
    return 0;
  });

  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Duplicate Parents");
  worksheet.addRow([...DUPLICATE_PARENTS_REPORT_FIELDNAMES]);
  for (const row of detailRows) {
    worksheet.addRow(DUPLICATE_PARENTS_REPORT_FIELDNAMES.map((f) => row[f]));
  }

  const headerFont: Partial<ExcelJS.Font> = { bold: true, color: { argb: "FFFFFFFF" }, size: 10 };
  const headerFill: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF4472C4" } };
  for (let c = 1; c <= DUPLICATE_PARENTS_REPORT_FIELDNAMES.length; c++) {
    const cell = worksheet.getRow(1).getCell(c);
    cell.font = headerFont;
    cell.fill = headerFill;
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  }
  worksheet.views = [{ state: "frozen", ySplit: 1 }];

  const dupFill: ExcelJS.Fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF0E6" } };
  const matchedOnColIdx = DUPLICATE_PARENTS_REPORT_FIELDNAMES.indexOf("Matched_On") + 1;
  for (let r = 2; r <= worksheet.rowCount; r++) {
    for (let c = 1; c <= DUPLICATE_PARENTS_REPORT_FIELDNAMES.length; c++) {
      worksheet.getRow(r).getCell(c).fill = dupFill;
    }
    worksheet.getRow(r).getCell(matchedOnColIdx).alignment = { wrapText: true };
  }

  worksheet.columns.forEach((column) => {
    let maxLen = 0;
    column.eachCell?.({ includeEmpty: false }, (cell) => {
      maxLen = Math.max(maxLen, cell.value ? String(cell.value).length : 0);
    });
    column.width = Math.max(12, Math.min(maxLen + 4, 60));
  });

  const buffer = (await workbook.xlsx.writeBuffer()) as unknown as ArrayBuffer;
  return { buffer, crossServiceCount: crossCount, intraFileGroupCount: intraCount };
}
