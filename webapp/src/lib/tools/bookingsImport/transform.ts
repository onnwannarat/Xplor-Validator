import { COLUMN_MAP, TEMPLATE_COLUMNS } from "./constants";
import { formatDmy, parseBookingDate, yearEndDate } from "./dateParsing";
import type { RawRow } from "./parseInput";
import type { ServiceMapping } from "@/lib/validator/serviceMapping";

export type TemplateRow = Record<(typeof TEMPLATE_COLUMNS)[number], string>;

function formatDateField(raw: string): string {
  const parsed = parseBookingDate(raw);
  return parsed ? formatDmy(parsed) : "";
}

/**
 * Transforms de-duplicated raw rows into the template format: maps columns,
 * resolves QK Service IDs to Xplor IDs, formats dates, defaults blank EndDate,
 * and normalises WeekType "single" -> "CASUAL". Port of process_df.
 */
export function transformRows(
  rawRows: RawRow[],
  serviceMap: ServiceMapping,
  defaultEndDate = "",
): { rows: TemplateRow[]; unmappedIds: Set<string> } {
  const unmappedIds = new Set<string>();
  const rows: TemplateRow[] = [];

  for (const raw of rawRows) {
    const qkId = (raw["Service Legacy ID"] ?? "").trim();
    const [xplorIdLookup, serviceNameLookup] = serviceMap.lookupByQk(qkId);
    const mapped = xplorIdLookup !== null ? { xplorId: xplorIdLookup, serviceName: serviceNameLookup ?? "" } : null;
    const xplorId = mapped?.xplorId ?? "";
    if (!mapped) unmappedIds.add(qkId);

    const originalChildId = raw["Child Legacy ID"] ?? "";
    const childLegacyId = xplorId ? `${xplorId}_${originalChildId}` : originalChildId;

    const startDate = formatDateField(raw["Start Date"] ?? "");
    let endDate = formatDateField(raw["End Date"] ?? "");
    if (!endDate) endDate = yearEndDate(startDate, defaultEndDate);

    const freqRaw = (raw["Frequency"] ?? "").trim();
    const weekType = freqRaw.toLowerCase() === "single" ? "CASUAL" : freqRaw;

    const row: Partial<TemplateRow> = {
      ServiceID: xplorId,
      Service_Name: mapped ? mapped.serviceName : raw["Service Name"] ?? "",
      Child_Legacy_Id: childLegacyId,
      Child_First_Name: raw["Child First Name"] ?? "",
      Child_Last_Name: raw["Child Last Name"] ?? "",
      StartDate: startDate,
      EndDate: endDate,
      ImportedFee: raw["Fee Name"] ?? "",
      ImportedRoom: raw["Room Name"] ?? "",
      WeekType: weekType,
      MON1: raw["Monday1"] ?? "",
      TUE1: raw["Tuesday1"] ?? "",
      WED1: raw["Wednesday1"] ?? "",
      THU1: raw["Thursday1"] ?? "",
      FRI1: raw["Friday1"] ?? "",
      SAT1: raw["Saturday1"] ?? "",
      SUN1: raw["Sunday1"] ?? "",
      MON2: raw["Monday2"] ?? "",
      TUE2: raw["Tuesday2"] ?? "",
      WED2: raw["Wednesday2"] ?? "",
      THU2: raw["Thursday2"] ?? "",
      FRI2: raw["Friday2"] ?? "",
      SAT2: raw["Saturday2"] ?? "",
      SUN2: raw["Sunday2"] ?? "",
      QKCreatedDate: formatDateField(raw["Created Date"] ?? ""),
      QKCreatedVia: raw["Created via"] ?? "",
    };

    rows.push(row as TemplateRow);
  }

  return { rows, unmappedIds };
}

export { COLUMN_MAP };
