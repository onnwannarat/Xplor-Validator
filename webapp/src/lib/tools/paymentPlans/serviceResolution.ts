import Papa from "papaparse";
import type { ServiceMapping } from "@/lib/validator/serviceMapping";

/** Builds {qkServiceName (lowercase) -> [xplorId, name]} by joining a QikKids Service IDs export (Name, ServiceId) with the serviceIDs.csv mapping. */
export async function buildServiceNameMap(
  qkServiceIdsFile: File,
  serviceMap: ServiceMapping,
): Promise<Map<string, [string, string]>> {
  const text = (await qkServiceIdsFile.text()).replace(/^﻿/, "");
  const result = Papa.parse<Record<string, string>>(text, { header: true, skipEmptyLines: true });

  const nameMap = new Map<string, [string, string]>();
  for (const row of result.data) {
    const qkName = (row["Name"] ?? "").trim();
    const qkId = (row["ServiceId"] ?? "").trim();
    if (!qkName || !qkId) continue;
    const [xplorId, name] = serviceMap.lookupByQk(qkId);
    if (xplorId !== null) {
      nameMap.set(qkName.toLowerCase(), [xplorId, name ?? ""]);
    }
  }
  return nameMap;
}
