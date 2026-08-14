import Papa from "papaparse";

/** Minimal CSV template for the service ID mapping — the only 3 columns actually required. */
export const SERVICE_MAP_CSV_TEMPLATE = [
  "Service Name,QKServiceID,Xplor Service ID",
  "Example Early Learning Centre,182,122956",
].join("\r\n");

/**
 * Loads serviceIDs.csv and provides lookups between QK legacy IDs, Xplor Service
 * IDs, and Service Names. Port of validator_v2.py's ServiceMapping class.
 *
 * Only 3 columns are required: Service Name, QKServiceID, Xplor Service ID.
 * "Service Type" and "QKDBID" are also read if present (for backward compatibility
 * with older exports), but are entirely optional.
 */
export class ServiceMapping {
  private qkToXplor = new Map<string, string>();
  private qkToName = new Map<string, string>();
  private xplorToName = new Map<string, string>();
  private nameToXplor = new Map<string, string>(); // lowercase name -> xplor id
  private loaded = false;

  static async fromCsvText(csvText: string): Promise<ServiceMapping> {
    const mapping = new ServiceMapping();
    const text = csvText.charCodeAt(0) === 0xfeff ? csvText.slice(1) : csvText;
    const result = Papa.parse<Record<string, string>>(text, { header: true, skipEmptyLines: true });

    for (const row of result.data) {
      const name = String(row["Service Name"] ?? row["Service_Name"] ?? "").trim();
      const xplor = String(row["Xplor Service ID"] ?? row["Xplor_Service_ID"] ?? "").trim();
      const qk = String(row["QKServiceID"] ?? row["QK_Service_ID"] ?? "").trim();
      const qkdb = String(row["QKDBID"] ?? "").trim();

      if (xplor) mapping.xplorToName.set(xplor, name);
      if (name && xplor) mapping.nameToXplor.set(name.toLowerCase(), xplor);
      if (qk && xplor) {
        mapping.qkToXplor.set(qk, xplor);
        mapping.qkToName.set(qk, name);
      }
      if (qkdb && xplor) {
        mapping.qkToXplor.set(qkdb, xplor);
        mapping.qkToName.set(qkdb, name);
      }
    }
    mapping.loaded = mapping.xplorToName.size > 0;
    return mapping;
  }

  static empty(): ServiceMapping {
    return new ServiceMapping();
  }

  /** Builds a mapping from up to a handful of manually-entered rows (Service Name, QK Service ID, Xplor Service ID). */
  static fromEntries(
    entries: { serviceName: string; qkServiceId: string; xplorServiceId: string }[],
  ): ServiceMapping {
    const mapping = new ServiceMapping();
    for (const entry of entries) {
      const name = entry.serviceName.trim();
      const xplor = entry.xplorServiceId.trim();
      const qk = entry.qkServiceId.trim();
      if (!xplor) continue;

      mapping.xplorToName.set(xplor, name);
      if (name) mapping.nameToXplor.set(name.toLowerCase(), xplor);
      if (qk) {
        mapping.qkToXplor.set(qk, xplor);
        mapping.qkToName.set(qk, name);
      }
    }
    mapping.loaded = mapping.xplorToName.size > 0;
    return mapping;
  }

  get isLoaded(): boolean {
    return this.loaded;
  }

  /** Returns [xplorServiceId, serviceName] for a given QK Service ID, or [null, null]. */
  lookupByQk(qkId: string): [string | null, string | null] {
    const key = qkId.trim();
    return [this.qkToXplor.get(key) ?? null, this.qkToName.get(key) ?? null];
  }

  /** Returns [xplorServiceId, canonicalName] for a service name (case-insensitive), or [null, null]. */
  lookupByName(name: string): [string | null, string | null] {
    const key = name.trim().toLowerCase();
    const xplor = this.nameToXplor.get(key);
    if (xplor) return [xplor, this.xplorToName.get(xplor) ?? ""];
    return [null, null];
  }

  isValidXplorId(xplorId: string): boolean {
    return this.xplorToName.has(xplorId.trim());
  }

  /** Returns the Service Name for a given Xplor Service ID, or the ID itself if unknown. */
  getNameByXplor(xplorId: string): string {
    return this.xplorToName.get(xplorId.trim()) ?? xplorId.trim();
  }

  /** Returns the full set of known Service Names, in their original casing (for exact-match lookups). */
  getAllServiceNames(): Set<string> {
    return new Set(this.xplorToName.values());
  }

  /** Returns a copy of the Xplor Service ID -> Service Name map. */
  getXplorIdToNameMap(): Map<string, string> {
    return new Map(this.xplorToName);
  }
}
