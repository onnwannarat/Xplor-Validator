import Papa from "papaparse";
import { loadInputBytes } from "@/lib/validator/parse";
import type { ServiceMapping } from "@/lib/validator/serviceMapping";

export interface NameMismatch {
  qkServiceId: string;
  xplorServiceId: string;
  serviceName: string;
  nameInQk: string;
  possibleMatch: string;
}

function parseCsvBytes(bytes: ArrayBuffer): Record<string, string>[] {
  const text = new TextDecoder("utf-8").decode(bytes).replace(/^﻿/, "");
  const result = Papa.parse<Record<string, string>>(text, { header: true, skipEmptyLines: true });
  return result.data;
}

/** {xplorServiceId: set of fee names}, from an "Xplor Active Fees" export (columns: Service ID, Fee Name). */
export function buildXplorFees(bytes: ArrayBuffer): Map<string, Set<string>> {
  const map = new Map<string, Set<string>>();
  for (const row of parseCsvBytes(bytes)) {
    const sid = (row["Service ID"] ?? "").trim();
    const fee = (row["Fee Name"] ?? "").trim();
    if (sid && fee) {
      if (!map.has(sid)) map.set(sid, new Set());
      map.get(sid)!.add(fee);
    }
  }
  return map;
}

/** {centreName: set of room names}, from an "Xplor Active Rooms" export (columns: Centre_Name, Room_Name). */
export function buildXplorRooms(bytes: ArrayBuffer): Map<string, Set<string>> {
  const map = new Map<string, Set<string>>();
  for (const row of parseCsvBytes(bytes)) {
    const centre = (row["Centre_Name"] ?? "").trim();
    const room = (row["Room_Name"] ?? "").trim();
    if (centre && room) {
      if (!map.has(centre)) map.set(centre, new Set());
      map.get(centre)!.add(room);
    }
  }
  return map;
}

/** {qkServiceId: set of fee/room names actually used}, from one or more QikKids booking exports. */
export async function buildQkUsage(
  files: { bytes: ArrayBuffer; filename: string }[],
): Promise<{ qkFees: Map<string, Set<string>>; qkRooms: Map<string, Set<string>> }> {
  const qkFees = new Map<string, Set<string>>();
  const qkRooms = new Map<string, Set<string>>();

  for (const { bytes, filename } of files) {
    const { rows } = await loadInputBytes(bytes, filename);
    for (const row of rows) {
      const qkId = (row["Service Legacy ID"] ?? "").trim();
      if (!qkId) continue;
      const fee = (row["Fee Name"] ?? "").trim();
      const room = (row["Room Name"] ?? "").trim();
      if (fee) {
        if (!qkFees.has(qkId)) qkFees.set(qkId, new Set());
        qkFees.get(qkId)!.add(fee);
      }
      if (room) {
        if (!qkRooms.has(qkId)) qkRooms.set(qkId, new Set());
        qkRooms.get(qkId)!.add(room);
      }
    }
  }
  return { qkFees, qkRooms };
}

const STOPWORDS = new Set(["", "the", "a", "an"]);

/** Suggests likely matches for an unmatched name by shared significant words. Port of _word_overlap_suggestions. */
export function wordOverlapSuggestions(name: string, candidates: Set<string>): string[] {
  const words = new Set(
    name
      .toLowerCase()
      .split(/\W+/)
      .filter((w) => !STOPWORDS.has(w)),
  );
  const significantWords = [...words].filter((w) => w.length > 2);
  return [...candidates].sort().filter((c) => significantWords.some((w) => c.toLowerCase().includes(w)));
}

function buildMismatches(
  serviceMap: ServiceMapping,
  usageByQkId: Map<string, Set<string>>,
  xplorNamesFor: (xplorId: string, serviceName: string) => Set<string>,
): NameMismatch[] {
  const mismatches: NameMismatch[] = [];
  const qkIds = [...usageByQkId.keys()].sort();

  for (const qkId of qkIds) {
    const [xplorId, serviceName] = serviceMap.lookupByQk(qkId);
    if (xplorId === null) continue;

    const xplorNames = xplorNamesFor(xplorId, serviceName ?? "");
    const namesInQk = [...(usageByQkId.get(qkId) ?? [])].sort();

    for (const name of namesInQk) {
      if (!xplorNames.has(name)) {
        const suggestions = wordOverlapSuggestions(name, xplorNames);
        mismatches.push({
          qkServiceId: qkId,
          xplorServiceId: xplorId,
          serviceName: serviceName ?? "",
          nameInQk: name,
          possibleMatch: suggestions.length > 0 ? suggestions.join("; ") : "(no match found)",
        });
      }
    }
  }
  return mismatches;
}

export function buildFeeMismatches(
  serviceMap: ServiceMapping,
  xplorFees: Map<string, Set<string>>,
  qkFees: Map<string, Set<string>>,
): NameMismatch[] {
  return buildMismatches(serviceMap, qkFees, (xplorId) => xplorFees.get(xplorId) ?? new Set());
}

export function buildRoomMismatches(
  serviceMap: ServiceMapping,
  xplorRooms: Map<string, Set<string>>,
  qkRooms: Map<string, Set<string>>,
): NameMismatch[] {
  // Rooms are keyed by service (centre) name in the Xplor export, not by Xplor Service ID.
  return buildMismatches(serviceMap, qkRooms, (_xplorId, serviceName) => xplorRooms.get(serviceName) ?? new Set());
}
