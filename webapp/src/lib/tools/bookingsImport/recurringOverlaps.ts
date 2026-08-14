import { DAY_COLS } from "./constants";
import { dayVal, normalizeRoom } from "./helpers";
import { parseDmyStrict } from "./dateParsing";
import type { TemplateRow } from "./transform";

export interface ConflictRow {
  row: TemplateRow;
  conflictReason: string;
}

export interface RecurringOverlapResult {
  cleanRows: TemplateRow[];
  conflicts: ConflictRow[];
}

/**
 * Finds recurring bookings for the same child+service that overlap in both date
 * range and booked weekdays (same room, same fee). Within each mutually-conflicting
 * group, keeps the row with the furthest EndDate and removes the rest.
 * Port of detect_recurring_schedule_overlaps.
 */
export function detectRecurringScheduleOverlaps(recurringRows: TemplateRow[]): RecurringOverlapResult {
  const groups = new Map<string, number[]>();
  recurringRows.forEach((row, idx) => {
    const key = `${row.Child_Legacy_Id}${row.ServiceID}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key)!.push(idx);
  });

  const removedIndices = new Set<number>();
  const reasonByIndex = new Map<number, string>();

  for (const idxList of groups.values()) {
    if (idxList.length < 2) continue;

    const parsed = new Map<number, { start: Date | null; end: Date | null }>();
    for (const idx of idxList) {
      const row = recurringRows[idx];
      parsed.set(idx, { start: parseDmyStrict(row.StartDate), end: parseDmyStrict(row.EndDate) });
    }

    const adjacency = new Map<number, Set<number>>(idxList.map((i) => [i, new Set<number>()]));
    const pairReason = new Map<string, string>();

    for (let i = 0; i < idxList.length; i++) {
      const idxA = idxList[i];
      const rowA = recurringRows[idxA];
      const { start: startA, end: endA } = parsed.get(idxA)!;
      if (!startA || !endA) continue;

      for (let j = i + 1; j < idxList.length; j++) {
        const idxB = idxList[j];
        const rowB = recurringRows[idxB];
        const { start: startB, end: endB } = parsed.get(idxB)!;
        if (!startB || !endB) continue;

        if (!(startA <= endB && startB <= endA)) continue; // date ranges must overlap
        if (normalizeRoom(rowA.ImportedRoom) !== normalizeRoom(rowB.ImportedRoom)) continue;
        if (rowA.ImportedFee.trim().toLowerCase() !== rowB.ImportedFee.trim().toLowerCase()) continue;

        const sharedDays = DAY_COLS.filter((c) => dayVal(rowA[c]) === 1 && dayVal(rowB[c]) === 1);
        if (sharedDays.length === 0) continue;

        const ovStart = new Date(Math.max(startA.getTime(), startB.getTime()));
        const ovEnd = new Date(Math.min(endA.getTime(), endB.getTime()));
        const fmt = (d: Date) => `${String(d.getUTCDate()).padStart(2, "0")}/${String(d.getUTCMonth() + 1).padStart(2, "0")}/${d.getUTCFullYear()}`;
        const reason = `Overlapping recurring schedules: date overlap ${fmt(ovStart)}–${fmt(ovEnd)}, shared booked days: ${sharedDays.join(", ")}`;

        adjacency.get(idxA)!.add(idxB);
        adjacency.get(idxB)!.add(idxA);
        pairReason.set(`${idxA}|${idxB}`, reason);
      }
    }

    const visited = new Set<number>();
    for (const seed of idxList) {
      if (visited.has(seed) || adjacency.get(seed)!.size === 0) continue;

      const component = new Set<number>();
      const queue = [seed];
      while (queue.length > 0) {
        const cur = queue.pop()!;
        if (component.has(cur)) continue;
        component.add(cur);
        visited.add(cur);
        for (const n of adjacency.get(cur)!) {
          if (!component.has(n)) queue.push(n);
        }
      }
      if (component.size < 2) continue;

      const endKey = (idx: number): number => parsed.get(idx)!.end?.getTime() ?? -Infinity;
      // Furthest EndDate wins; ties keep the first-created (lowest index) row.
      let keepIdx = -1;
      for (const idx of component) {
        if (keepIdx === -1) {
          keepIdx = idx;
          continue;
        }
        if (endKey(idx) > endKey(keepIdx) || (endKey(idx) === endKey(keepIdx) && idx < keepIdx)) {
          keepIdx = idx;
        }
      }
      const kept = parsed.get(keepIdx)!.end;
      const keptEndStr = kept ? `${String(kept.getUTCDate()).padStart(2, "0")}/${String(kept.getUTCMonth() + 1).padStart(2, "0")}/${kept.getUTCFullYear()}` : "";

      for (const idx of component) {
        if (idx === keepIdx) continue;
        const pairKey = [...pairReason.keys()].find((k) => {
          const [a, b] = k.split("|").map(Number);
          return a === idx || b === idx;
        });
        let reason = pairKey ? pairReason.get(pairKey)! : "Overlapping recurring schedule.";
        reason += ` Removed in favour of the row with the furthest EndDate (${keptEndStr}), which was kept in the bookings import.`;
        removedIndices.add(idx);
        reasonByIndex.set(idx, reason);
      }
    }
  }

  const conflicts: ConflictRow[] = [...removedIndices]
    .sort((a, b) => a - b)
    .map((idx) => ({ row: recurringRows[idx], conflictReason: reasonByIndex.get(idx)! }));
  const cleanRows = recurringRows.filter((_, idx) => !removedIndices.has(idx));

  return { cleanRows, conflicts };
}
