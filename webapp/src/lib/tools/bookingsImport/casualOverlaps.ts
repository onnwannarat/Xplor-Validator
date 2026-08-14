import { DAY_COLS, WEEKDAY_TO_COL } from "./constants";
import { dayVal, normalizeRoom } from "./helpers";
import { parseDmyStrict, weekdayIndexMondayZero } from "./dateParsing";
import type { TemplateRow } from "./transform";

export interface CasualOverlapResult {
  isOverlap: boolean;
  reason: string;
}

/**
 * Checks every casual booking against all recurring bookings for the same child.
 * A casual booking overlaps when: same child, same service, same room, the casual
 * date falls within the recurring date range, and the casual date's weekday is a
 * booked day in the recurring pattern (Week 1 or 2, depending on Fortnightly cycle
 * position). Port of detect_casual_overlaps.
 */
export function detectCasualOverlaps(casualRows: TemplateRow[], recurringRows: TemplateRow[]): CasualOverlapResult[] {
  const recByChild = new Map<string, TemplateRow[]>();
  for (const row of recurringRows) {
    const key = row.Child_Legacy_Id.trim();
    if (!recByChild.has(key)) recByChild.set(key, []);
    recByChild.get(key)!.push(row);
  }

  return casualRows.map((cas) => {
    const childId = cas.Child_Legacy_Id.trim();
    const relevant = recByChild.get(childId);
    if (!relevant) return { isOverlap: false, reason: "" };

    const casDate = parseDmyStrict(cas.StartDate);
    if (!casDate) return { isOverlap: false, reason: "" };

    const casService = (cas.ServiceID ?? "").trim();
    const casFee = (cas.ImportedFee ?? "").trim();
    const casRoom = (cas.ImportedRoom ?? "").trim();

    for (const rec of relevant) {
      const recStart = parseDmyStrict(rec.StartDate);
      const recEnd = parseDmyStrict(rec.EndDate);
      if (!recStart || !recEnd) continue;

      if (!(recStart <= casDate && casDate <= recEnd)) continue;
      if ((rec.ServiceID ?? "").trim() !== casService) continue;
      if (normalizeRoom(rec.ImportedRoom ?? "") !== normalizeRoom(casRoom)) continue;

      const casWeekday = weekdayIndexMondayZero(casDate); // 0=Mon ... 6=Sun
      const dayPrefix = WEEKDAY_TO_COL[casWeekday];
      const weekType = (rec.WeekType ?? "").trim().toLowerCase();

      let suffix: "1" | "2";
      if (weekType === "fortnightly") {
        const deltaDays = Math.round((casDate.getTime() - recStart.getTime()) / 86400000);
        const weekInCycle = Math.floor(deltaDays / 7) % 2;
        suffix = weekInCycle === 0 ? "1" : "2";
      } else {
        suffix = "1";
      }

      const colToCheck = `${dayPrefix}${suffix}` as (typeof DAY_COLS)[number];
      if (dayVal(rec[colToCheck]) !== 1) continue;

      const dayCols = DAY_COLS.filter((c) => dayVal(rec[c]) === 1);
      const daysStr = dayCols.length > 0 ? dayCols.join(", ") : "—";
      const recFee = (rec.ImportedFee ?? "").trim();
      const period =
        `Casual date ${cas.StartDate} (${dayPrefix}, Week ${suffix} of ${rec.WeekType ?? ""}) falls within ` +
        `recurring period ${rec.StartDate} – ${rec.EndDate} (Days: ${daysStr}, Room: ${casRoom}`;

      const reason =
        recFee === casFee
          ? `${period}, Fee: ${casFee})`
          : `${period}) — same room, different fee: casual fee '${casFee}' vs recurring fee '${recFee}'`;

      return { isOverlap: true, reason };
    }

    return { isOverlap: false, reason: "" };
  });
}
