import { parseAmount, parseName, type Row } from "./parsers";

export interface CentreGroup {
  centreName: string;
  rows: Row[];
}

export interface ProcessFileResult {
  createdGroups: CentreGroup[];
  skippedCentres: Set<string>;
}

function findCentreColumn(headers: string[]): string | null {
  return headers.find((h) => ["centername", "centrename"].includes(h.toLowerCase().replace(/ /g, ""))) ?? null;
}

/**
 * Groups one input file's rows by centre (excluding demo accounts), keeping only
 * centres that match a known Service Name. Port of _process_input_df.
 */
export function processBalanceAdjustmentRows(rows: Row[], serviceNames: Set<string>): ProcessFileResult {
  if (rows.length === 0) return { createdGroups: [], skippedCentres: new Set() };

  const headers = Object.keys(rows[0]);
  const centreCol = findCentreColumn(headers);
  if (centreCol === null || !headers.includes("Account Name")) {
    return { createdGroups: [], skippedCentres: new Set() };
  }

  const cleanedRows = rows
    .map((row) => ({ ...row, [centreCol]: (row[centreCol] ?? "").trim() }))
    .filter((row) => !(row["Account Name"] ?? "").toLowerCase().includes("demo parent"));

  const seenCentres: string[] = [];
  for (const row of cleanedRows) {
    const centre = row[centreCol];
    if (centre && !seenCentres.includes(centre)) seenCentres.push(centre);
  }

  const createdGroups: CentreGroup[] = [];
  const skippedCentres = new Set<string>();

  for (const centreName of seenCentres) {
    if (!serviceNames.has(centreName)) {
      skippedCentres.add(centreName);
      continue;
    }
    const groupRows = cleanedRows.filter((row) => row[centreCol] === centreName);
    createdGroups.push({ centreName, rows: groupRows });
  }

  return { createdGroups, skippedCentres };
}

export interface ResolvedRow {
  centreName: string;
  firstName: string;
  lastName: string;
  credit: number | null;
  owing: number | null;
}

/** Resolves Credit/Owing per row, from explicit Credit/Owing columns or a signed Amount Due column. */
export function resolveRows(centreName: string, rows: Row[]): ResolvedRow[] {
  const headers = rows.length > 0 ? Object.keys(rows[0]) : [];
  const hasCredit = headers.includes("Credit");
  const hasOwing = headers.includes("Owing");
  const hasAmountDue = headers.includes("Amount Due");

  return rows.map((row) => {
    const [firstName, lastName] = parseName(row["Account Name"]);
    let credit: number | null = null;
    let owing: number | null = null;

    if (hasCredit || hasOwing) {
      credit = parseAmount(row["Credit"]);
      owing = parseAmount(row["Owing"]);
    } else if (hasAmountDue) {
      const amount = parseAmount(row["Amount Due"]);
      if (amount === null) {
        credit = null;
        owing = null;
      } else if (amount < 0) {
        credit = Math.abs(amount);
        owing = null;
      } else {
        owing = amount;
      }
    }

    return { centreName, firstName, lastName, credit, owing };
  });
}
