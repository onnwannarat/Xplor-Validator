import { readBalanceAdjustmentFile } from "./parsers";
import { processBalanceAdjustmentRows, type CentreGroup } from "./process";
import { writeCentreOutput, writeConsolidatedOutput, writeDuplicateReport } from "./writers";

export interface CreatedFile {
  centre: string;
  filename: string;
  buffer: ArrayBuffer;
  rows: number;
}

export interface BalanceAdjustmentsResult {
  totalOutputs: number;
  totalRows: number;
  skippedCentres: Set<string>;
  createdFiles: CreatedFile[];
  errors: string[];
  consolidated: { filename: string; buffer: ArrayBuffer } | null;
  duplicate: { filename: string; buffer: ArrayBuffer } | null;
}

/** Orchestrates the full balance-adjustments run, entirely client-side. Port of main(). */
export async function runBalanceAdjustments(
  files: File[],
  serviceNames: Set<string>,
  templateBytes: ArrayBuffer,
): Promise<BalanceAdjustmentsResult> {
  const skippedCentres = new Set<string>();
  const createdFiles: CreatedFile[] = [];
  const allGroups: CentreGroup[] = [];
  const errors: string[] = [];
  let totalRows = 0;

  for (const file of files) {
    try {
      const rows = await readBalanceAdjustmentFile(file);
      const { createdGroups, skippedCentres: skipped } = processBalanceAdjustmentRows(rows, serviceNames);
      for (const s of skipped) skippedCentres.add(s);

      for (const group of createdGroups) {
        const { filename, buffer } = await writeCentreOutput(group, templateBytes);
        createdFiles.push({ centre: group.centreName, filename, buffer, rows: group.rows.length });
        totalRows += group.rows.length;
        allGroups.push(group);
      }
    } catch (err) {
      errors.push(`${file.name}: ${err instanceof Error ? err.message : String(err)}`);
    }
  }

  let consolidated: BalanceAdjustmentsResult["consolidated"] = null;
  let duplicate: BalanceAdjustmentsResult["duplicate"] = null;
  if (allGroups.length > 0) {
    consolidated = { filename: "All_Services_Balance_Import.xlsx", buffer: await writeConsolidatedOutput(allGroups, templateBytes) };
    const dupBuffer = await writeDuplicateReport(allGroups, templateBytes);
    if (dupBuffer) duplicate = { filename: "Duplicate_Accounts_Report.xlsx", buffer: dupBuffer };
  }

  return {
    totalOutputs: createdFiles.length,
    totalRows,
    skippedCentres,
    createdFiles,
    errors,
    consolidated,
    duplicate,
  };
}
