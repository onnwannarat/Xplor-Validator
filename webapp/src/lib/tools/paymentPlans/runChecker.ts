import type { ServiceMapping } from "@/lib/validator/serviceMapping";
import { ColumnKey } from "./constants";
import { processCsv, type ProcessResult } from "./processCsv";
import { buildServiceNameMap } from "./serviceResolution";
import { buildPaymentPlanSplitCsvs, translateErrorRows, type SplitCsvResult } from "./splitCsv";
import { buildPaymentPlanErrorReport } from "./errorReport";

export interface PaymentPlanRunResult {
  result: ProcessResult;
  split: SplitCsvResult;
  errorReport: ArrayBuffer;
}

/** Orchestrates the full payment-plan check, entirely client-side. Port of run_payment_plan_checker. */
export async function runPaymentPlanChecker(
  planFile: File,
  qkServiceIdsFile: File,
  serviceMap: ServiceMapping,
  columns: Record<ColumnKey, string>,
): Promise<PaymentPlanRunResult> {
  const result = await processCsv(planFile, columns);
  const serviceNameMap = await buildServiceNameMap(qkServiceIdsFile, serviceMap);
  const split = buildPaymentPlanSplitCsvs(result.processedRows, columns, serviceMap, serviceNameMap);
  translateErrorRows(result.errors, split.rowMap);
  const errorReport = await buildPaymentPlanErrorReport(result.errors, planFile.name);

  return { result, split, errorReport };
}
