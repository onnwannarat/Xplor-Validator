"use client";

import { useState } from "react";
import Link from "next/link";
import { AlertCircle, ArrowLeft, CheckCircle2, ChevronDown, Download, Loader2, PlayCircle } from "lucide-react";
import { ServiceMapUploader } from "@/components/ServiceMapUploader";
import { FileUploader } from "@/components/FileUploader";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { useServiceMapStore } from "@/lib/state/serviceMapStore";
import { COLUMN_LABELS, DEFAULT_COLUMNS, ERROR_PRIORITY, type ColumnKey } from "@/lib/tools/paymentPlans/constants";
import { runPaymentPlanChecker, type PaymentPlanRunResult } from "@/lib/tools/paymentPlans/runChecker";
import { buildOutputZip } from "@/lib/validator/reports/zipBundle";
import { downloadBlob } from "@/lib/download";
import { cn } from "@/lib/utils";

export default function PaymentPlansImportPage() {
  const { mapping } = useServiceMapStore();

  const [qkServiceIdsFile, setQkServiceIdsFile] = useState<File[]>([]);
  const [planFile, setPlanFile] = useState<File[]>([]);
  const [columns, setColumns] = useState<Record<ColumnKey, string>>({ ...DEFAULT_COLUMNS });
  const [showColumnMapping, setShowColumnMapping] = useState(false);
  const [isRunning, setIsRunning] = useState(false);
  const [runError, setRunError] = useState<string | null>(null);
  const [result, setResult] = useState<PaymentPlanRunResult | null>(null);

  const hasMapping = !!mapping?.isLoaded;
  const canRun = hasMapping && qkServiceIdsFile.length > 0 && planFile.length > 0 && !isRunning;

  async function handleRun() {
    if (!canRun || !mapping) return;
    setIsRunning(true);
    setRunError(null);
    setResult(null);
    try {
      const res = await runPaymentPlanChecker(planFile[0], qkServiceIdsFile[0], mapping, columns);
      setResult(res);
    } catch (err) {
      setRunError(err instanceof Error ? err.message : "Could not process this file.");
    } finally {
      setIsRunning(false);
    }
  }

  async function handleDownload() {
    if (!result) return;
    const zipBlob = await buildOutputZip([
      ...result.split.files.map((f) => ({ filename: `split_csvs/${f.filename}`, content: f.content })),
      { filename: `${planFile[0]?.name.replace(/\.csv$/i, "")}_error_report.xlsx`, content: result.errorReport },
    ]);
    downloadBlob(zipBlob, "payment_plans_output.zip");
  }

  const totalErrors = result ? Object.values(result.result.errors).reduce((sum, arr) => sum + arr.length, 0) : 0;

  return (
    <div className="mx-auto flex max-w-6xl flex-col gap-6 px-4 py-8 sm:px-6">
      <Link
        href="/"
        className="transition-standard inline-flex w-fit items-center gap-1.5 text-sm font-medium text-muted-foreground hover:text-foreground focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring/50 rounded-md"
      >
        <ArrowLeft className="size-4" aria-hidden />
        All tools
      </Link>

      <div className="rounded-2xl bg-gradient-to-br from-primary via-primary to-accent p-6 text-primary-foreground sm:p-8">
        <h1 className="text-2xl font-semibold sm:text-3xl">💳 Payment Plans Import</h1>
        <p className="mt-1 max-w-2xl opacity-90">
          Validates payment plan CSVs against Onboarding Tool error rules, auto-fixes date formats and weekday
          abbreviations, and splits output by service — entirely in this browser tab.
        </p>
      </div>

      <div className="grid grid-cols-1 gap-6 lg:grid-cols-[300px_1fr]">
        <aside className="flex flex-col gap-4">
          <ServiceMapUploader />
        </aside>

        <div className="flex flex-col gap-6">
          <section className="flex flex-col gap-4 rounded-xl border border-border bg-card p-5">
            <FileUploader
              label="QikKids Service IDs CSV"
              hint="The service ID export from QikKids (columns: Dbid, ServiceId, Name). Used to map service names to Xplor IDs."
              accept=".csv"
              files={qkServiceIdsFile}
              onChange={setQkServiceIdsFile}
              required
            />
            <FileUploader
              label="Payment Plan CSV"
              hint="The raw payment plan export from QikKids."
              accept=".csv"
              files={planFile}
              onChange={setPlanFile}
              required
            />
          </section>

          <section className="rounded-xl border border-border bg-card p-5">
            <button
              type="button"
              onClick={() => setShowColumnMapping((v) => !v)}
              className="flex w-full items-center justify-between text-sm font-semibold text-foreground"
              aria-expanded={showColumnMapping}
            >
              ⚙️ Column mapping (customise if your CSV uses different column names)
              <ChevronDown className={cn("size-4 transition-transform", showColumnMapping && "rotate-180")} />
            </button>
            {showColumnMapping && (
              <div className="mt-4 grid grid-cols-1 gap-3 sm:grid-cols-3">
                {(Object.keys(COLUMN_LABELS) as ColumnKey[]).map((key) => (
                  <div key={key} className="flex flex-col gap-1">
                    <Label htmlFor={`col-${key}`} className="text-xs text-muted-foreground">
                      {COLUMN_LABELS[key]}
                    </Label>
                    <Input
                      id={`col-${key}`}
                      value={columns[key]}
                      onChange={(e) => setColumns((prev) => ({ ...prev, [key]: e.target.value }))}
                      className="h-8 text-sm"
                    />
                  </div>
                ))}
              </div>
            )}
          </section>

          <div className="flex items-center gap-3">
            <Button size="lg" onClick={handleRun} disabled={!canRun} className="gap-2">
              {isRunning ? <Loader2 className="size-4 animate-spin" /> : <PlayCircle className="size-4" />}
              {isRunning ? "Validating…" : "Run checker"}
            </Button>
            {!hasMapping && (
              <p className="text-sm text-muted-foreground">Set a service ID mapping in the sidebar to enable this.</p>
            )}
          </div>

          {runError && (
            <div className="flex items-center gap-2 rounded-xl bg-destructive/10 px-4 py-3 text-sm text-destructive">
              <AlertCircle className="size-4 shrink-0" />
              <span>{runError}</span>
            </div>
          )}

          {result && (
            <section className="flex flex-col gap-5">
              <div className="grid grid-cols-2 gap-3 sm:grid-cols-5">
                {[
                  ["Total rows", result.result.stats.total],
                  ["Date format fixed", result.result.stats.date_fixed],
                  ["Weekday abbreviated", result.result.stats.weekday_fixed],
                  ["Trailing spaces removed", result.result.stats.spaces_fixed],
                  ["Total errors", totalErrors],
                ].map(([label, value]) => (
                  <div key={label as string} className="rounded-xl border border-border bg-card p-4">
                    <p className="text-2xl font-semibold tabular-nums text-foreground">{value}</p>
                    <p className="text-xs text-muted-foreground">{label}</p>
                  </div>
                ))}
              </div>

              <div
                className={cn(
                  "flex items-center gap-2 rounded-xl px-4 py-3 text-sm font-medium",
                  totalErrors === 0 ? "bg-success/10 text-success" : "bg-destructive/10 text-destructive",
                )}
              >
                {totalErrors === 0 ? <CheckCircle2 className="size-4" /> : <AlertCircle className="size-4" />}
                {totalErrors === 0
                  ? "No errors found — this file appears ready for import."
                  : `${totalErrors} error(s) found — review and fix before importing.`}
              </div>

              <div>
                <Button onClick={handleDownload} className="gap-2">
                  <Download className="size-4" />
                  Download outputs (.zip)
                </Button>
                {result.split.files.length === 0 && (
                  <p className="mt-1 text-xs text-warning-foreground">
                    No services could be matched — check that service names in the payment plan CSV match the
                    QikKids Service IDs CSV.
                  </p>
                )}
              </div>

              <div className="overflow-x-auto rounded-xl border border-border">
                <table className="w-full text-sm">
                  <thead className="bg-muted">
                    <tr>
                      <th className="p-2 text-left font-medium">Error category</th>
                      <th className="p-2 text-left font-medium">Count</th>
                    </tr>
                  </thead>
                  <tbody>
                    {ERROR_PRIORITY.map(({ key, label }) => {
                      const n = result.result.errors[key].length;
                      return (
                        <tr key={key} className="border-t border-border">
                          <td className="p-2">{label}</td>
                          <td className={cn("p-2 font-medium tabular-nums", n > 0 ? "text-destructive" : "text-success")}>{n}</td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </section>
          )}
        </div>
      </div>
    </div>
  );
}
