"use client";

import { useMemo, useState } from "react";
import Link from "next/link";
import { AlertCircle, ArrowLeft, CheckCircle2, Download, Loader2, PlayCircle } from "lucide-react";
import { ServiceMapUploader } from "@/components/ServiceMapUploader";
import { FileUploader } from "@/components/FileUploader";
import { SummaryMetrics } from "@/components/SummaryMetrics";
import { ResultsTable } from "@/components/ResultsTable";
import { Button } from "@/components/ui/button";
import { useServiceMapStore } from "@/lib/state/serviceMapStore";
import { runValidation } from "@/lib/validator/pipeline";
import { ServiceMapping } from "@/lib/validator/serviceMapping";
import { buildSplitCsvs } from "@/lib/validator/reports/splitCsv";
import { buildClientExcelReport, buildExcelReport } from "@/lib/validator/reports/excelReport";
import { buildDuplicateParentsReport } from "@/lib/validator/reports/duplicateParentsReport";
import { buildOutputZip, downloadBlob } from "@/lib/validator/reports/zipBundle";
import type { ValidationResult } from "@/lib/validator/types";
import { cn } from "@/lib/utils";

type ImportScope = "active" | "activeAndWaitlist";

export default function ParentAndChildImportPage() {
  const { mapping } = useServiceMapStore();
  const activeMapping = mapping?.isLoaded ? mapping : ServiceMapping.empty();

  const [migrationFile, setMigrationFile] = useState<File[]>([]);
  const [scope, setScope] = useState<ImportScope>("active");
  const [existingFiles, setExistingFiles] = useState<File[]>([]);
  const [isRunning, setIsRunning] = useState(false);
  const [isPackaging, setIsPackaging] = useState(false);
  const [runError, setRunError] = useState<string | null>(null);
  const [result, setResult] = useState<ValidationResult | null>(null);
  const [dupCounts, setDupCounts] = useState<{ cross: number; intra: number } | null>(null);

  const canRun = migrationFile.length > 0 && !isRunning;

  const counts = useMemo(() => {
    if (!result) return null;
    return {
      total: result.recorder.issues.length,
      errors: result.recorder.errorCount(),
      warnings: result.recorder.warningCount(),
      fixed: result.recorder.fixedCount(),
    };
  }, [result]);

  async function handleRun() {
    if (migrationFile.length === 0) return;
    setIsRunning(true);
    setRunError(null);
    setResult(null);
    setDupCounts(null);
    try {
      const existingBytes = await Promise.all(
        existingFiles.map(async (f) => ({ bytes: await f.arrayBuffer(), filename: f.name })),
      );
      const res = await runValidation(migrationFile[0], activeMapping, {
        includeWaitlist: scope === "activeAndWaitlist",
        existingFiles: existingBytes,
      });
      setResult(res);
    } catch (err) {
      setRunError(err instanceof Error ? err.message : "Could not process this file.");
    } finally {
      setIsRunning(false);
    }
  }

  async function handleDownload() {
    if (!result) return;
    setIsPackaging(true);
    try {
      const { files: csvFiles, rowNumMap } = buildSplitCsvs(result.rows, result.fieldnames, activeMapping);
      const [auditBuffer, clientBuffer, dupReport] = await Promise.all([
        buildExcelReport(result.recorder, result.rows, activeMapping, rowNumMap),
        buildClientExcelReport(result.recorder, result.rows, activeMapping, rowNumMap),
        buildDuplicateParentsReport(result.recorder, activeMapping, result.rows),
      ]);
      setDupCounts({ cross: dupReport.crossServiceCount, intra: dupReport.intraFileGroupCount });

      const ts = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
      const zipBlob = await buildOutputZip([
        ...csvFiles.map((f) => ({ filename: `split_csvs/${f.filename}`, content: f.content })),
        { filename: `validation_audit_report_${ts}.xlsx`, content: auditBuffer },
        { filename: `client_audit_report_${ts}.xlsx`, content: clientBuffer },
        { filename: `duplicate_parents_report_${ts}.xlsx`, content: dupReport.buffer },
      ]);
      downloadBlob(zipBlob, `parent_child_import_output_${ts}.zip`);
    } finally {
      setIsPackaging(false);
    }
  }

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
        <h1 className="text-2xl font-semibold sm:text-3xl">🔍 Parent and Child Import</h1>
        <p className="mt-1 max-w-2xl opacity-90">
          Upload your migration file — it&apos;s validated and transformed entirely in this browser tab. Your
          original file is never modified, uploaded, or logged.
        </p>
      </div>

      <div className="grid grid-cols-1 gap-6 lg:grid-cols-[300px_1fr]">
        <aside className="flex flex-col gap-4">
          <ServiceMapUploader />
        </aside>

        <div className="flex flex-col gap-6">
          <section className="flex flex-col gap-4 rounded-xl border border-border bg-card p-5">
            <h2 className="text-sm font-semibold text-foreground">Step 1 — Upload your migration file</h2>
            <FileUploader
              label="Migration file (CSV or XLSX)"
              accept=".csv,.xlsx,.xls"
              files={migrationFile}
              onChange={setMigrationFile}
              required
            />
          </section>

          <section className="flex flex-col gap-3 rounded-xl border border-border bg-card p-5">
            <h2 className="text-sm font-semibold text-foreground">Step 2 — Import scope</h2>
            <div className="flex flex-wrap gap-2">
              {(
                [
                  ["active", "Active only"],
                  ["activeAndWaitlist", "Active + Waitlist"],
                ] as [ImportScope, string][]
              ).map(([value, label]) => (
                <button
                  key={value}
                  type="button"
                  onClick={() => setScope(value)}
                  className={cn(
                    "transition-standard rounded-full border px-4 py-1.5 text-sm font-medium focus-visible:outline-none focus-visible:ring-3 focus-visible:ring-ring/50",
                    scope === value
                      ? "border-primary bg-primary text-primary-foreground"
                      : "border-border bg-background text-foreground hover:bg-muted",
                  )}
                  aria-pressed={scope === value}
                >
                  {label}
                </button>
              ))}
            </div>
          </section>

          <section className="flex flex-col gap-4 rounded-xl border border-border bg-card p-5">
            <div>
              <h2 className="text-sm font-semibold text-foreground">Step 3 (optional) — Cross-check against existing service data</h2>
              <p className="text-sm text-muted-foreground">
                Provide CSV/XLSX files from other services already in Xplor. Any parent whose name plus at least one
                of DOB, contact, or email matches will be flagged as a duplicate.
              </p>
            </div>
            <FileUploader
              label="Existing service files"
              accept=".csv,.xlsx,.xls"
              multiple
              files={existingFiles}
              onChange={setExistingFiles}
            />
          </section>

          <div className="flex items-center gap-3">
            <Button size="lg" onClick={handleRun} disabled={!canRun} className="gap-2">
              {isRunning ? <Loader2 className="size-4 animate-spin" /> : <PlayCircle className="size-4" />}
              {isRunning ? "Validating…" : "Run validation"}
            </Button>
            {!activeMapping.isLoaded && (
              <p className="text-sm text-muted-foreground">
                Tip: set a service ID mapping first to enable service-ID resolution and per-service output.
              </p>
            )}
          </div>

          {runError && (
            <div className="flex items-center gap-2 rounded-xl bg-destructive/10 px-4 py-3 text-sm text-destructive">
              <AlertCircle className="size-4 shrink-0" />
              <span>Could not process this file: {runError}</span>
            </div>
          )}

          {result && counts && (
            <section className="flex flex-col gap-5">
              <div
                className={cn(
                  "flex items-center gap-2 rounded-xl px-4 py-3 text-sm font-medium",
                  counts.errors > 0
                    ? "bg-destructive/10 text-destructive"
                    : counts.warnings > 0
                      ? "bg-warning/15 text-warning-foreground"
                      : "bg-success/10 text-success",
                )}
              >
                {counts.errors === 0 && counts.warnings === 0 ? <CheckCircle2 className="size-4" /> : <AlertCircle className="size-4" />}
                {counts.errors > 0
                  ? `${counts.errors} error(s) found — please resolve all errors before importing.`
                  : counts.warnings > 0
                    ? `${counts.warnings} warning(s) found — please review before importing.`
                    : "No errors or warnings — this file appears ready for import."}
              </div>

              <SummaryMetrics
                total={counts.total}
                errors={counts.errors}
                warnings={counts.warnings}
                fixed={counts.fixed}
                intraDupes={dupCounts?.intra}
                crossDupes={dupCounts?.cross}
              />

              <div>
                <Button onClick={handleDownload} disabled={isPackaging} className="gap-2">
                  {isPackaging ? <Loader2 className="size-4 animate-spin" /> : <Download className="size-4" />}
                  {isPackaging ? "Packaging…" : "Download all outputs (.zip)"}
                </Button>
              </div>

              <div>
                <h2 className="mb-3 text-sm font-semibold text-foreground">Validation results</h2>
                <ResultsTable issues={result.recorder.issues} />
              </div>
            </section>
          )}
        </div>
      </div>
    </div>
  );
}
