"use client";

import { useState } from "react";
import Link from "next/link";
import { AlertCircle, ArrowLeft, CheckCircle2, ChevronDown, Download, Loader2, PlayCircle, UploadCloud } from "lucide-react";
import { ServiceMapUploader } from "@/components/ServiceMapUploader";
import { FileUploader } from "@/components/FileUploader";
import { Button } from "@/components/ui/button";
import { useServiceMapStore } from "@/lib/state/serviceMapStore";
import { runBalanceAdjustments, type BalanceAdjustmentsResult } from "@/lib/tools/balanceAdjustments/run";
import { buildOutputZip } from "@/lib/validator/reports/zipBundle";
import { downloadBlob } from "@/lib/download";
import { cn } from "@/lib/utils";

const BUNDLED_TEMPLATE_URL = "/balance-adjustments-template.xlsx";

export default function BalanceAdjustmentsPage() {
  const { mapping } = useServiceMapStore();

  const [inputFiles, setInputFiles] = useState<File[]>([]);
  const [templateOverride, setTemplateOverride] = useState<File[]>([]);
  const [showTemplateSection, setShowTemplateSection] = useState(false);
  const [isRunning, setIsRunning] = useState(false);
  const [runError, setRunError] = useState<string | null>(null);
  const [result, setResult] = useState<BalanceAdjustmentsResult | null>(null);

  const hasMapping = !!mapping?.isLoaded;
  const canRun = hasMapping && inputFiles.length > 0 && !isRunning;

  async function handleRun() {
    if (!canRun || !mapping) return;
    setIsRunning(true);
    setRunError(null);
    setResult(null);
    try {
      const templateBytes =
        templateOverride.length > 0
          ? await templateOverride[0].arrayBuffer()
          : await fetch(BUNDLED_TEMPLATE_URL).then((r) => r.arrayBuffer());

      const res = await runBalanceAdjustments(inputFiles, mapping.getAllServiceNames(), templateBytes);
      setResult(res);
    } catch (err) {
      setRunError(err instanceof Error ? err.message : "Could not process these files.");
    } finally {
      setIsRunning(false);
    }
  }

  async function handleDownload() {
    if (!result) return;
    // Two input files can produce the same per-centre filename (e.g. the same centre appears in
    // both) — disambiguate so neither silently overwrites the other when the zip is extracted.
    const seenNames = new Map<string, number>();
    const uniqueName = (filename: string) => {
      const count = seenNames.get(filename) ?? 0;
      seenNames.set(filename, count + 1);
      if (count === 0) return filename;
      const dotIdx = filename.lastIndexOf(".");
      return dotIdx === -1 ? `${filename} (${count + 1})` : `${filename.slice(0, dotIdx)} (${count + 1})${filename.slice(dotIdx)}`;
    };

    const entries = [
      ...result.createdFiles.map((f) => ({ filename: `per_centre/${uniqueName(f.filename)}`, content: f.buffer })),
      ...(result.consolidated ? [{ filename: result.consolidated.filename, content: result.consolidated.buffer }] : []),
      ...(result.duplicate ? [{ filename: result.duplicate.filename, content: result.duplicate.buffer }] : []),
    ];
    const zipBlob = await buildOutputZip(entries);
    downloadBlob(zipBlob, "balance_adjustments_output.zip");
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
        <h1 className="text-2xl font-semibold sm:text-3xl">⚖️ Balance Adjustments</h1>
        <p className="mt-1 max-w-2xl opacity-90">
          Maps centre names to Xplor service names, removes demo accounts, and produces one styled XLSX output file
          per centre — entirely in this browser tab.
        </p>
      </div>

      <div className="grid grid-cols-1 gap-6 lg:grid-cols-[300px_1fr]">
        <aside className="flex flex-col gap-4">
          <ServiceMapUploader />
        </aside>

        <div className="flex flex-col gap-6">
          <section className="flex flex-col gap-4 rounded-xl border border-border bg-card p-5">
            <FileUploader
              label="Balance Adjustment Files (CSV or XLSX, one or more)"
              hint="Columns required: Center Name (or Centre Name), Account Name. Credit/Owing or Amount Due columns are used for amounts."
              accept=".csv,.xlsx,.xls"
              multiple
              files={inputFiles}
              onChange={setInputFiles}
              required
            />
          </section>

          <section className="rounded-xl border border-border bg-card p-5">
            <button
              type="button"
              onClick={() => setShowTemplateSection((v) => !v)}
              className="flex w-full items-center justify-between text-sm font-semibold text-foreground"
              aria-expanded={showTemplateSection}
            >
              Advanced — override upload template
              <ChevronDown className={cn("size-4 transition-transform", showTemplateSection && "rotate-180")} />
            </button>
            {!showTemplateSection && (
              <p className="mt-1 flex items-center gap-1.5 text-xs text-muted-foreground">
                <UploadCloud className="size-3.5" />
                Using the bundled Balance Adjustments Details Upload Template.
              </p>
            )}
            {showTemplateSection && (
              <div className="mt-4">
                <FileUploader
                  label="Template XLSX (optional override)"
                  hint="Overrides the bundled Balance Adjustments Details Upload Template.xlsx"
                  accept=".xlsx"
                  files={templateOverride}
                  onChange={setTemplateOverride}
                />
              </div>
            )}
          </section>

          <div className="flex items-center gap-3">
            <Button size="lg" onClick={handleRun} disabled={!canRun} className="gap-2">
              {isRunning ? <Loader2 className="size-4 animate-spin" /> : <PlayCircle className="size-4" />}
              {isRunning ? "Processing…" : "Run balance adjustments"}
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
              <div className="grid grid-cols-2 gap-3 sm:max-w-sm">
                <div className="rounded-xl border border-border bg-card p-4">
                  <p className="text-2xl font-semibold tabular-nums text-foreground">{result.totalOutputs}</p>
                  <p className="text-xs text-muted-foreground">Output files created</p>
                </div>
                <div className="rounded-xl border border-border bg-card p-4">
                  <p className="text-2xl font-semibold tabular-nums text-foreground">{result.totalRows}</p>
                  <p className="text-xs text-muted-foreground">Total data rows written</p>
                </div>
              </div>

              {result.totalOutputs > 0 && (
                <div>
                  <Button onClick={handleDownload} className="gap-2">
                    <Download className="size-4" />
                    Download outputs (.zip)
                  </Button>
                </div>
              )}

              {result.createdFiles.length > 0 && (
                <div>
                  <h2 className="mb-2 text-sm font-semibold text-foreground">Created files</h2>
                  <ul className="flex flex-col gap-1">
                    {[...result.createdFiles]
                      .sort((a, b) => a.centre.localeCompare(b.centre))
                      .map((f, idx) => (
                        <li key={`${f.filename}-${idx}`} className="flex items-center gap-2 text-sm text-muted-foreground">
                          <CheckCircle2 className="size-4 shrink-0 text-success" />
                          <span className="font-medium text-foreground">{f.centre}</span> — {f.filename} ({f.rows} rows)
                        </li>
                      ))}
                  </ul>
                </div>
              )}

              {result.skippedCentres.size > 0 && (
                <div className="flex flex-col gap-2">
                  <div className="flex items-center gap-2 rounded-xl bg-warning/15 px-4 py-3 text-sm font-medium text-warning-foreground">
                    <AlertCircle className="size-4 shrink-0" />
                    {result.skippedCentres.size} centre(s) skipped — name not found in the service ID mapping.
                  </div>
                  <details className="text-sm">
                    <summary className="cursor-pointer text-muted-foreground">
                      Skipped centres ({result.skippedCentres.size})
                    </summary>
                    <ul className="mt-2 flex flex-col gap-1 pl-4 text-muted-foreground">
                      {[...result.skippedCentres].sort().map((c) => (
                        <li key={c}>- {c}</li>
                      ))}
                    </ul>
                  </details>
                  <p className="text-xs text-muted-foreground">
                    Tip: Centre names must match the Service Name exactly (case-sensitive).
                  </p>
                </div>
              )}

              {result.errors.length > 0 && (
                <details className="text-sm">
                  <summary className="cursor-pointer text-destructive">File read errors ({result.errors.length})</summary>
                  <ul className="mt-2 flex flex-col gap-1 pl-4 text-destructive">
                    {result.errors.map((e) => (
                      <li key={e}>{e}</li>
                    ))}
                  </ul>
                </details>
              )}
            </section>
          )}
        </div>
      </div>
    </div>
  );
}
