"use client";

import { useState } from "react";
import Link from "next/link";
import { AlertCircle, AlertTriangle, ArrowLeft, Download, Loader2, PlayCircle } from "lucide-react";
import { ServiceMapUploader } from "@/components/ServiceMapUploader";
import { FileUploader } from "@/components/FileUploader";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { Label } from "@/components/ui/label";
import { useServiceMapStore } from "@/lib/state/serviceMapStore";
import { runBookingsImport, type BookingsImportResult } from "@/lib/tools/bookingsImport/run";
import { buildOutputZip } from "@/lib/validator/reports/zipBundle";
import { downloadBlob } from "@/lib/download";

function isoToDmy(iso: string): string {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return `${d}/${m}/${y}`;
}

export default function BookingsImportPage() {
  const { mapping } = useServiceMapStore();

  const [bookingsFiles, setBookingsFiles] = useState<File[]>([]);
  const [defaultEndDateIso, setDefaultEndDateIso] = useState("");
  const [isRunning, setIsRunning] = useState(false);
  const [runError, setRunError] = useState<string | null>(null);
  const [result, setResult] = useState<BookingsImportResult | null>(null);

  const hasMapping = !!mapping?.isLoaded;
  const canRun = hasMapping && bookingsFiles.length > 0 && !isRunning;

  async function handleRun() {
    if (!canRun || !mapping) return;
    setIsRunning(true);
    setRunError(null);
    setResult(null);
    try {
      const res = await runBookingsImport(bookingsFiles, mapping, isoToDmy(defaultEndDateIso));
      setResult(res);
    } catch (err) {
      setRunError(err instanceof Error ? err.message : "Could not process these files.");
    } finally {
      setIsRunning(false);
    }
  }

  async function handleDownload() {
    if (!result) return;
    const entries = [
      ...result.recurringFiles.map((f) => ({ filename: `Recurring/${f.filename}`, content: f.content })),
      ...result.casualFiles.map((f) => ({ filename: `Casual/${f.filename}`, content: f.content })),
      { filename: result.duplicateReportFilename, content: result.duplicateReportBuffer },
      ...(result.removedOverlapReportBuffer && result.removedOverlapReportFilename
        ? [{ filename: result.removedOverlapReportFilename, content: result.removedOverlapReportBuffer }]
        : []),
    ];
    const zipBlob = await buildOutputZip(entries);
    downloadBlob(zipBlob, "bookings_import_output.zip");
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
        <h1 className="text-2xl font-semibold sm:text-3xl">📅 Bookings Import</h1>
        <p className="mt-1 max-w-2xl opacity-90">
          Processes QikKids booking exports: removes duplicates and schedule overlaps, maps service IDs, and splits
          output CSVs per service — entirely in this browser tab.
        </p>
      </div>

      <div className="grid grid-cols-1 gap-6 lg:grid-cols-[300px_1fr]">
        <aside className="flex flex-col gap-4">
          <ServiceMapUploader />
        </aside>

        <div className="flex flex-col gap-6">
          <section className="flex flex-col gap-4 rounded-xl border border-border bg-card p-5">
            <FileUploader
              label="QikKids Booking Files (CSV or XLSX, one or more)"
              hint="Recurring and casual files can be mixed. Required columns include Service Legacy ID, Child Legacy ID, Fee Name, Room Name, Frequency."
              accept=".csv,.xlsx,.xls"
              multiple
              files={bookingsFiles}
              onChange={setBookingsFiles}
              required
            />
          </section>

          <section className="flex flex-col gap-2 rounded-xl border border-border bg-card p-5">
            <Label htmlFor="default-end-date" className="text-sm font-semibold text-foreground">
              Default booking end date (optional)
            </Label>
            <p className="text-sm text-muted-foreground">
              Applied to any booking with a blank End Date. Leave empty to fall back to 31/12 of that booking&apos;s
              own Start Date year.
            </p>
            <Input
              id="default-end-date"
              type="date"
              value={defaultEndDateIso}
              onChange={(e) => setDefaultEndDateIso(e.target.value)}
              className="w-48"
            />
          </section>

          <div className="flex items-center gap-3">
            <Button size="lg" onClick={handleRun} disabled={!canRun} className="gap-2">
              {isRunning ? <Loader2 className="size-4 animate-spin" /> : <PlayCircle className="size-4" />}
              {isRunning ? "Processing…" : "Run bookings import"}
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
                  ["Input files", result.nInputFiles],
                  ["Raw rows", result.nRawRows],
                  ["Recurring bookings", result.nRecurring],
                  ["Casual bookings", result.nCasual],
                  ["Duplicates removed", result.nDupeRows],
                ].map(([label, value]) => (
                  <div key={label as string} className="rounded-xl border border-border bg-card p-4">
                    <p className="text-2xl font-semibold tabular-nums text-foreground">{value}</p>
                    <p className="text-xs text-muted-foreground">{label}</p>
                  </div>
                ))}
              </div>

              {result.nSchedConflictRows > 0 && (
                <div className="flex items-center gap-2 rounded-xl bg-warning/15 px-4 py-3 text-sm text-warning-foreground">
                  <AlertTriangle className="size-4 shrink-0" />
                  {result.nSchedConflictRows} recurring booking(s) removed due to schedule conflicts across{" "}
                  {result.nSchedConflictGroups} group(s). See the duplicate report&apos;s &quot;Recurring Schedule
                  Overlaps&quot; sheet.
                </div>
              )}

              {result.nCasualRemoved > 0 && (
                <div className="flex items-center gap-2 rounded-xl bg-warning/15 px-4 py-3 text-sm text-warning-foreground">
                  <AlertTriangle className="size-4 shrink-0" />
                  {result.nCasualRemoved} casual booking(s) removed due to overlap with recurring bookings. See the
                  removed-overlap report.
                </div>
              )}

              {result.unmappedIds.size > 0 && (
                <div className="flex flex-col gap-2">
                  <div className="flex items-center gap-2 rounded-xl bg-destructive/10 px-4 py-3 text-sm font-medium text-destructive">
                    <AlertCircle className="size-4 shrink-0" />
                    {result.unmappedIds.size} QK Service ID(s) could not be mapped — those rows have a blank
                    ServiceID.
                  </div>
                  <details className="text-sm">
                    <summary className="cursor-pointer text-muted-foreground">
                      Unmapped Service IDs ({result.unmappedIds.size})
                    </summary>
                    <ul className="mt-2 flex flex-col gap-1 pl-4 text-muted-foreground">
                      {[...result.unmappedIds].sort().map((id) => (
                        <li key={id}>- {id}</li>
                      ))}
                    </ul>
                  </details>
                </div>
              )}

              <div>
                <Button onClick={handleDownload} className="gap-2">
                  <Download className="size-4" />
                  Download outputs (.zip)
                </Button>
                <p className="mt-1 text-xs text-muted-foreground">Upload everything in Recurring/ first, then Casual/.</p>
              </div>

              <div className="grid grid-cols-1 gap-4 sm:grid-cols-2">
                <div>
                  <h2 className="mb-2 text-sm font-semibold text-foreground">Recurring files ({result.recurringFiles.length})</h2>
                  {result.recurringFiles.length > 0 ? (
                    <ul className="flex flex-col gap-1 text-sm text-muted-foreground">
                      {[...result.recurringFiles].sort((a, b) => a.filename.localeCompare(b.filename)).map((f) => (
                        <li key={f.filename}>
                          <span className="font-medium text-foreground">{f.filename}</span> — {f.rows} rows
                        </li>
                      ))}
                    </ul>
                  ) : (
                    <p className="text-sm text-muted-foreground">No recurring files produced.</p>
                  )}
                </div>
                <div>
                  <h2 className="mb-2 text-sm font-semibold text-foreground">Casual files ({result.casualFiles.length})</h2>
                  {result.casualFiles.length > 0 ? (
                    <ul className="flex flex-col gap-1 text-sm text-muted-foreground">
                      {[...result.casualFiles].sort((a, b) => a.filename.localeCompare(b.filename)).map((f) => (
                        <li key={f.filename}>
                          <span className="font-medium text-foreground">{f.filename}</span> — {f.rows} rows
                        </li>
                      ))}
                    </ul>
                  ) : (
                    <p className="text-sm text-muted-foreground">No casual files produced.</p>
                  )}
                </div>
              </div>
            </section>
          )}
        </div>
      </div>
    </div>
  );
}
