"use client";

import { useState } from "react";
import Link from "next/link";
import { AlertCircle, ArrowLeft, CheckCircle2, Download, Loader2, PlayCircle } from "lucide-react";
import { ServiceMapUploader } from "@/components/ServiceMapUploader";
import { FileUploader } from "@/components/FileUploader";
import { Button } from "@/components/ui/button";
import { useServiceMapStore } from "@/lib/state/serviceMapStore";
import { buildFeeMismatches, buildQkUsage, buildRoomMismatches, buildXplorFees, buildXplorRooms, type NameMismatch } from "@/lib/tools/roomFeeNames/logic";
import { buildNameMismatchReport } from "@/lib/tools/roomFeeNames/report";
import { downloadBlob } from "@/lib/download";
import { cn } from "@/lib/utils";

export default function RoomFeeNamesPage() {
  const { mapping } = useServiceMapStore();

  const [feesFile, setFeesFile] = useState<File[]>([]);
  const [roomsFile, setRoomsFile] = useState<File[]>([]);
  const [bookingsFiles, setBookingsFiles] = useState<File[]>([]);
  const [isRunning, setIsRunning] = useState(false);
  const [runError, setRunError] = useState<string | null>(null);
  const [result, setResult] = useState<{ report: ArrayBuffer; fees: NameMismatch[]; rooms: NameMismatch[] } | null>(
    null,
  );

  const hasMapping = !!mapping?.isLoaded;
  const canRun = hasMapping && feesFile.length > 0 && roomsFile.length > 0 && bookingsFiles.length > 0 && !isRunning;

  async function handleRun() {
    if (!canRun || !mapping) return;
    setIsRunning(true);
    setRunError(null);
    setResult(null);
    try {
      const xplorFees = buildXplorFees(await feesFile[0].arrayBuffer());
      const xplorRooms = buildXplorRooms(await roomsFile[0].arrayBuffer());
      const bookingsBytes = await Promise.all(
        bookingsFiles.map(async (f) => ({ bytes: await f.arrayBuffer(), filename: f.name })),
      );
      const { qkFees, qkRooms } = await buildQkUsage(bookingsBytes);

      const fees = buildFeeMismatches(mapping, xplorFees, qkFees);
      const rooms = buildRoomMismatches(mapping, xplorRooms, qkRooms);
      const report = await buildNameMismatchReport(fees, rooms);

      setResult({ report, fees, rooms });
    } catch (err) {
      setRunError(err instanceof Error ? err.message : "Could not process these files.");
    } finally {
      setIsRunning(false);
    }
  }

  function handleDownload() {
    if (!result) return;
    downloadBlob(
      new Blob([result.report], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }),
      "Fee_Room_Name_Mismatch_Report.xlsx",
    );
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
        <h1 className="text-2xl font-semibold sm:text-3xl">🏷️ Room & Fee Names checking</h1>
        <p className="mt-1 max-w-2xl opacity-90">
          Compares fee and room names used in QikKids bookings against those defined in Xplor, entirely in this
          browser tab, and produces a mismatch report with suggested matches.
        </p>
      </div>

      <div className="grid grid-cols-1 gap-6 lg:grid-cols-[300px_1fr]">
        <aside className="flex flex-col gap-4">
          <ServiceMapUploader />
        </aside>

        <div className="flex flex-col gap-6">
          <section className="grid grid-cols-1 gap-4 rounded-xl border border-border bg-card p-5 sm:grid-cols-2">
            <FileUploader
              label="Xplor Active Fees CSV"
              hint="Columns required: Service ID, Fee Name"
              accept=".csv"
              files={feesFile}
              onChange={setFeesFile}
              required
            />
            <FileUploader
              label="Xplor Active Rooms CSV"
              hint="Columns required: Centre_Name, Room_Name"
              accept=".csv"
              files={roomsFile}
              onChange={setRoomsFile}
              required
            />
          </section>

          <section className="flex flex-col gap-4 rounded-xl border border-border bg-card p-5">
            <FileUploader
              label="QikKids Bookings (CSV or XLSX, one or more)"
              hint="Recurring and/or casual bookings exports. Columns required: Service Legacy ID, Fee Name, Room Name"
              accept=".csv,.xlsx,.xls"
              multiple
              files={bookingsFiles}
              onChange={setBookingsFiles}
              required
            />
          </section>

          <div className="flex items-center gap-3">
            <Button size="lg" onClick={handleRun} disabled={!canRun} className="gap-2">
              {isRunning ? <Loader2 className="size-4 animate-spin" /> : <PlayCircle className="size-4" />}
              {isRunning ? "Comparing…" : "Run check"}
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
                  <p className="text-2xl font-semibold tabular-nums text-foreground">{result.fees.length}</p>
                  <p className="text-xs text-muted-foreground">Fee mismatches</p>
                </div>
                <div className="rounded-xl border border-border bg-card p-4">
                  <p className="text-2xl font-semibold tabular-nums text-foreground">{result.rooms.length}</p>
                  <p className="text-xs text-muted-foreground">Room mismatches</p>
                </div>
              </div>

              <div
                className={cn(
                  "flex items-center gap-2 rounded-xl px-4 py-3 text-sm font-medium",
                  result.fees.length === 0 && result.rooms.length === 0
                    ? "bg-success/10 text-success"
                    : "bg-warning/15 text-warning-foreground",
                )}
              >
                {result.fees.length === 0 && result.rooms.length === 0 ? (
                  <CheckCircle2 className="size-4" />
                ) : (
                  <AlertCircle className="size-4" />
                )}
                {result.fees.length === 0 && result.rooms.length === 0
                  ? "No mismatches found — all fee and room names match Xplor."
                  : `${result.fees.length} fee mismatch(es) and ${result.rooms.length} room mismatch(es) found. Download the report for details.`}
              </div>

              <div>
                <Button onClick={handleDownload} className="gap-2">
                  <Download className="size-4" />
                  Download mismatch report (.xlsx)
                </Button>
              </div>
            </section>
          )}
        </div>
      </div>
    </div>
  );
}
