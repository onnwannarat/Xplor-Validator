"use client";

import { useEffect, useRef } from "react";
import { AlertTriangle, CheckCircle2, Download, Plus, RotateCcw, Trash2, UploadCloud } from "lucide-react";
import {
  MAX_MANUAL_ENTRIES,
  rehydrateServiceMapStore,
  useServiceMapStore,
  type ManualServiceEntry,
} from "@/lib/state/serviceMapStore";
import { SERVICE_MAP_CSV_TEMPLATE } from "@/lib/validator/serviceMapping";
import { downloadTextFile } from "@/lib/download";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import { cn } from "@/lib/utils";

function newEntry(): ManualServiceEntry {
  return { id: crypto.randomUUID(), serviceName: "", qkServiceId: "", xplorServiceId: "" };
}

/**
 * Sidebar-equivalent for the service ID mapping, shared across every tool page
 * within the tab. Entirely optional: up to MAX_MANUAL_ENTRIES services can be
 * typed in directly, or a serviceIDs.csv can be uploaded (required beyond that,
 * but always available even for a single service).
 */
export function ServiceMapUploader() {
  const { source, fileName, manualEntries, mapping, isLoading, error, setSource, setFile, setManualEntries, clear } =
    useServiceMapStore();
  const inputRef = useRef<HTMLInputElement>(null);
  const hydrated = useRef(false);

  useEffect(() => {
    if (!hydrated.current) {
      hydrated.current = true;
      void rehydrateServiceMapStore();
    }
  }, []);

  const rows = manualEntries.length > 0 ? manualEntries : [newEntry()];
  const configuredCount = manualEntries.filter((e) => e.xplorServiceId.trim()).length;

  function updateRow(id: string, patch: Partial<ManualServiceEntry>) {
    const base = manualEntries.length > 0 ? manualEntries : rows;
    setManualEntries(base.map((r) => (r.id === id ? { ...r, ...patch } : r)));
  }

  function addRow() {
    setManualEntries([...rows, newEntry()]);
  }

  function removeRow(id: string) {
    const next = rows.filter((r) => r.id !== id);
    setManualEntries(next.length > 0 ? next : [newEntry()]);
  }

  return (
    <div className="flex flex-col gap-3 rounded-xl border border-border bg-card p-4">
      <div className="flex items-center justify-between">
        <div>
          <h2 className="text-sm font-semibold text-foreground">Service ID mapping</h2>
          <p className="text-xs text-muted-foreground">Optional — improves service-ID resolution and per-service output.</p>
        </div>
        {(fileName || configuredCount > 0) && (
          <Button variant="ghost" size="icon-sm" aria-label="Reset service ID mapping" onClick={clear}>
            <RotateCcw className="size-3.5" />
          </Button>
        )}
      </div>

      <div className="grid grid-cols-2 gap-1 rounded-lg bg-muted p-1 text-sm">
        <button
          type="button"
          onClick={() => setSource("manual")}
          aria-pressed={source === "manual"}
          className={cn(
            "transition-standard rounded-md px-3 py-1.5 font-medium focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring/50",
            source === "manual" ? "bg-card text-foreground shadow-sm" : "text-muted-foreground hover:text-foreground",
          )}
        >
          Enter manually
        </button>
        <button
          type="button"
          onClick={() => setSource("csv")}
          aria-pressed={source === "csv"}
          className={cn(
            "transition-standard rounded-md px-3 py-1.5 font-medium focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring/50",
            source === "csv" ? "bg-card text-foreground shadow-sm" : "text-muted-foreground hover:text-foreground",
          )}
        >
          Upload CSV
        </button>
      </div>

      {source === "manual" ? (
        <div className="flex flex-col gap-2">
          <p className="text-xs text-muted-foreground">
            Up to {MAX_MANUAL_ENTRIES} services. Working with more? Switch to <span className="font-medium">Upload CSV</span> above.
          </p>
          <ul className="flex flex-col gap-2">
            {rows.map((row, idx) => (
              <li key={row.id} className="flex flex-col gap-1.5 rounded-lg border border-border p-2.5">
                <div className="flex items-center justify-between">
                  <span className="text-xs font-medium text-muted-foreground">Service {idx + 1}</span>
                  {rows.length > 1 && (
                    <button
                      type="button"
                      aria-label={`Remove service ${idx + 1}`}
                      onClick={() => removeRow(row.id)}
                      className="transition-standard rounded-md p-1 text-muted-foreground hover:bg-destructive/10 hover:text-destructive focus-visible:ring-2 focus-visible:ring-ring/50"
                    >
                      <Trash2 className="size-3.5" />
                    </button>
                  )}
                </div>
                <Input
                  placeholder="Service Name"
                  value={row.serviceName}
                  onChange={(e) => updateRow(row.id, { serviceName: e.target.value })}
                  className="h-8 text-sm"
                />
                <div className="grid grid-cols-2 gap-1.5">
                  <Input
                    placeholder="QK Service ID"
                    value={row.qkServiceId}
                    onChange={(e) => updateRow(row.id, { qkServiceId: e.target.value })}
                    className="h-8 text-sm"
                  />
                  <Input
                    placeholder="Xplor Service ID"
                    value={row.xplorServiceId}
                    onChange={(e) => updateRow(row.id, { xplorServiceId: e.target.value })}
                    className="h-8 text-sm"
                  />
                </div>
              </li>
            ))}
          </ul>
          {rows.length < MAX_MANUAL_ENTRIES && (
            <Button variant="outline" size="sm" onClick={addRow} className="justify-center gap-1.5">
              <Plus className="size-3.5" />
              Add service
            </Button>
          )}
          {configuredCount > 0 && (
            <div className="flex items-center gap-2 rounded-lg bg-success/10 px-3 py-2 text-sm text-success">
              <CheckCircle2 className="size-4 shrink-0" aria-hidden />
              <span>
                {configuredCount} service{configuredCount === 1 ? "" : "s"} configured
              </span>
            </div>
          )}
        </div>
      ) : (
        <div className="flex flex-col gap-2">
          <p className="text-xs text-muted-foreground">Required columns: Service Name, QKServiceID, Xplor Service ID.</p>
          <button
            type="button"
            onClick={() => downloadTextFile(SERVICE_MAP_CSV_TEMPLATE, "serviceIDs_template.csv")}
            className="transition-standard inline-flex w-fit items-center gap-1.5 text-xs font-medium text-primary hover:underline focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring/50 rounded-sm"
          >
            <Download className="size-3.5" aria-hidden />
            Download CSV template
          </button>
          {fileName ? (
            <div className="flex items-center gap-2 rounded-lg bg-success/10 px-3 py-2 text-sm text-success">
              <CheckCircle2 className="size-4 shrink-0" aria-hidden />
              <span className="truncate">Using: {fileName}</span>
            </div>
          ) : (
            <Button
              variant="outline"
              onClick={() => inputRef.current?.click()}
              disabled={isLoading}
              className="justify-center gap-1.5"
            >
              <UploadCloud className="size-4" />
              {isLoading ? "Loading…" : "Upload serviceIDs.csv"}
            </Button>
          )}
          {error && (
            <div className="flex items-center gap-2 rounded-lg bg-destructive/10 px-3 py-2 text-sm text-destructive">
              <AlertTriangle className="size-4 shrink-0" aria-hidden />
              <span>{error}</span>
            </div>
          )}
          <input
            ref={inputRef}
            type="file"
            accept=".csv"
            className="sr-only"
            onChange={(e) => {
              const file = e.target.files?.[0];
              if (file) void setFile(file);
              e.target.value = "";
            }}
          />
        </div>
      )}

      {!mapping?.isLoaded && (
        <p className="text-xs text-muted-foreground">
          No mapping set yet — you can still run validation, but service IDs won&apos;t be resolved or corrected.
        </p>
      )}
    </div>
  );
}
