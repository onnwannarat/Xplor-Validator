import { create } from "zustand";
import { ServiceMapping } from "../validator/serviceMapping";

export interface ManualServiceEntry {
  id: string;
  serviceName: string;
  qkServiceId: string;
  xplorServiceId: string;
}

export const MAX_MANUAL_ENTRIES = 5;

type Source = "manual" | "csv";

interface ServiceMapState {
  source: Source;
  fileName: string | null;
  csvText: string | null;
  manualEntries: ManualServiceEntry[];
  mapping: ServiceMapping | null;
  isLoading: boolean;
  error: string | null;
  setSource: (source: Source) => void;
  setFile: (file: File) => Promise<void>;
  setManualEntries: (entries: ManualServiceEntry[]) => void;
  clear: () => void;
}

const SESSION_KEY = "xplor-validator:service-map";

function buildManualMapping(entries: ManualServiceEntry[]): ServiceMapping {
  return ServiceMapping.fromEntries(
    entries.filter((e) => e.xplorServiceId.trim()).map((e) => ({
      serviceName: e.serviceName,
      qkServiceId: e.qkServiceId,
      xplorServiceId: e.xplorServiceId,
    })),
  );
}

function persist(state: Pick<ServiceMapState, "source" | "fileName" | "csvText" | "manualEntries">) {
  sessionStorage.setItem(SESSION_KEY, JSON.stringify(state));
}

/**
 * Holds the service ID mapping for the current browser tab, mirroring
 * shared/service_map.py's Streamlit session-state pattern: set once, available
 * to every tool page without re-entering. Backed by sessionStorage (not localStorage
 * or any server) so it never survives closing the tab and never leaves the browser.
 *
 * The mapping is entirely optional and can come from either source:
 *  - up to MAX_MANUAL_ENTRIES rows typed directly into the UI, or
 *  - an uploaded serviceIDs.csv (required once there are more than MAX_MANUAL_ENTRIES services).
 *
 * Holds only the service ID *mapping* (centre names + IDs) — never enrolment PII.
 */
export const useServiceMapStore = create<ServiceMapState>((set, get) => ({
  source: "manual",
  fileName: null,
  csvText: null,
  manualEntries: [],
  mapping: null,
  isLoading: false,
  error: null,

  setSource: (source) => {
    set({ source, error: null });
    const { fileName, csvText, manualEntries } = get();
    persist({ source, fileName, csvText, manualEntries });
  },

  setFile: async (file: File) => {
    set({ isLoading: true, error: null });
    try {
      const csvText = await file.text();
      const mapping = await ServiceMapping.fromCsvText(csvText);
      if (!mapping.isLoaded) {
        set({ isLoading: false, error: "This file doesn't look like a valid service ID mapping CSV." });
        return;
      }
      const { manualEntries } = get();
      persist({ source: "csv", fileName: file.name, csvText, manualEntries });
      set({ fileName: file.name, csvText, mapping, isLoading: false, source: "csv" });
    } catch (err) {
      set({ isLoading: false, error: err instanceof Error ? err.message : "Could not read this file." });
    }
  },

  setManualEntries: (entries) => {
    const mapping = buildManualMapping(entries);
    const { fileName, csvText } = get();
    persist({ source: "manual", fileName, csvText, manualEntries: entries });
    set({ manualEntries: entries, mapping, source: "manual", error: null });
  },

  clear: () => {
    sessionStorage.removeItem(SESSION_KEY);
    set({ fileName: null, csvText: null, manualEntries: [], mapping: null, error: null, source: "manual" });
  },
}));

/** Rehydrates the store from sessionStorage on mount (call once from the app shell). */
export async function rehydrateServiceMapStore(): Promise<void> {
  const raw = sessionStorage.getItem(SESSION_KEY);
  if (!raw) return;
  try {
    const parsed = JSON.parse(raw) as {
      source: Source;
      fileName: string | null;
      csvText: string | null;
      manualEntries: ManualServiceEntry[];
    };
    if (parsed.source === "csv" && parsed.csvText) {
      const mapping = await ServiceMapping.fromCsvText(parsed.csvText);
      useServiceMapStore.setState({ ...parsed, mapping });
    } else if (parsed.manualEntries?.length) {
      const mapping = buildManualMapping(parsed.manualEntries);
      useServiceMapStore.setState({ ...parsed, mapping });
    } else {
      useServiceMapStore.setState(parsed);
    }
  } catch {
    sessionStorage.removeItem(SESSION_KEY);
  }
}
