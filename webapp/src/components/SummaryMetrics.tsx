import { AlertCircle, AlertTriangle, CheckCircle2, ListChecks, Users } from "lucide-react";
import { cn } from "@/lib/utils";

interface Metric {
  label: string;
  value: number;
  icon: React.ComponentType<{ className?: string }>;
  tone: "default" | "error" | "warning" | "success";
  helpText?: string;
}

const TONE_CLASSES: Record<Metric["tone"], string> = {
  default: "bg-muted text-foreground",
  error: "bg-destructive/10 text-destructive",
  warning: "bg-warning/15 text-warning-foreground",
  success: "bg-success/10 text-success",
};

export function SummaryMetrics({
  total,
  errors,
  warnings,
  fixed,
  intraDupes,
  crossDupes,
}: {
  total: number;
  errors: number;
  warnings: number;
  fixed: number;
  intraDupes?: number;
  crossDupes?: number;
}) {
  const metrics: Metric[] = [
    { label: "Total entries", value: total, icon: ListChecks, tone: "default" },
    { label: "Errors", value: errors, icon: AlertCircle, tone: "error", helpText: "Must be resolved before importing" },
    { label: "Warnings", value: warnings, icon: AlertTriangle, tone: "warning", helpText: "Review recommended" },
    { label: "Auto-fixed", value: fixed, icon: CheckCircle2, tone: "success", helpText: "Corrected automatically" },
  ];
  if (typeof intraDupes === "number") {
    metrics.push({ label: "Within-upload dupes", value: intraDupes, icon: Users, tone: "default" });
  }
  if (typeof crossDupes === "number") {
    metrics.push({ label: "Cross-service dupes", value: crossDupes, icon: Users, tone: "default" });
  }

  return (
    <div className="grid grid-cols-2 gap-3 sm:grid-cols-3 lg:grid-cols-6">
      {metrics.map(({ label, value, icon: Icon, tone, helpText }) => (
        <div
          key={label}
          className="transition-standard flex flex-col gap-2 rounded-xl border border-border bg-card p-4 hover:shadow-sm"
          title={helpText}
        >
          <div className={cn("flex size-8 items-center justify-center rounded-lg", TONE_CLASSES[tone])}>
            <Icon className="size-4" />
          </div>
          <div>
            <p className="text-2xl font-semibold tabular-nums text-foreground">{value.toLocaleString()}</p>
            <p className="text-xs text-muted-foreground">{label}</p>
          </div>
        </div>
      ))}
    </div>
  );
}
