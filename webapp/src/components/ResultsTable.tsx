"use client";

import { useMemo, useState } from "react";
import { Badge } from "@/components/ui/badge";
import { Table, TableBody, TableCell, TableHead, TableHeader, TableRow } from "@/components/ui/table";
import { cn } from "@/lib/utils";
import type { Issue, Severity } from "@/lib/validator/types";

const SEVERITY_STYLES: Record<Severity, string> = {
  ERROR: "bg-destructive/10 text-destructive",
  WARNING: "bg-warning/15 text-warning-foreground",
  FIXED: "bg-success/10 text-success",
};

function FilterSelect({
  label,
  value,
  onChange,
  options,
}: {
  label: string;
  value: string;
  onChange: (value: string) => void;
  options: string[];
}) {
  return (
    <label className="flex flex-col gap-1 text-sm">
      <span className="text-xs font-medium text-muted-foreground">{label}</span>
      <select
        value={value}
        onChange={(e) => onChange(e.target.value)}
        className="transition-standard h-9 rounded-lg border border-input bg-background px-2.5 text-sm outline-none focus-visible:ring-3 focus-visible:ring-ring/50"
      >
        {options.map((opt) => (
          <option key={opt} value={opt}>
            {opt}
          </option>
        ))}
      </select>
    </label>
  );
}

export function ResultsTable({ issues }: { issues: Issue[] }) {
  const [severity, setSeverity] = useState("All");
  const [child, setChild] = useState("All children");
  const [field, setField] = useState("All fields");

  const childOptions = useMemo(
    () => ["All children", ...Array.from(new Set(issues.map((i) => i.Child_Name))).sort()],
    [issues],
  );
  const fieldOptions = useMemo(
    () => ["All fields", ...Array.from(new Set(issues.map((i) => i.Field))).sort()],
    [issues],
  );

  const filtered = useMemo(
    () =>
      issues.filter(
        (i) =>
          (severity === "All" || i.Severity_Level === severity) &&
          (child === "All children" || i.Child_Name === child) &&
          (field === "All fields" || i.Field === field),
      ),
    [issues, severity, child, field],
  );

  if (issues.length === 0) {
    return <p className="text-sm text-muted-foreground">No issues to display.</p>;
  }

  return (
    <div className="flex flex-col gap-4">
      <div className="flex flex-wrap gap-3">
        <FilterSelect label="Severity" value={severity} onChange={setSeverity} options={["All", "ERROR", "WARNING", "FIXED"]} />
        <FilterSelect label="Child" value={child} onChange={setChild} options={childOptions} />
        <FilterSelect label="Field" value={field} onChange={setField} options={fieldOptions} />
      </div>

      <p className="text-sm text-muted-foreground">
        Showing {filtered.length.toLocaleString()} of {issues.length.toLocaleString()} entries
      </p>

      <div className="max-h-[600px] overflow-y-auto rounded-xl border border-border">
        <Table>
          <TableHeader className="sticky top-0 bg-card">
            <TableRow>
              <TableHead className="w-16">Row</TableHead>
              <TableHead>Child</TableHead>
              <TableHead>Field</TableHead>
              <TableHead>Issue</TableHead>
              <TableHead className="w-24">Severity</TableHead>
              <TableHead>Action taken</TableHead>
            </TableRow>
          </TableHeader>
          <TableBody>
            {filtered.map((issue, idx) => (
              <TableRow key={`${issue.Row}-${issue.Field}-${idx}`}>
                <TableCell className="tabular-nums text-muted-foreground">{issue.Row}</TableCell>
                <TableCell className="max-w-40 truncate">{issue.Child_Name}</TableCell>
                <TableCell className="max-w-48 truncate">{issue.Field}</TableCell>
                <TableCell className="max-w-md whitespace-normal text-foreground">{issue.Issue_Description}</TableCell>
                <TableCell>
                  <Badge className={cn("border-transparent", SEVERITY_STYLES[issue.Severity_Level])}>
                    {issue.Severity_Level}
                  </Badge>
                </TableCell>
                <TableCell className="max-w-md whitespace-normal text-muted-foreground">{issue.Action_Taken}</TableCell>
              </TableRow>
            ))}
          </TableBody>
        </Table>
      </div>
    </div>
  );
}
