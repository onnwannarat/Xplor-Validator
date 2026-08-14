import Link from "next/link";
import { ArrowRight, Sparkles } from "lucide-react";
import { Card, CardContent, CardDescription, CardHeader, CardTitle } from "@/components/ui/card";
import { Badge } from "@/components/ui/badge";
import { cn } from "@/lib/utils";

const TOOLS = [
  {
    icon: "🔍",
    name: "Parent and Child Import",
    href: "/parent-child-import",
    description:
      "Validates your full migration CSV/XLSX against Xplor import rules. Flags errors, warnings, and auto-fixes, then produces split output CSVs and audit reports — ready to download.",
    available: true,
  },
  {
    icon: "📅",
    name: "Bookings Import",
    href: "/bookings-import",
    description:
      "Processes recurring and casual QikKids booking exports, removes duplicates and schedule overlaps, and produces split files per service ready for Xplor import.",
    available: true,
  },
  {
    icon: "💳",
    name: "Payment Plans Import",
    href: "/payment-plans-import",
    description:
      "Validates payment plan CSVs against Onboarding Tool error rules, auto-fixes date formats and weekday abbreviations, and splits output by service.",
    available: true,
  },
  {
    icon: "⚖️",
    name: "Balance Adjustments",
    href: "/balance-adjustments",
    description:
      "Reads balance adjustment data, maps centre names to Xplor service names, removes demo accounts, and produces one styled import file per centre.",
    available: true,
  },
  {
    icon: "🏷️",
    name: "Room & Fee Names checking",
    href: "/room-fee-names",
    description:
      "Compares fee and room names used in QikKids bookings against those defined in Xplor. Produces a mismatch report with fuzzy-match suggestions.",
    available: true,
  },
];

export default function Home() {
  return (
    <div className="mx-auto flex max-w-6xl flex-col gap-10 px-4 py-10 sm:px-6">
      <section className="flex flex-col gap-4 rounded-2xl bg-gradient-to-br from-primary via-primary to-accent p-8 text-primary-foreground sm:p-10">
        <div className="flex items-center gap-2 text-sm font-medium opacity-90">
          <Sparkles className="size-4" aria-hidden />
          Runs entirely in your browser
        </div>
        <h1 className="text-3xl font-semibold tracking-tight sm:text-4xl">Xplor Data Migration Tools</h1>
        <p className="max-w-2xl text-base opacity-90">
          A suite of tools to validate, clean, and prepare QikKids data for Xplor import. Files are parsed and
          validated locally — nothing is uploaded to a server, logged, or stored.
        </p>
      </section>

      <section className="flex flex-col gap-4">
        <h2 className="text-lg font-semibold text-foreground">Available tools</h2>
        <div className="grid grid-cols-1 gap-4 sm:grid-cols-2 lg:grid-cols-3">
          {TOOLS.map((tool) => {
            const card = (
              <Card
                className={cn(
                  "transition-standard h-full ring-1 ring-border",
                  tool.available ? "hover:-translate-y-0.5 hover:shadow-md hover:ring-primary/40" : "opacity-60",
                )}
              >
                <CardHeader>
                  <div className="flex items-start justify-between gap-2">
                    <span className="text-2xl" aria-hidden>
                      {tool.icon}
                    </span>
                    {!tool.available && <Badge variant="secondary">Coming soon</Badge>}
                  </div>
                  <CardTitle className="flex items-center gap-1.5 text-base">
                    {tool.name}
                    {tool.available && <ArrowRight className="size-4 text-primary transition-transform group-hover:translate-x-0.5" />}
                  </CardTitle>
                  <CardDescription>{tool.description}</CardDescription>
                </CardHeader>
                <CardContent />
              </Card>
            );
            return tool.available ? (
              <Link key={tool.name} href={tool.href} className="group rounded-xl focus-visible:outline-none focus-visible:ring-3 focus-visible:ring-ring/50">
                {card}
              </Link>
            ) : (
              <div key={tool.name} aria-disabled className="rounded-xl">
                {card}
              </div>
            );
          })}
        </div>
      </section>
    </div>
  );
}
