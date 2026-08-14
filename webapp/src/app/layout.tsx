import type { Metadata } from "next";
import { Geist, Geist_Mono } from "next/font/google";
import Link from "next/link";
import { ShieldCheck } from "lucide-react";
import { TooltipProvider } from "@/components/ui/tooltip";
import "./globals.css";

const geistSans = Geist({
  variable: "--font-geist-sans",
  subsets: ["latin"],
});

const geistMono = Geist_Mono({
  variable: "--font-geist-mono",
  subsets: ["latin"],
});

export const metadata: Metadata = {
  title: "Xplor Data Migration Tools",
  description: "Validate and prepare QikKids exports for Xplor import — entirely in your browser.",
};

export default function RootLayout({ children }: LayoutProps<"/">) {
  return (
    <html
      lang="en"
      data-scroll-behavior="smooth"
      className={`${geistSans.variable} ${geistMono.variable} h-full antialiased`}
    >
      <body className="min-h-full flex flex-col bg-background text-foreground">
        <TooltipProvider>
          <header className="border-b border-border bg-card/80 backdrop-blur">
            <div className="mx-auto flex max-w-6xl items-center justify-between px-4 py-3 sm:px-6">
              <Link
                href="/"
                aria-label="Go to homepage"
                className="transition-standard -ml-1.5 flex items-center gap-2 rounded-lg py-1 pl-1.5 pr-2.5 font-semibold text-foreground hover:bg-muted focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring/50"
              >
                <span className="flex size-8 items-center justify-center rounded-lg bg-primary text-primary-foreground">X</span>
                <span>Xplor Data Migration Tools</span>
              </Link>
              <div className="flex items-center gap-1.5 rounded-full bg-success/10 px-3 py-1 text-xs font-medium text-success">
                <ShieldCheck className="size-3.5" aria-hidden />
                Files never leave your browser
              </div>
            </div>
          </header>
          <main className="flex-1">{children}</main>
          <footer className="border-t border-border py-6 text-center text-xs text-muted-foreground">
            Xplor Data Migration Tools · All processing happens locally in your browser · Nothing is uploaded, logged, or stored
          </footer>
        </TooltipProvider>
      </body>
    </html>
  );
}
