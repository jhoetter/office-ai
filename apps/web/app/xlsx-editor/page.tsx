"use client";

import dynamic from "next/dynamic";
import Link from "next/link";
import { ArrowLeft } from "lucide-react";
import { ThemeToggle } from "@officeai/ui";

const XlsxEditor = dynamic(() => import("./XlsxEditor").then((m) => m.XlsxEditor), {
  ssr: false,
  loading: () => (
    <div className="flex h-full items-center justify-center text-sm text-secondary">Loading editor…</div>
  ),
});

export default function XlsxEditorPage() {
  return (
    <main className="mx-auto flex h-screen max-w-[1200px] flex-col px-6 py-6">
      <header className="flex items-center justify-between pb-3">
        <div className="flex items-center gap-3">
          <Link
            href="/"
            className="inline-flex items-center gap-1.5 text-sm text-secondary transition-colors hover:text-foreground"
          >
            <ArrowLeft size={14} />
            Back
          </Link>
          <span className="text-sm text-tertiary">·</span>
          <span className="text-sm font-medium">XLSX editor</span>
          <span className="rounded-full bg-[var(--ai-violet-light)] px-2 py-0.5 text-[10px] font-medium uppercase tracking-wide text-[var(--ai-violet)]">
            AI-native
          </span>
        </div>
        <ThemeToggle />
      </header>
      <div className="flex-1 min-h-0">
        <XlsxEditor />
      </div>
    </main>
  );
}
