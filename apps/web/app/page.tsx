"use client";

import Link from "next/link";
import { useEffect, useState } from "react";
import { FileSpreadsheet, FileText, FolderOpen, Loader2, Plus, Presentation, Sparkles } from "lucide-react";
import { Button, ThemeToggle } from "@officeai/ui";

type Kind = "docx" | "xlsx" | "pptx";

interface SampleFileEntry {
  readonly name: string;
  readonly url: string;
  readonly kind: Kind;
  readonly size: number;
  readonly modifiedAt: string;
}

interface NewAction {
  readonly id: Kind;
  readonly title: string;
  readonly subtitle: string;
  readonly href: string;
  readonly icon: typeof FileText;
  readonly accent: string;
}

// The "create new" tiles open each editor with `?new=1`, which the
// editor reads as "bootstrap the blank builder instead of the
// synthetic welcome sample" — so users land in a truly empty file
// rather than the demo content. Plain `/editor`, `/xlsx-editor` and
// `/pptx-editor` (no query param) still open the welcome sample, so
// direct-link smoke tests keep working.
const NEW_ACTIONS: ReadonlyArray<NewAction> = [
  {
    id: "docx",
    title: "New document",
    subtitle: "Word-compatible .docx",
    href: "/editor?new=1",
    icon: FileText,
    accent: "text-[var(--office-blue)]",
  },
  {
    id: "xlsx",
    title: "New spreadsheet",
    subtitle: "Excel-compatible .xlsx",
    href: "/xlsx-editor?new=1",
    icon: FileSpreadsheet,
    accent: "text-emerald-600 dark:text-emerald-400",
  },
  {
    id: "pptx",
    title: "New presentation",
    subtitle: "PowerPoint-compatible .pptx",
    href: "/pptx-editor?new=1",
    icon: Presentation,
    accent: "text-orange-600 dark:text-orange-400",
  },
];

const KIND_META: Record<Kind, { editorPath: string; icon: typeof FileText; label: string; accent: string }> =
  {
    docx: {
      editorPath: "/editor",
      icon: FileText,
      label: "Word document",
      accent: "text-[var(--office-blue)]",
    },
    xlsx: {
      editorPath: "/xlsx-editor",
      icon: FileSpreadsheet,
      label: "Excel workbook",
      accent: "text-emerald-600 dark:text-emerald-400",
    },
    pptx: {
      editorPath: "/pptx-editor",
      icon: Presentation,
      label: "PowerPoint deck",
      accent: "text-orange-600 dark:text-orange-400",
    },
  };

function formatBytes(n: number): string {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / (1024 * 1024)).toFixed(1)} MB`;
}

function formatDate(iso: string): string {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return "—";
  return d.toLocaleDateString(undefined, { year: "numeric", month: "short", day: "numeric" });
}

function sampleHref(file: SampleFileEntry): string {
  const meta = KIND_META[file.kind];
  const params = new URLSearchParams({ src: file.url, name: file.name });
  return `${meta.editorPath}?${params.toString()}`;
}

export default function HomePage() {
  const [files, setFiles] = useState<ReadonlyArray<SampleFileEntry> | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    void (async () => {
      try {
        const res = await fetch("/api/sample-files");
        if (!res.ok) throw new Error(`Failed to list samples (${res.status})`);
        const data = (await res.json()) as { files: SampleFileEntry[] };
        if (!cancelled) setFiles(data.files);
      } catch (err) {
        if (!cancelled) setError(err instanceof Error ? err.message : String(err));
      }
    })();
    return () => {
      cancelled = true;
    };
  }, []);

  return (
    <main className="mx-auto flex min-h-full max-w-content flex-col px-6 py-10">
      <header className="flex items-center justify-between">
        <div className="flex items-center gap-2">
          <div className="h-6 w-6 rounded-md bg-[var(--office-blue)]" aria-hidden />
          <span className="font-semibold tracking-tight">officeAI</span>
        </div>
        <ThemeToggle />
      </header>

      <section className="mt-12 flex flex-col gap-2">
        <span className="inline-flex w-fit items-center gap-1.5 rounded-full bg-[var(--ai-violet-light)] px-2.5 py-0.5 text-xs font-medium text-[var(--ai-violet)]">
          <Sparkles size={12} />
          AI-native office editors
        </span>
        <h1 className="text-3xl font-semibold tracking-tight text-foreground sm:text-4xl">
          Start a new file or open a sample.
        </h1>
        <p className="max-w-prose text-sm text-secondary">
          Word-, Excel- and PowerPoint-compatible editors built around an OOXML-faithful core. Every change —
          human or AI — flows through the same typed command bus.
        </p>
      </section>

      <section className="mt-8">
        <h2 className="mb-3 text-xs font-semibold uppercase tracking-wide text-tertiary">Create new</h2>
        <div className="grid grid-cols-1 gap-3 sm:grid-cols-3">
          {NEW_ACTIONS.map((action) => {
            const Icon = action.icon;
            return (
              <Link
                key={action.id}
                href={action.href}
                className="group flex items-center gap-3 rounded-lg border border-divider bg-surface p-4 transition hover:border-[var(--office-blue)] hover:shadow-sm"
              >
                <div className="flex h-10 w-10 items-center justify-center rounded-md bg-hover">
                  <Icon size={20} className={action.accent} />
                </div>
                <div className="flex flex-1 flex-col">
                  <span className="text-sm font-medium text-foreground">{action.title}</span>
                  <span className="text-xs text-secondary">{action.subtitle}</span>
                </div>
                <Plus size={16} className="text-tertiary transition group-hover:text-[var(--office-blue)]" />
              </Link>
            );
          })}
        </div>
      </section>

      <section className="mt-10 flex-1">
        <div className="mb-3 flex items-center justify-between">
          <h2 className="flex items-center gap-2 text-xs font-semibold uppercase tracking-wide text-tertiary">
            <FolderOpen size={14} />
            Sample files
          </h2>
          {files !== null ? (
            <span className="text-xs text-tertiary">
              {files.length} {files.length === 1 ? "file" : "files"}
            </span>
          ) : null}
        </div>

        {error ? (
          <div className="rounded-lg border border-divider bg-surface p-6 text-sm text-secondary">
            Couldn&apos;t load samples: {error}
          </div>
        ) : files === null ? (
          <div className="flex items-center gap-2 rounded-lg border border-divider bg-surface p-6 text-sm text-secondary">
            <Loader2 size={14} className="animate-spin" />
            Listing sample files…
          </div>
        ) : files.length === 0 ? (
          <div className="rounded-lg border border-divider bg-surface p-6 text-sm text-secondary">
            No sample files yet. Drop .docx, .xlsx or .pptx files into{" "}
            <code className="font-mono text-xs">apps/web/public/sample-files/</code> and refresh.
          </div>
        ) : (
          <SampleFileTable files={files} />
        )}
      </section>

      <footer className="mt-12 flex flex-wrap items-center gap-3 border-t border-divider pt-6 text-xs text-tertiary">
        <span>
          OOXML round-trip · typed command bus · headless <code className="font-mono">office-agent</code> CLI
          for server-side AI workflows.
        </span>
      </footer>
    </main>
  );
}

function SampleFileTable({ files }: { files: ReadonlyArray<SampleFileEntry> }) {
  return (
    <div className="overflow-hidden rounded-lg border border-divider bg-surface">
      <table className="w-full text-sm">
        <thead>
          <tr className="border-b border-divider text-left text-xs uppercase tracking-wide text-tertiary">
            <th className="px-4 py-2 font-medium">Name</th>
            <th className="px-4 py-2 font-medium">Type</th>
            <th className="hidden px-4 py-2 font-medium sm:table-cell">Modified</th>
            <th className="hidden px-4 py-2 text-right font-medium sm:table-cell">Size</th>
            <th className="px-4 py-2" />
          </tr>
        </thead>
        <tbody>
          {files.map((file) => {
            const meta = KIND_META[file.kind];
            const Icon = meta.icon;
            return (
              <tr key={file.url} className="border-b border-divider last:border-b-0 hover:bg-hover">
                <td className="px-4 py-2.5">
                  <Link
                    href={sampleHref(file)}
                    className="flex items-center gap-2.5 text-foreground hover:underline"
                  >
                    <Icon size={16} className={meta.accent} />
                    <span className="truncate">{file.name}</span>
                  </Link>
                </td>
                <td className="px-4 py-2.5 text-xs text-secondary">{meta.label}</td>
                <td className="hidden px-4 py-2.5 text-xs text-secondary sm:table-cell">
                  {formatDate(file.modifiedAt)}
                </td>
                <td className="hidden px-4 py-2.5 text-right text-xs text-secondary sm:table-cell">
                  {formatBytes(file.size)}
                </td>
                <td className="px-4 py-2.5 text-right">
                  <Link href={sampleHref(file)}>
                    <Button variant="ghost" size="sm">
                      Open
                    </Button>
                  </Link>
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}
