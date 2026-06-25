"use client";

import Link from "next/link";
import { useRouter } from "next/navigation";
import { useCallback, useEffect, useState } from "react";
import {
  BookOpen,
  FileSpreadsheet,
  FileText,
  FolderOpen,
  Loader2,
  Plus,
  Presentation,
  Sparkles,
} from "@officeai/ui/sonaloop-icons";
import { Button, ThemeToggle } from "@officeai/ui";
import { LocaleToggle, useTranslator } from "@/lib/i18n";
import { openFile } from "@/lib/files/file-service";
import { editorPathForFormat, sampleHrefForFormat } from "@/lib/sessions/editor-routing";
import { SessionBrowser } from "./session-browser";

type Kind = "docx" | "xlsx" | "pptx" | "pdf";

interface SampleFileEntry {
  readonly name: string;
  readonly url: string;
  readonly kind: Kind;
  readonly size: number;
  readonly modifiedAt: string;
}

interface NewAction {
  readonly id: Kind;
  readonly titleKey: string;
  readonly subtitleKey: string;
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
//
// PDF is intentionally NOT a "create new" option — there's no useful
// blank-PDF starting point, but PDFs from the sample-files folder
// still open through the `/pdf-viewer` route via `KIND_META.pdf`.
const NEW_ACTIONS: ReadonlyArray<NewAction> = [
  {
    id: "docx",
    titleKey: "home.newDocument",
    subtitleKey: "home.subDocx",
    href: "/editor?new=1",
    icon: FileText,
    accent: "text-[var(--office-blue)]",
  },
  {
    id: "xlsx",
    titleKey: "home.newSpreadsheet",
    subtitleKey: "home.subXlsx",
    href: "/xlsx-editor?new=1",
    icon: FileSpreadsheet,
    accent: "text-emerald-600 dark:text-emerald-400",
  },
  {
    id: "pptx",
    titleKey: "home.newPresentation",
    subtitleKey: "home.subPptx",
    href: "/pptx-editor?new=1",
    icon: Presentation,
    accent: "text-orange-600 dark:text-orange-400",
  },
];

const KIND_META: Record<
  Kind,
  { editorPath: string; icon: typeof FileText; labelKey: string; accent: string }
> = {
  docx: {
    editorPath: editorPathForFormat("docx"),
    icon: FileText,
    labelKey: "common.kindDocx",
    accent: "text-[var(--office-blue)]",
  },
  xlsx: {
    editorPath: editorPathForFormat("xlsx"),
    icon: FileSpreadsheet,
    labelKey: "common.kindXlsx",
    accent: "text-emerald-600 dark:text-emerald-400",
  },
  pptx: {
    editorPath: editorPathForFormat("pptx"),
    icon: Presentation,
    labelKey: "common.kindPptx",
    accent: "text-orange-600 dark:text-orange-400",
  },
  pdf: {
    editorPath: editorPathForFormat("pdf"),
    icon: BookOpen,
    labelKey: "common.kindPdf",
    accent: "text-rose-600 dark:text-rose-400",
  },
};

function formatBytes(n: number): string {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / (1024 * 1024)).toFixed(1)} MB`;
}

function formatDate(iso: string, locale?: string): string {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return "—";
  return d.toLocaleDateString(locale, { year: "numeric", month: "short", day: "numeric" });
}

function sampleHref(file: SampleFileEntry): string {
  return sampleHrefForFormat({ format: file.kind, url: file.url, name: file.name });
}

export default function HomePage() {
  const { t } = useTranslator();
  const router = useRouter();
  const [files, setFiles] = useState<ReadonlyArray<SampleFileEntry> | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [openingPdf, setOpeningPdf] = useState(false);

  // Open a PDF from the user's disk and route into the viewer. We
  // deliberately go through a `blob:` URL rather than threading the
  // bytes via React state because /pdf-viewer is a sibling page —
  // the only handoff that survives a Next client navigation without
  // mutating its URL contract is the `?src=` param the viewer
  // already consumes for sample-files clicks. The blob stays valid
  // for the lifetime of the document, which matches the editor
  // session. We don't capture the FileSystemFileHandle here because
  // it isn't serialisable across the navigation; if the user wants
  // save-back-to-disk they can re-open via the in-editor "Open"
  // button which keeps the handle locally.
  const handleOpenPdf = useCallback(async () => {
    if (openingPdf) return;
    setOpeningPdf(true);
    try {
      const opened = await openFile({
        description: "PDF document",
        mimeToExt: { "application/pdf": [".pdf"] },
        accept: ".pdf,application/pdf",
      });
      if (!opened) return;
      const blob = new Blob([opened.bytes as BlobPart], { type: "application/pdf" });
      const url = URL.createObjectURL(blob);
      const params = new URLSearchParams({ src: url, name: opened.name });
      router.push(`/pdf-viewer?${params.toString()}`);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setOpeningPdf(false);
    }
  }, [openingPdf, router]);

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
          <span className="font-semibold tracking-tight">{t("common.appName")}</span>
        </div>
        <div className="flex items-center gap-2">
          <LocaleToggle />
          <ThemeToggle />
        </div>
      </header>

      <section className="mt-12 flex flex-col gap-2">
        <span className="inline-flex w-fit items-center gap-1.5 rounded-full bg-[var(--ai-violet-light)] px-2.5 py-0.5 text-xs font-medium text-[var(--ai-violet)]">
          <Sparkles size={12} />
          {t("home.badge")}
        </span>
        <h1 className="text-3xl font-semibold tracking-tight text-foreground sm:text-4xl">
          {t("home.title")}
        </h1>
        <p className="max-w-prose text-sm text-secondary">{t("home.intro")}</p>
      </section>

      <section className="mt-8">
        <h2 className="mb-3 text-xs font-semibold uppercase tracking-wide text-tertiary">
          {t("home.createNew")}
        </h2>
        <div className="grid grid-cols-1 gap-3 sm:grid-cols-2 lg:grid-cols-4">
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
                  <span className="text-sm font-medium text-foreground">{t(action.titleKey)}</span>
                  <span className="text-xs text-secondary">{t(action.subtitleKey)}</span>
                </div>
                <Plus size={16} className="text-tertiary transition group-hover:text-[var(--office-blue)]" />
              </Link>
            );
          })}
          <button
            type="button"
            onClick={handleOpenPdf}
            disabled={openingPdf}
            className="group flex items-center gap-3 rounded-lg border border-divider bg-surface p-4 text-left transition hover:border-[var(--office-blue)] hover:shadow-sm disabled:cursor-progress disabled:opacity-70"
          >
            <div className="flex h-10 w-10 items-center justify-center rounded-md bg-hover">
              {openingPdf ? (
                <Loader2 size={20} className="animate-spin text-rose-600 dark:text-rose-400" />
              ) : (
                <BookOpen size={20} className="text-rose-600 dark:text-rose-400" />
              )}
            </div>
            <div className="flex flex-1 flex-col">
              <span className="text-sm font-medium text-foreground">{t("home.openPdf")}</span>
              <span className="text-xs text-secondary">{t("home.subOpenPdf")}</span>
            </div>
            <FolderOpen
              size={16}
              className="text-tertiary transition group-hover:text-[var(--office-blue)]"
            />
          </button>
        </div>
      </section>

      <SessionBrowser />

      <section className="mt-10 flex-1">
        <div className="mb-3 flex items-center justify-between">
          <h2 className="flex items-center gap-2 text-xs font-semibold uppercase tracking-wide text-tertiary">
            <FolderOpen size={14} />
            {t("home.sampleFiles")}
          </h2>
          {files !== null ? (
            <span className="text-xs text-tertiary">
              {files.length} {files.length === 1 ? t("home.fileOne") : t("home.fileMany")}
            </span>
          ) : null}
        </div>

        {error ? (
          <div className="rounded-lg border border-divider bg-surface p-6 text-sm text-secondary">
            {t("home.loadError")}: {error}
          </div>
        ) : files === null ? (
          <div className="flex items-center gap-2 rounded-lg border border-divider bg-surface p-6 text-sm text-secondary">
            <Loader2 size={14} className="animate-spin" />
            {t("home.listing")}
          </div>
        ) : files.length === 0 ? (
          <div className="rounded-lg border border-divider bg-surface p-6 text-sm text-secondary">
            {t("home.noSamplesLong")}
          </div>
        ) : (
          <SampleFileTable files={files} />
        )}
      </section>

      <footer className="mt-12 flex flex-wrap items-center gap-3 border-t border-divider pt-6 text-xs text-tertiary">
        <span>{t("home.footer")}</span>
      </footer>
    </main>
  );
}

function SampleFileTable({ files }: { files: ReadonlyArray<SampleFileEntry> }) {
  const { t, locale } = useTranslator();
  return (
    <div className="overflow-hidden rounded-lg border border-divider bg-surface">
      <table className="w-full text-sm">
        <thead>
          <tr className="border-b border-divider text-left text-xs uppercase tracking-wide text-tertiary">
            <th className="px-4 py-2 font-medium">{t("home.tableName")}</th>
            <th className="px-4 py-2 font-medium">{t("home.tableType")}</th>
            <th className="hidden px-4 py-2 font-medium sm:table-cell">{t("home.tableModified")}</th>
            <th className="hidden px-4 py-2 text-right font-medium sm:table-cell">{t("home.tableSize")}</th>
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
                <td className="px-4 py-2.5 text-xs text-secondary">{t(meta.labelKey)}</td>
                <td className="hidden px-4 py-2.5 text-xs text-secondary sm:table-cell">
                  {formatDate(file.modifiedAt, locale)}
                </td>
                <td className="hidden px-4 py-2.5 text-right text-xs text-secondary sm:table-cell">
                  {formatBytes(file.size)}
                </td>
                <td className="px-4 py-2.5 text-right">
                  <Link href={sampleHref(file)}>
                    <Button variant="ghost" size="sm">
                      {t("common.open")}
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
