"use client";

import Link from "next/link";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  AlertTriangle,
  ArrowRight,
  Clock3,
  Database,
  FileImage,
  FileSpreadsheet,
  FileText,
  FileType2,
  Loader2,
  Plus,
  Presentation,
  RefreshCw,
  Search,
  Upload,
} from "@officeai/ui/sonaloop-icons";
import { Button, buttonVariants } from "@officeai/ui";
import { useTranslator } from "@/lib/i18n";
import { editorHrefForSessionDocument } from "@/lib/sessions/editor-routing";
import {
  sessionBrowserCounts,
  type WebDocumentEntry,
  type WebOfficeFormat,
  type WebSessionEntry,
  type WebSessionsPayload,
} from "@/lib/sessions/web-sessions";

const CREATE_FORMATS: ReadonlyArray<WebOfficeFormat> = ["docx", "xlsx", "pptx", "pdf"];

const FORMAT_ICON_BY_FORMAT = {
  docx: FileText,
  xlsx: FileSpreadsheet,
  pptx: Presentation,
  pdf: FileType2,
  image: FileImage,
} satisfies Record<WebOfficeFormat, typeof FileText>;

const FORMAT_TONE_BY_FORMAT: Record<WebOfficeFormat, string> = {
  docx: "bg-blue-50 text-blue-700 border-blue-100 dark:bg-blue-950/40 dark:text-blue-200 dark:border-blue-900/60",
  xlsx: "bg-emerald-50 text-emerald-700 border-emerald-100 dark:bg-emerald-950/40 dark:text-emerald-200 dark:border-emerald-900/60",
  pptx: "bg-amber-50 text-amber-700 border-amber-100 dark:bg-amber-950/40 dark:text-amber-200 dark:border-amber-900/60",
  pdf: "bg-red-50 text-red-700 border-red-100 dark:bg-red-950/40 dark:text-red-200 dark:border-red-900/60",
  image: "bg-sky-50 text-sky-700 border-sky-100 dark:bg-sky-950/40 dark:text-sky-200 dark:border-sky-900/60",
};

export function SessionBrowser() {
  const { t, locale } = useTranslator();
  const [payload, setPayload] = useState<WebSessionsPayload | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);
  const [importing, setImporting] = useState(false);
  const [creatingFormat, setCreatingFormat] = useState<WebOfficeFormat | null>(null);
  const [query, setQuery] = useState("");
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  const refresh = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const res = await fetch("/api/sessions", { cache: "no-store" });
      const data = (await res.json()) as WebSessionsPayload | { readonly message?: string };
      if (!res.ok) {
        throw new Error(
          "message" in data && data.message ? data.message : `Session API failed (${res.status})`
        );
      }
      setPayload(data as WebSessionsPayload);
    } catch (err) {
      setPayload(null);
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setLoading(false);
    }
  }, []);

  useEffect(() => {
    void refresh();
  }, [refresh]);

  const counts = useMemo(() => (payload ? sessionBrowserCounts(payload) : null), [payload]);
  const visibleDocuments = useMemo(
    () => (payload ? filterDocuments(payload, query) : null),
    [payload, query]
  );
  const sessionById = useMemo(
    () => new Map(payload?.sessions.map((session) => [session.sessionId, session]) ?? []),
    [payload]
  );
  const importFile = useCallback(
    async (file: File) => {
      setImporting(true);
      setError(null);
      try {
        const form = new FormData();
        form.set("file", file);
        const res = await fetch("/api/sessions/import", { method: "POST", body: form });
        const data = (await res.json()) as { readonly message?: string };
        if (!res.ok) {
          throw new Error(
            "message" in data && data.message ? data.message : `Session import failed (${res.status})`
          );
        }
        await refresh();
      } catch (err) {
        setError(err instanceof Error ? err.message : String(err));
      } finally {
        setImporting(false);
      }
    },
    [refresh]
  );

  const onImportInput = useCallback(
    (event: React.ChangeEvent<HTMLInputElement>) => {
      const file = event.currentTarget.files?.[0];
      event.currentTarget.value = "";
      if (file) void importFile(file);
    },
    [importFile]
  );
  const createDocument = useCallback(
    async (format: WebOfficeFormat) => {
      setCreatingFormat(format);
      setError(null);
      try {
        const res = await fetch("/api/sessions/create", {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify({ format }),
        });
        const data = (await res.json()) as { readonly message?: string };
        if (!res.ok) {
          throw new Error(
            "message" in data && data.message ? data.message : `Session create failed (${res.status})`
          );
        }
        await refresh();
      } catch (err) {
        setError(err instanceof Error ? err.message : String(err));
      } finally {
        setCreatingFormat(null);
      }
    },
    [refresh]
  );
  return (
    <section className="mt-10">
      <div className="mb-3 flex flex-wrap items-center justify-between gap-3">
        <h2 className="flex items-center gap-2 text-xs font-semibold uppercase tracking-wide text-tertiary">
          <Database size={14} />
          {t("home.workspace")}
        </h2>
        <div className="flex items-center gap-2">
          <div className="flex flex-wrap items-center gap-1">
            {CREATE_FORMATS.map((format) => (
              <Button
                key={format}
                variant="ghost"
                size="sm"
                onClick={() => void createDocument(format)}
                disabled={creatingFormat !== null || importing}
              >
                {creatingFormat === format ? (
                  <Loader2 size={14} className="mr-1.5 animate-spin" />
                ) : (
                  <Plus size={14} className="mr-1.5" />
                )}
                {format.toUpperCase()}
              </Button>
            ))}
          </div>
          <input
            ref={fileInputRef}
            className="hidden"
            type="file"
            accept=".docx,.xlsx,.pptx,.pdf,.png,.jpg,.jpeg,.webp,.gif,.svg,.bmp,.tif,.tiff,.heic,.heif,application/pdf,image/*"
            onChange={onImportInput}
          />
          <Button
            variant="secondary"
            size="sm"
            onClick={() => fileInputRef.current?.click()}
            disabled={importing || creatingFormat !== null}
          >
            {importing ? (
              <Loader2 size={14} className="mr-1.5 animate-spin" />
            ) : (
              <Upload size={14} className="mr-1.5" />
            )}
            {t("home.importDocument")}
          </Button>
          <Button
            variant="ghost"
            size="sm"
            onClick={refresh}
            disabled={loading || importing || creatingFormat !== null}
          >
            {loading ? (
              <Loader2 size={14} className="mr-1.5 animate-spin" />
            ) : (
              <RefreshCw size={14} className="mr-1.5" />
            )}
            {t("home.refresh")}
          </Button>
        </div>
      </div>

      <div className="sl-card overflow-hidden rounded-lg border border-divider bg-surface p-0">
        {error ? (
          <div className="sl-empty sl-empty--error flex items-center gap-2 p-6 text-sm text-secondary">
            <AlertTriangle size={16} className="text-amber-600 dark:text-amber-400" />
            {t("home.workspaceError")}: {error}
          </div>
        ) : loading && payload === null ? (
          <div className="flex items-center gap-2 p-6 text-sm text-secondary">
            <Loader2 size={14} className="animate-spin" />
            {t("home.workspaceLoading")}
          </div>
        ) : payload && payload.documents.length === 0 ? (
          <div className="sl-empty sl-empty--no-results p-6 text-sm text-secondary">
            {t("home.workspaceEmpty")}
          </div>
        ) : payload && counts && visibleDocuments ? (
          <div>
            <div className="border-b border-divider p-3">
              <div className="flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
                <label className="flex min-w-0 flex-1 items-center gap-2 rounded-md border border-divider bg-background px-3 py-2 text-sm text-secondary">
                  <Search size={14} />
                  <input
                    value={query}
                    onChange={(event) => setQuery(event.currentTarget.value)}
                    placeholder={t("home.searchSessions")}
                    className="min-w-0 flex-1 bg-transparent text-foreground outline-none placeholder:text-tertiary"
                  />
                </label>
                <div className="text-xs text-tertiary">
                  {counts.documents}{" "}
                  {counts.documents === 1 ? t("home.workspaceDocument") : t("home.workspaceDocuments")}
                </div>
              </div>
            </div>
            <div className="divide-y divide-divider">
              {visibleDocuments.length === 0 ? (
                <div className="p-6 text-sm text-secondary">{t("home.workspaceNoMatches")}</div>
              ) : (
                visibleDocuments.map((document) => (
                  <DocumentListRow
                    key={document.documentId}
                    document={document}
                    session={sessionById.get(document.sessionId)}
                    locale={locale}
                  />
                ))
              )}
            </div>
          </div>
        ) : null}
      </div>
    </section>
  );
}

function DocumentListRow({
  document,
  session,
  locale,
}: {
  readonly document: WebDocumentEntry;
  readonly session?: WebSessionEntry;
  readonly locale?: string;
}) {
  const { t } = useTranslator();
  const Icon = FORMAT_ICON_BY_FORMAT[document.format];
  const diagnosticsLabel =
    document.diagnostics.length > 0
      ? `${document.diagnostics.length} ${
          document.diagnostics.length === 1 ? t("home.diagnostic") : t("home.diagnostics")
        }`
      : null;
  return (
    <div className="flex items-center gap-3 px-4 py-3">
      <Link
        href={editorHrefForSessionDocument(document)}
        aria-hidden="true"
        tabIndex={-1}
        className={`flex h-10 w-10 shrink-0 items-center justify-center rounded-md border ${FORMAT_TONE_BY_FORMAT[document.format]}`}
      >
        <Icon size={22} />
      </Link>
      <div className="min-w-0 flex-1">
        <Link href={editorHrefForSessionDocument(document)}>
          <span className="block truncate text-sm font-medium text-foreground hover:underline">
            {document.name}
          </span>
        </Link>
        <div className="mt-0.5 flex flex-wrap items-center gap-x-2 gap-y-1 text-xs text-tertiary">
          <span className="uppercase">{document.format}</span>
          <span>
            {t("home.revision")} {document.revision}
          </span>
          <span className="inline-flex items-center gap-1">
            <Clock3 size={11} />
            {formatDateTime(document.updatedAt, locale)}
          </span>
          {session ? <span className="truncate">{session.title}</span> : null}
          {document.pendingChangeCount > 0 ? (
            <span>
              {t("home.pending")}: {document.pendingChangeCount}
            </span>
          ) : null}
          {diagnosticsLabel ? <span>{diagnosticsLabel}</span> : null}
        </div>
      </div>
      <Link
        href={editorHrefForSessionDocument(document)}
        className={buttonVariants({ variant: "secondary", size: "sm", className: "shrink-0" })}
      >
        {t("home.openInEditor")}
        <ArrowRight size={14} className="ml-1.5" />
      </Link>
    </div>
  );
}

function formatDateTime(iso: string, locale?: string): string {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return "-";
  return d.toLocaleString(locale, {
    year: "numeric",
    month: "short",
    day: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function filterDocuments(payload: WebSessionsPayload, query: string): ReadonlyArray<WebDocumentEntry> {
  const needle = query.trim().toLowerCase();
  const sessionById = new Map(payload.sessions.map((session) => [session.sessionId, session]));
  const documents = needle
    ? payload.documents.filter((document) =>
        [
          document.name,
          document.format,
          document.status,
          document.documentId,
          sessionById.get(document.sessionId)?.title ?? "",
        ]
          .join(" ")
          .toLowerCase()
          .includes(needle)
      )
    : payload.documents;
  return [...documents].sort((a, b) => Date.parse(b.updatedAt) - Date.parse(a.updatedAt));
}
