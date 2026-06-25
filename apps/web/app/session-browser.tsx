"use client";

import Link from "next/link";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import {
  AlertTriangle,
  ArrowRight,
  Clock3,
  Database,
  FileArchive,
  Info,
  Loader2,
  Plus,
  RefreshCw,
  Upload,
} from "@officeai/ui/sonaloop-icons";
import { Button } from "@officeai/ui";
import { useTranslator } from "@/lib/i18n";
import { editorHrefForSessionDocument, inspectorHrefForDocumentId } from "@/lib/sessions/editor-routing";
import { formatParityDiagnostics, formatParityFor } from "@/lib/sessions/format-parity";
import {
  documentsForSession,
  sessionBrowserCounts,
  type WebDocumentEntry,
  type WebOfficeFormat,
  type WebSessionEntry,
  type WebSessionsPayload,
} from "@/lib/sessions/web-sessions";

const CREATE_FORMATS: ReadonlyArray<WebOfficeFormat> = ["docx", "xlsx", "pptx", "pdf"];

export function SessionBrowser() {
  const { t, locale } = useTranslator();
  const [payload, setPayload] = useState<WebSessionsPayload | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);
  const [importing, setImporting] = useState(false);
  const [creatingFormat, setCreatingFormat] = useState<WebOfficeFormat | null>(null);
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
            accept=".docx,.xlsx,.pptx,.pdf,application/pdf"
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
        ) : payload && payload.sessions.length === 0 ? (
          <div className="sl-empty sl-empty--no-results p-6 text-sm text-secondary">
            {t("home.workspaceEmpty")}
          </div>
        ) : payload && counts ? (
          <div>
            <div className="grid grid-cols-2 border-b border-divider text-xs text-secondary sm:grid-cols-4">
              <Metric label={t("home.sessions")} value={counts.sessions} />
              <Metric label={t("home.documents")} value={counts.documents} />
              <Metric label={t("home.pending")} value={counts.pending} />
              <Metric label={t("home.diagnostics")} value={counts.diagnostics} />
            </div>
            <div className="divide-y divide-divider">
              {payload.sessions.map((session) => (
                <SessionRow
                  key={session.sessionId}
                  session={session}
                  documents={documentsForSession(payload, session.sessionId)}
                  locale={locale}
                />
              ))}
            </div>
          </div>
        ) : null}
      </div>
    </section>
  );
}

function Metric({ label, value }: { readonly label: string; readonly value: number }) {
  return (
    <div className="px-4 py-3">
      <div className="text-lg font-semibold text-foreground">{value}</div>
      <div className="text-[11px] uppercase tracking-wide text-tertiary">{label}</div>
    </div>
  );
}

function SessionRow({
  session,
  documents,
  locale,
}: {
  readonly session: WebSessionEntry;
  readonly documents: ReadonlyArray<WebDocumentEntry>;
  readonly locale?: string;
}) {
  const { t } = useTranslator();
  return (
    <div className="px-4 py-3">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <div>
          <div className="text-sm font-medium text-foreground">{session.title}</div>
          <div className="mt-0.5 flex items-center gap-1.5 text-xs text-tertiary">
            <Clock3 size={12} />
            {formatDateTime(session.updatedAt, locale)}
          </div>
        </div>
        <div className="text-xs text-secondary">
          {session.documentCount}{" "}
          {session.documentCount === 1 ? t("home.workspaceDocument") : t("home.workspaceDocuments")}
        </div>
      </div>
      <div className="mt-3 overflow-hidden rounded-md border border-divider">
        {documents.length === 0 ? (
          <div className="px-3 py-2 text-xs text-secondary">{t("home.workspaceNoDocuments")}</div>
        ) : (
          <table className="sl-table w-full text-sm">
            <tbody>
              {documents.map((document) => (
                <DocumentRow key={document.documentId} document={document} locale={locale} />
              ))}
            </tbody>
          </table>
        )}
      </div>
    </div>
  );
}

function DocumentRow({
  document,
  locale,
}: {
  readonly document: WebDocumentEntry;
  readonly locale?: string;
}) {
  const { t } = useTranslator();
  const parity = formatParityFor(document.format);
  const latestDiagnostic = document.diagnostics[0] ?? formatParityDiagnostics(document.format)[0];
  return (
    <tr className="border-b border-divider last:border-b-0">
      <td className="px-3 py-2.5">
        <div className="flex min-w-0 items-center gap-2">
          <span className="rounded border border-divider px-1.5 py-0.5 text-[11px] font-medium uppercase text-secondary">
            {document.format}
          </span>
          <div className="min-w-0">
            <Link
              href={editorHrefForSessionDocument(document)}
              className="block truncate text-sm font-medium text-foreground hover:underline"
            >
              {document.name}
            </Link>
            <div className="text-xs text-tertiary">
              {t("home.revision")} {document.revision} · {formatDateTime(document.updatedAt, locale)}
            </div>
            <div className="mt-0.5 flex flex-wrap items-center gap-x-2 gap-y-1 text-xs text-tertiary">
              <span>{parity.title}</span>
              <Link
                href={inspectorHrefForDocumentId(document.documentId)}
                className="inline-flex items-center gap-1 text-secondary hover:text-foreground"
              >
                <Info size={11} />
                {t("sessionDetail.inspector")}
              </Link>
            </div>
          </div>
        </div>
      </td>
      <td className="hidden px-3 py-2.5 text-xs text-secondary md:table-cell">
        <div className="flex items-center gap-1.5">
          <FileArchive size={12} />
          {document.artifacts.hasOriginal ? t("home.original") : t("home.created")} /{" "}
          {document.artifacts.hasWorking ? t("home.working") : t("home.noWorking")}
        </div>
      </td>
      <td className="px-3 py-2.5 text-right text-xs text-secondary">
        <div>
          {t("home.pending")}: {document.pendingChangeCount}
        </div>
        <div>
          {t("home.exports")}: {document.exportCount}
        </div>
      </td>
      <td className="hidden max-w-[220px] px-3 py-2.5 text-xs text-secondary lg:table-cell">
        {latestDiagnostic ? (
          <span className={latestDiagnostic.level === "error" ? "text-red-600 dark:text-red-400" : ""}>
            {latestDiagnostic.code}
          </span>
        ) : (
          t("home.noDiagnostics")
        )}
      </td>
      <td className="px-3 py-2.5 text-right">
        <Link href={editorHrefForSessionDocument(document)}>
          <Button variant="secondary" size="sm">
            {t("home.openInEditor")}
            <ArrowRight size={14} className="ml-1.5" />
          </Button>
        </Link>
      </td>
    </tr>
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
