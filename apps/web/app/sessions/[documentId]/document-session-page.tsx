"use client";

import Link from "next/link";
import { useCallback, useEffect, useState } from "react";
import {
  Activity,
  AlertTriangle,
  ArrowLeft,
  ArrowRight,
  Check,
  Clock3,
  Database,
  Download,
  FileArchive,
  History,
  Info,
  Loader2,
  RefreshCw,
  Undo2,
  X,
} from "@officeai/ui/sonaloop-icons";
import { Button, ThemeToggle } from "@officeai/ui";
import { downloadBlob } from "@/lib/files/file-service";
import { LocaleToggle, useTranslator } from "@/lib/i18n";
import { editorHrefForSessionDocument } from "@/lib/sessions/editor-routing";
import { formatParityDiagnostics, formatParityFor, type WebFormatParity } from "@/lib/sessions/format-parity";
import type {
  WebCommandLogEntry,
  WebDocumentPayload,
  WebPendingChangeEntry,
} from "@/lib/sessions/web-sessions";

export function DocumentSessionPage({ documentId }: { readonly documentId: string }) {
  const { t, locale } = useTranslator();
  const [payload, setPayload] = useState<WebDocumentPayload | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [loading, setLoading] = useState(true);

  const refresh = useCallback(async () => {
    setLoading(true);
    setError(null);
    try {
      const res = await fetch(`/api/sessions/${encodeURIComponent(documentId)}`, { cache: "no-store" });
      const data = (await res.json()) as WebDocumentPayload | { readonly message?: string };
      if (!res.ok) {
        throw new Error(
          "message" in data && data.message ? data.message : `Document API failed (${res.status})`
        );
      }
      setPayload(data as WebDocumentPayload);
    } catch (err) {
      setPayload(null);
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setLoading(false);
    }
  }, [documentId]);

  useEffect(() => {
    void refresh();
  }, [refresh]);

  return (
    <main className="sl-app-main mx-auto flex min-h-full max-w-content flex-col px-6 py-10">
      <header className="sl-app-topbar flex items-center justify-between gap-3">
        <Link
          href="/"
          className="inline-flex items-center gap-2 text-sm text-secondary hover:text-foreground"
        >
          <ArrowLeft size={15} />
          {t("sessionDetail.back")}
        </Link>
        <div className="flex items-center gap-2">
          <LocaleToggle />
          <ThemeToggle />
        </div>
      </header>

      {error ? (
        <div className="sl-empty sl-empty--error mt-10 flex items-center gap-2 rounded-lg border border-divider bg-surface p-6 text-sm text-secondary">
          <AlertTriangle size={16} className="text-amber-600 dark:text-amber-400" />
          {t("sessionDetail.loadError")}: {error}
        </div>
      ) : loading && payload === null ? (
        <div className="sl-card mt-10 flex items-center gap-2 rounded-lg border border-divider bg-surface p-6 text-sm text-secondary">
          <Loader2 size={14} className="animate-spin" />
          {t("sessionDetail.loading")}
        </div>
      ) : payload ? (
        <DocumentDetail payload={payload} locale={locale} onRefresh={refresh} loading={loading} />
      ) : null}
    </main>
  );
}

function DocumentDetail({
  payload,
  locale,
  onRefresh,
  loading,
}: {
  readonly payload: WebDocumentPayload;
  readonly locale?: string;
  readonly onRefresh: () => Promise<void>;
  readonly loading: boolean;
}) {
  const { t } = useTranslator();
  const { document, session } = payload;
  const formatParity = formatParityFor(document.format);
  const diagnostics = [...document.diagnostics, ...formatParityDiagnostics(document.format)];
  const [exporting, setExporting] = useState(false);
  const [exportError, setExportError] = useState<string | null>(null);
  const [reviewingChangeId, setReviewingChangeId] = useState<string | null>(null);
  const [reviewError, setReviewError] = useState<string | null>(null);
  const exportDocument = useCallback(async () => {
    setExporting(true);
    setExportError(null);
    try {
      const res = await fetch(`/api/sessions/${encodeURIComponent(document.documentId)}/export`, {
        method: "POST",
      });
      if (!res.ok) {
        const data = (await res.json().catch(() => ({}))) as { readonly message?: string };
        throw new Error(
          "message" in data && data.message ? data.message : `Document export failed (${res.status})`
        );
      }
      const blob = await res.blob();
      downloadBlob(
        blob,
        filenameFromContentDisposition(res.headers.get("content-disposition"), document.name)
      );
      await onRefresh();
    } catch (err) {
      setExportError(err instanceof Error ? err.message : String(err));
    } finally {
      setExporting(false);
    }
  }, [document.documentId, document.name, onRefresh]);

  const reviewChange = useCallback(
    async (mutationId: string, decision: "approve" | "reject" | "undo") => {
      setReviewingChangeId(mutationId);
      setReviewError(null);
      try {
        const res = await fetch(
          `/api/sessions/${encodeURIComponent(document.documentId)}/changes/${encodeURIComponent(mutationId)}`,
          {
            method: "POST",
            headers: { "content-type": "application/json" },
            body: JSON.stringify({ decision }),
          }
        );
        if (!res.ok) {
          const data = (await res.json().catch(() => ({}))) as { readonly message?: string };
          throw new Error(
            "message" in data && data.message ? data.message : `Change review failed (${res.status})`
          );
        }
        await onRefresh();
      } catch (err) {
        setReviewError(err instanceof Error ? err.message : String(err));
      } finally {
        setReviewingChangeId(null);
      }
    },
    [document.documentId, onRefresh]
  );

  return (
    <>
      <section className="mt-10 flex flex-col gap-3 border-b border-divider pb-6">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div className="sl-badge inline-flex w-fit items-center gap-2 rounded-full border border-divider bg-surface px-2.5 py-1 text-xs font-medium uppercase text-secondary">
            <Info size={13} />
            {t("sessionDetail.inspector")} · {document.format}
          </div>
          <div className="flex items-center gap-2">
            <Link href={editorHrefForSessionDocument(document)}>
              <Button variant="primary" size="sm">
                {t("sessionDetail.openInEditor")}
                <ArrowRight size={14} className="ml-1.5" />
              </Button>
            </Link>
            <Button variant="secondary" size="sm" onClick={() => void exportDocument()} disabled={exporting}>
              {exporting ? (
                <Loader2 size={14} className="mr-1.5 animate-spin" />
              ) : (
                <Download size={14} className="mr-1.5" />
              )}
              {t("common.export")}
            </Button>
            <Button variant="ghost" size="sm" onClick={onRefresh} disabled={loading || exporting}>
              {loading ? (
                <Loader2 size={14} className="mr-1.5 animate-spin" />
              ) : (
                <RefreshCw size={14} className="mr-1.5" />
              )}
              {t("home.refresh")}
            </Button>
          </div>
        </div>
        <p className="max-w-prose text-sm text-secondary">{t("sessionDetail.inspectorIntro")}</p>
        <h1 className="text-3xl font-semibold tracking-tight text-foreground sm:text-4xl">{document.name}</h1>
        {exportError ? (
          <div className="flex items-center gap-2 text-sm text-amber-700 dark:text-amber-300">
            <AlertTriangle size={15} />
            {t("sessionDetail.exportError")}: {exportError}
          </div>
        ) : null}
        <div className="flex flex-wrap items-center gap-x-4 gap-y-2 text-sm text-secondary">
          <span className="inline-flex items-center gap-1.5">
            <Database size={14} />
            {session.title}
          </span>
          <span className="inline-flex items-center gap-1.5">
            <Clock3 size={14} />
            {formatDateTime(document.updatedAt, locale)}
          </span>
        </div>
      </section>

      <section className="sl-card mt-6 grid grid-cols-2 overflow-hidden rounded-lg border border-divider bg-surface p-0 text-sm sm:grid-cols-5">
        <Metric label={t("home.revision")} value={document.revision} />
        <Metric label={t("home.pending")} value={document.pendingChangeCount} />
        <Metric label={t("home.diagnostics")} value={document.diagnostics.length} />
        <Metric label={t("sessionDetail.commands")} value={document.commandLogCount} />
        <Metric label={t("home.exports")} value={document.exportCount} />
      </section>

      <section className="mt-6 grid gap-4 lg:grid-cols-[minmax(0,0.85fr)_minmax(0,1.15fr)]">
        <div className="space-y-4">
          <Panel title={t("sessionDetail.artifacts")} icon={<FileArchive size={14} />}>
            <div className="grid grid-cols-2 gap-3 text-sm">
              <ArtifactState label={t("home.original")} available={document.artifacts.hasOriginal} />
              <ArtifactState label={t("home.working")} available={document.artifacts.hasWorking} />
            </div>
          </Panel>

          <Panel title={t("sessionDetail.formatParity")} icon={<FileArchive size={14} />}>
            <FormatParityPanel parity={formatParity} />
          </Panel>

          <Panel title={t("home.exports")} icon={<History size={14} />}>
            {document.exports.length === 0 ? (
              <Empty>{t("sessionDetail.noExports")}</Empty>
            ) : (
              <div className="divide-y divide-divider text-sm">
                {document.exports.map((entry) => (
                  <div
                    key={`${entry.exportedAt}-${entry.bytes}`}
                    className="flex items-center justify-between gap-3 py-2"
                  >
                    <span className="text-secondary">{formatDateTime(entry.exportedAt, locale)}</span>
                    <span className="font-mono text-xs text-tertiary">{formatBytes(entry.bytes)}</span>
                  </div>
                ))}
              </div>
            )}
          </Panel>
        </div>

        <div className="space-y-4">
          <Panel title={t("sessionDetail.pendingChanges")} icon={<Activity size={14} />}>
            {reviewError ? (
              <div className="mb-3 flex items-center gap-2 text-sm text-amber-700 dark:text-amber-300">
                <AlertTriangle size={15} />
                {t("sessionDetail.reviewError")}: {reviewError}
              </div>
            ) : null}
            {document.pendingChanges.length === 0 ? (
              <Empty>{t("sessionDetail.noPendingChanges")}</Empty>
            ) : (
              <PendingList
                entries={document.pendingChanges}
                locale={locale}
                reviewingChangeId={reviewingChangeId}
                onReview={reviewChange}
              />
            )}
          </Panel>

          <Panel title={t("sessionDetail.commandLog")} icon={<History size={14} />}>
            {document.commandLog.length === 0 ? (
              <Empty>{t("sessionDetail.noCommandLog")}</Empty>
            ) : (
              <CommandLog entries={document.commandLog} locale={locale} />
            )}
          </Panel>

          <Panel title={t("home.diagnostics")} icon={<AlertTriangle size={14} />}>
            {diagnostics.length === 0 ? (
              <Empty>{t("home.noDiagnostics")}</Empty>
            ) : (
              <div className="divide-y divide-divider text-sm">
                {diagnostics.map((diagnostic) => (
                  <div key={`${diagnostic.level}-${diagnostic.code}-${diagnostic.message}`} className="py-2">
                    <div className="flex items-center justify-between gap-3">
                      <span className="font-medium text-foreground">{diagnostic.code}</span>
                      <span className="text-xs uppercase text-tertiary">{diagnostic.level}</span>
                    </div>
                    <p className="mt-1 text-xs text-secondary">{diagnostic.message}</p>
                  </div>
                ))}
              </div>
            )}
          </Panel>
        </div>
      </section>
    </>
  );
}

function Metric({ label, value }: { readonly label: string; readonly value: number }) {
  return (
    <div className="border-b border-r border-divider px-4 py-3 last:border-r-0 sm:border-b-0">
      <div className="text-lg font-semibold text-foreground">{value}</div>
      <div className="text-[11px] uppercase tracking-wide text-tertiary">{label}</div>
    </div>
  );
}

function Panel({
  title,
  icon,
  children,
}: {
  readonly title: string;
  readonly icon: React.ReactNode;
  readonly children: React.ReactNode;
}) {
  return (
    <section className="sl-app-inspector rounded-lg border border-divider bg-surface">
      <h2 className="sl-app-inspector__header mb-0 flex items-center gap-2 text-xs font-semibold uppercase tracking-wide text-tertiary">
        {icon}
        {title}
      </h2>
      <div className="sl-app-inspector__body">{children}</div>
    </section>
  );
}

function ArtifactState({ label, available }: { readonly label: string; readonly available: boolean }) {
  const { t } = useTranslator();
  return (
    <div className="sl-card rounded-md border border-divider px-3 py-2">
      <div className="text-xs uppercase tracking-wide text-tertiary">{label}</div>
      <div className={available ? "mt-1 text-sm font-medium text-foreground" : "mt-1 text-sm text-secondary"}>
        {available ? t("sessionDetail.available") : t("sessionDetail.missing")}
      </div>
    </div>
  );
}

function FormatParityPanel({ parity }: { readonly parity: WebFormatParity }) {
  return (
    <div className="space-y-3 text-sm">
      <div className="font-medium text-foreground">{parity.title}</div>
      <div className="divide-y divide-divider rounded-md border border-divider">
        {parity.rows.map((row) => (
          <div key={row.label} className="grid gap-2 px-3 py-2 sm:grid-cols-[9rem_minmax(0,1fr)]">
            <div>
              <div className="text-xs uppercase text-tertiary">{row.label}</div>
              <span
                className={`mt-1 inline-flex rounded border px-1.5 py-0.5 text-[11px] uppercase ${statusClass(row.status)}`}
              >
                {row.status}
              </span>
            </div>
            <div className="text-secondary">{row.detail}</div>
          </div>
        ))}
      </div>
      {parity.knownLimits.length > 0 ? (
        <div className="rounded-md border border-amber-200 bg-amber-50 px-3 py-2 text-xs text-amber-900 dark:border-amber-900/60 dark:bg-amber-950/30 dark:text-amber-200">
          {parity.knownLimits.join(" ")}
        </div>
      ) : null}
    </div>
  );
}

function statusClass(status: WebFormatParity["rows"][number]["status"]): string {
  switch (status) {
    case "full":
      return "border-emerald-300 text-emerald-700 dark:border-emerald-700 dark:text-emerald-300";
    case "partial":
      return "border-sky-300 text-sky-700 dark:border-sky-700 dark:text-sky-300";
    case "review-only":
      return "border-amber-300 text-amber-800 dark:border-amber-700 dark:text-amber-300";
    case "planned":
      return "border-divider text-secondary";
  }
}

function PendingList({
  entries,
  locale,
  reviewingChangeId,
  onReview,
}: {
  readonly entries: ReadonlyArray<WebPendingChangeEntry>;
  readonly locale?: string;
  readonly reviewingChangeId: string | null;
  readonly onReview: (mutationId: string, decision: "approve" | "reject" | "undo") => void;
}) {
  const { t } = useTranslator();
  return (
    <div className="divide-y divide-divider text-sm">
      {entries.map((entry) => {
        const isReviewing = reviewingChangeId === entry.id;
        const isBlocked = reviewingChangeId !== null;
        return (
          <OperationRow
            key={entry.id}
            title={entry.operation}
            status={entry.status}
            source={entry.source}
            actorId={entry.actorId}
            timestamp={entry.timestamp ? new Date(entry.timestamp).toISOString() : undefined}
            hasDiff={entry.hasDiff}
            diffSummary={entry.diffSummary}
            rejection={entry.rejection?.message}
            locale={locale}
            actions={
              entry.status === "pending" ? (
                <div className="mt-2 flex flex-wrap gap-2">
                  <Button
                    variant="secondary"
                    size="sm"
                    onClick={() => onReview(entry.id, "approve")}
                    disabled={isBlocked}
                  >
                    {isReviewing ? (
                      <Loader2 size={14} className="mr-1.5 animate-spin" />
                    ) : (
                      <Check size={14} className="mr-1.5" />
                    )}
                    {t("sessionDetail.approveChange")}
                  </Button>
                  <Button
                    variant="ghost"
                    size="sm"
                    onClick={() => onReview(entry.id, "reject")}
                    disabled={isBlocked}
                  >
                    {isReviewing ? (
                      <Loader2 size={14} className="mr-1.5 animate-spin" />
                    ) : (
                      <X size={14} className="mr-1.5" />
                    )}
                    {t("sessionDetail.rejectChange")}
                  </Button>
                </div>
              ) : (
                <div className="mt-2 flex flex-wrap gap-2">
                  <Button
                    variant="ghost"
                    size="sm"
                    onClick={() => onReview(entry.id, "undo")}
                    disabled={isBlocked}
                  >
                    {isReviewing ? (
                      <Loader2 size={14} className="mr-1.5 animate-spin" />
                    ) : (
                      <Undo2 size={14} className="mr-1.5" />
                    )}
                    {t("sessionDetail.undoReview")}
                  </Button>
                </div>
              )
            }
          />
        );
      })}
    </div>
  );
}

function CommandLog({
  entries,
  locale,
}: {
  readonly entries: ReadonlyArray<WebCommandLogEntry>;
  readonly locale?: string;
}) {
  return (
    <div className="divide-y divide-divider text-sm">
      {entries.map((entry) => (
        <OperationRow
          key={entry.id}
          title={`${entry.operation} · ${entry.stage}`}
          status={entry.status}
          source={entry.source}
          actorId={entry.actorId}
          timestamp={entry.recordedAt}
          hasDiff={entry.hasDiff}
          diagnosticCount={entry.diagnostics.length}
          targetRevision={entry.provenance?.targetRevision}
          exportCommandIds={entry.exportRef?.commandIds}
          locale={locale}
        />
      ))}
    </div>
  );
}

function OperationRow({
  title,
  status,
  source,
  actorId,
  timestamp,
  hasDiff,
  diffSummary,
  rejection,
  diagnosticCount,
  targetRevision,
  exportCommandIds,
  locale,
  actions,
}: {
  readonly title: string;
  readonly status: string;
  readonly source: string;
  readonly actorId?: string;
  readonly timestamp?: string;
  readonly hasDiff: boolean;
  readonly diffSummary?: string;
  readonly rejection?: string;
  readonly diagnosticCount?: number;
  readonly targetRevision?: number;
  readonly exportCommandIds?: ReadonlyArray<string>;
  readonly locale?: string;
  readonly actions?: React.ReactNode;
}) {
  const { t } = useTranslator();
  return (
    <div className="py-2">
      <div className="flex flex-wrap items-center justify-between gap-2">
        <span className="font-medium text-foreground">{title}</span>
        <span className="rounded border border-divider px-1.5 py-0.5 text-[11px] uppercase text-secondary">
          {status}
        </span>
      </div>
      <div className="mt-1 flex flex-wrap gap-x-3 gap-y-1 text-xs text-tertiary">
        <span>
          {t("sessionDetail.source")}: {source}
        </span>
        {actorId ? (
          <span>
            {t("sessionDetail.actor")}: {actorId}
          </span>
        ) : null}
        {timestamp ? <span>{formatDateTime(timestamp, locale)}</span> : null}
        <span>{hasDiff ? t("sessionDetail.diffAvailable") : t("sessionDetail.noDiff")}</span>
        {diagnosticCount !== undefined ? (
          <span>
            {t("home.diagnostics")}: {diagnosticCount}
          </span>
        ) : null}
        {targetRevision !== undefined ? <span>rev {targetRevision}</span> : null}
        {exportCommandIds && exportCommandIds.length > 0 ? (
          <span>export basis: {exportCommandIds.join(", ")}</span>
        ) : null}
      </div>
      {diffSummary ? <p className="mt-1 text-xs text-secondary">{diffSummary}</p> : null}
      {rejection ? (
        <p className="mt-1 text-xs text-secondary">
          {t("sessionDetail.rejection")}: {rejection}
        </p>
      ) : null}
      {actions}
    </div>
  );
}

function Empty({ children }: { readonly children: React.ReactNode }) {
  return <div className="py-2 text-sm text-secondary">{children}</div>;
}

function formatBytes(n: number): string {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / (1024 * 1024)).toFixed(1)} MB`;
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

function filenameFromContentDisposition(value: string | null, fallback: string): string {
  const match = value?.match(/filename="([^"]+)"/);
  return match?.[1] ?? fallback;
}
