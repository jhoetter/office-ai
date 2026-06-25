"use client";

import Link from "next/link";
import { useRouter, useSearchParams } from "next/navigation";
import { Suspense, useCallback, useMemo, useState } from "react";
import {
  AlertTriangle,
  ArrowLeft,
  ArrowRight,
  FileArchive,
  Loader2,
  MessageSquare,
  RefreshCw,
} from "@officeai/ui/sonaloop-icons";
import { Button } from "@officeai/ui";
import { useTranslator } from "@/lib/i18n";
import { LoadingScreen } from "@/lib/shell";
import { editorHrefForSessionDocument, inspectorHrefForDocumentId } from "@/lib/sessions/editor-routing";
import { SessionEditorLoadState, useSessionEditorDocument } from "@/lib/sessions/session-editor-client";
import { parseEmailMessage, type EmailAttachment } from "@/lib/viewers/email-format";

function EmailViewerPageInner(): React.ReactNode {
  const params = useSearchParams();
  const documentId = params.get("session");
  const state = useSessionEditorDocument({ documentId, expectedFormat: "email" });
  if (!documentId || state.loading || state.error || !state.document) {
    return <SessionEditorLoadState format="email" state={state} />;
  }
  return <EmailViewer document={state.document} reload={state.reload} />;
}

function EmailViewer({
  document,
  reload,
}: {
  readonly document: NonNullable<ReturnType<typeof useSessionEditorDocument>["document"]>;
  readonly reload: () => Promise<void>;
}) {
  const { t } = useTranslator();
  const router = useRouter();
  const [routingAttachment, setRoutingAttachment] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const message = useMemo(
    () => parseEmailMessage(document.bytes, document.filename),
    [document.bytes, document.filename]
  );

  const routeAttachment = useCallback(
    async (attachment: EmailAttachment) => {
      if (!attachment.routeFormat) return;
      setRoutingAttachment(attachment.id);
      setError(null);
      try {
        const form = new FormData();
        form.set(
          "file",
          new File([arrayBufferFromBytes(attachment.bytes)], attachment.filename, {
            type: attachment.contentType,
          })
        );
        if (document.sessionId) form.set("session_id", document.sessionId);
        form.set("title", message.subject || t("emailViewer.attachments"));
        const res = await fetch("/api/sessions/import", { method: "POST", body: form });
        const data = (await res.json()) as {
          readonly message?: string;
          readonly document?: {
            readonly documentId: string;
            readonly format: EmailAttachment["routeFormat"];
          };
        };
        if (!res.ok || !data.document?.format) {
          throw new Error(data.message ?? `Attachment import failed (${res.status})`);
        }
        router.push(
          editorHrefForSessionDocument({
            documentId: data.document.documentId,
            format: data.document.format,
          })
        );
      } catch (err) {
        setError(err instanceof Error ? err.message : String(err));
      } finally {
        setRoutingAttachment(null);
      }
    },
    [document.sessionId, message.subject, router, t]
  );

  return (
    <main className="min-h-screen bg-background text-foreground">
      <header className="border-b border-divider bg-surface">
        <div className="mx-auto flex max-w-6xl flex-wrap items-center justify-between gap-3 px-6 py-4">
          <div className="min-w-0">
            <div className="flex items-center gap-2 text-xs uppercase text-tertiary">
              <MessageSquare size={14} />
              {t("emailViewer.title")}
            </div>
            <h1 className="mt-1 truncate text-xl font-semibold">{message.subject}</h1>
            <p className="mt-1 text-sm text-secondary">{document.filename}</p>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            <Link href="/">
              <Button variant="ghost" size="sm">
                <ArrowLeft size={14} className="mr-1.5" />
                {t("common.back")}
              </Button>
            </Link>
            <Link href={inspectorHrefForDocumentId(document.documentId)}>
              <Button variant="secondary" size="sm">
                {t("sessionEditor.openInspector")}
              </Button>
            </Link>
            <Button variant="ghost" size="sm" onClick={() => void reload()}>
              <RefreshCw size={14} className="mr-1.5" />
              {t("home.refresh")}
            </Button>
          </div>
        </div>
      </header>

      <div className="mx-auto grid max-w-6xl gap-6 px-6 py-6 lg:grid-cols-[280px_minmax(0,1fr)]">
        <aside className="space-y-4">
          <section className="rounded-lg border border-divider bg-surface p-4">
            <dl className="space-y-3 text-sm">
              <HeaderLine label={t("emailViewer.from")} value={message.from} />
              <HeaderLine label={t("emailViewer.to")} value={message.to} />
              <HeaderLine label={t("emailViewer.date")} value={message.date} />
              <HeaderLine label={t("common.format")} value={message.sourceKind.toUpperCase()} />
            </dl>
          </section>
          <section className="rounded-lg border border-divider bg-surface p-4">
            <h2 className="flex items-center gap-2 text-sm font-semibold">
              <FileArchive size={14} />
              {t("emailViewer.attachments")}
            </h2>
            {message.attachments.length === 0 ? (
              <p className="mt-3 text-sm text-secondary">{t("emailViewer.noAttachments")}</p>
            ) : (
              <div className="mt-3 space-y-2">
                {message.attachments.map((attachment) => (
                  <AttachmentRow
                    key={attachment.id}
                    attachment={attachment}
                    busy={routingAttachment === attachment.id}
                    onRoute={() => void routeAttachment(attachment)}
                  />
                ))}
              </div>
            )}
          </section>
        </aside>

        <section className="min-w-0 rounded-lg border border-divider bg-surface p-5">
          {error ? (
            <div className="mb-4 flex items-center gap-2 rounded-md border border-amber-300 bg-amber-50 px-3 py-2 text-sm text-amber-900 dark:border-amber-800 dark:bg-amber-950 dark:text-amber-100">
              <AlertTriangle size={14} />
              {error}
            </div>
          ) : null}
          {message.bodyHtml ? (
            <iframe
              title={t("emailViewer.body")}
              className="h-[68vh] w-full rounded-md border border-divider bg-white"
              sandbox=""
              srcDoc={message.bodyHtml}
            />
          ) : (
            <pre className="min-h-[50vh] whitespace-pre-wrap break-words font-sans text-sm leading-6 text-foreground">
              {message.bodyText || t("emailViewer.noBody")}
            </pre>
          )}
          {message.diagnostics.length > 0 ? (
            <div className="mt-5 space-y-1 text-xs text-tertiary">
              {message.diagnostics.map((diagnostic) => (
                <div key={diagnostic}>{diagnostic}</div>
              ))}
            </div>
          ) : null}
        </section>
      </div>
    </main>
  );
}

function HeaderLine({ label, value }: { readonly label: string; readonly value: string }) {
  return (
    <div>
      <dt className="text-xs uppercase text-tertiary">{label}</dt>
      <dd className="mt-0.5 break-words text-foreground">{value || "-"}</dd>
    </div>
  );
}

function AttachmentRow({
  attachment,
  busy,
  onRoute,
}: {
  readonly attachment: EmailAttachment;
  readonly busy: boolean;
  readonly onRoute: () => void;
}) {
  const { t } = useTranslator();
  return (
    <div className="rounded-md border border-divider p-2 text-sm">
      <div className="truncate font-medium">{attachment.filename}</div>
      <div className="mt-0.5 text-xs text-tertiary">{formatBytes(attachment.size)}</div>
      <Button
        className="mt-2 w-full justify-center"
        variant="secondary"
        size="sm"
        disabled={!attachment.routeFormat || busy}
        onClick={onRoute}
      >
        {busy ? <Loader2 size={14} className="mr-1.5 animate-spin" /> : null}
        {t("emailViewer.openAttachment")}
        <ArrowRight size={14} className="ml-1.5" />
      </Button>
    </div>
  );
}

function formatBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

function arrayBufferFromBytes(bytes: Uint8Array): ArrayBuffer {
  const copy = new Uint8Array(bytes.byteLength);
  copy.set(bytes);
  return copy.buffer;
}

export default function EmailViewerPage(): React.ReactNode {
  return (
    <Suspense fallback={<LoadingScreen variant="splash" product="docx" show />}>
      <EmailViewerPageInner />
    </Suspense>
  );
}
