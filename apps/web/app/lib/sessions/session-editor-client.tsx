"use client";

import Link from "next/link";
import { useCallback, useEffect, useRef, useState, type ReactNode } from "react";
import { AlertTriangle, Loader2, RefreshCw } from "@officeai/ui/sonaloop-icons";
import { Button } from "@officeai/ui";
import { useTranslator } from "@/lib/i18n";
import { inspectorHrefForDocumentId } from "./editor-routing";
import type { WebOfficeFormat } from "./web-sessions";

export interface SessionEditorDocument {
  readonly documentId: string;
  readonly bytes: Uint8Array;
  readonly filename: string;
  readonly revision: number;
  readonly etag: string;
  readonly format: WebOfficeFormat;
}

export interface SessionEditorState {
  readonly document: SessionEditorDocument | null;
  readonly loading: boolean;
  readonly error: string | null;
  readonly reload: () => Promise<void>;
  readonly save: (bytes: Uint8Array, mime: string, filename: string) => Promise<void>;
}

export function useSessionEditorDocument(args: {
  readonly documentId: string | null;
  readonly expectedFormat: WebOfficeFormat;
}): SessionEditorState {
  const { documentId, expectedFormat } = args;
  const [document, setDocument] = useState<SessionEditorDocument | null>(null);
  const [loading, setLoading] = useState(Boolean(documentId));
  const [error, setError] = useState<string | null>(null);
  const documentRef = useRef<SessionEditorDocument | null>(null);

  useEffect(() => {
    documentRef.current = document;
  }, [document]);

  const reload = useCallback(async () => {
    if (!documentId) {
      setDocument(null);
      setLoading(false);
      setError(null);
      return;
    }
    setLoading(true);
    setError(null);
    try {
      const res = await fetch(`/api/sessions/${encodeURIComponent(documentId)}/bytes`, {
        cache: "no-store",
      });
      if (!res.ok) {
        const data = (await res.json().catch(() => ({}))) as { readonly message?: string };
        throw new Error(data.message ?? `Session bytes failed (${res.status})`);
      }
      const format = res.headers.get("x-officeai-format") as WebOfficeFormat | null;
      if (format !== expectedFormat) {
        throw new Error(`Session ${documentId} is ${format ?? "unknown"}, not ${expectedFormat}.`);
      }
      const revision = Number.parseInt(res.headers.get("x-officeai-revision") ?? "", 10);
      if (!Number.isInteger(revision)) {
        throw new Error(`Session ${documentId} did not return a revision.`);
      }
      const filename =
        decodeHeaderFilename(res.headers.get("x-officeai-filename")) ?? `${documentId}.${format}`;
      const etag = res.headers.get("etag") ?? revisionEtag(documentId, revision);
      const bytes = new Uint8Array(await res.arrayBuffer());
      setDocument({ documentId, bytes, filename, revision, etag, format });
    } catch (err) {
      setDocument(null);
      setError(err instanceof Error ? err.message : String(err));
    } finally {
      setLoading(false);
    }
  }, [documentId, expectedFormat]);

  useEffect(() => {
    void reload();
  }, [reload]);

  const save = useCallback(
    async (bytes: Uint8Array, mime: string, filename: string) => {
      const current = documentRef.current;
      if (!documentId || !current) {
        throw new Error("No session document is loaded.");
      }
      const res = await fetch(`/api/sessions/${encodeURIComponent(documentId)}/bytes`, {
        method: "PUT",
        headers: {
          "content-type": mime,
          "if-match": current.etag,
          "x-officeai-base-revision": String(current.revision),
          "x-officeai-filename": encodeURIComponent(filename),
        },
        body: arrayBufferFromBytes(bytes),
      });
      if (!res.ok) {
        const data = (await res.json().catch(() => ({}))) as { readonly message?: string };
        throw new Error(data.message ?? `Session save failed (${res.status})`);
      }
      const data = (await res.json()) as {
        readonly document: {
          readonly documentId: string;
          readonly format: WebOfficeFormat;
          readonly name: string;
          readonly revision: number;
        };
        readonly etag?: string;
      };
      const next: SessionEditorDocument = {
        documentId: data.document.documentId,
        bytes,
        filename: data.document.name,
        revision: data.document.revision,
        etag: data.etag ?? revisionEtag(data.document.documentId, data.document.revision),
        format: data.document.format,
      };
      setDocument(next);
    },
    [documentId]
  );

  return { document, loading, error, reload, save };
}

export function SessionEditorLoadState({
  format,
  state,
}: {
  readonly format: WebOfficeFormat;
  readonly state: SessionEditorState;
}): ReactNode {
  const { t } = useTranslator();
  const documentId = state.document?.documentId;
  if (state.loading) {
    return (
      <div className="flex h-screen w-full items-center justify-center bg-background text-sm text-secondary">
        <div className="flex items-center gap-2">
          <Loader2 size={14} className="animate-spin" />
          {t("sessionEditor.loading")} {format.toUpperCase()}
        </div>
      </div>
    );
  }
  if (!state.error) return null;
  return (
    <main className="flex h-screen w-full items-center justify-center bg-background px-6">
      <section className="w-full max-w-lg rounded-lg border border-divider bg-surface p-6 text-sm">
        <div className="flex items-start gap-3">
          <AlertTriangle size={18} className="mt-0.5 text-amber-600 dark:text-amber-400" />
          <div className="min-w-0">
            <h1 className="text-base font-semibold text-foreground">{t("sessionEditor.loadError")}</h1>
            <p className="mt-2 break-words text-secondary">{state.error}</p>
            <div className="mt-5 flex flex-wrap gap-2">
              <Button variant="secondary" size="sm" onClick={() => void state.reload()}>
                <RefreshCw size={14} className="mr-1.5" />
                {t("home.refresh")}
              </Button>
              {documentId ? (
                <Link href={inspectorHrefForDocumentId(documentId)}>
                  <Button variant="ghost" size="sm">
                    {t("sessionEditor.openInspector")}
                  </Button>
                </Link>
              ) : null}
            </div>
          </div>
        </div>
      </section>
    </main>
  );
}

function decodeHeaderFilename(value: string | null): string | null {
  if (!value) return null;
  try {
    return decodeURIComponent(value);
  } catch {
    return value;
  }
}

function revisionEtag(documentId: string, revision: number): string {
  return `"officeai:${documentId}:${revision}"`;
}

function arrayBufferFromBytes(bytes: Uint8Array): ArrayBuffer {
  const copy = new Uint8Array(bytes.byteLength);
  copy.set(bytes);
  return copy.buffer;
}
