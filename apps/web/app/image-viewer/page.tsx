"use client";

import Link from "next/link";
import { useSearchParams } from "next/navigation";
import { Suspense, useEffect, useMemo, useState } from "react";
import { ArrowLeft, Download, FileImage, Loader2, RefreshCw } from "@officeai/ui/sonaloop-icons";
import { Button } from "@officeai/ui";
import { useTranslator } from "@/lib/i18n";
import { LoadingScreen } from "@/lib/shell";
import { inspectorHrefForDocumentId } from "@/lib/sessions/editor-routing";
import { SessionEditorLoadState, useSessionEditorDocument } from "@/lib/sessions/session-editor-client";
import type { ImageViewerSource } from "@/lib/viewers/image-normalize";

function ImageViewerPageInner(): React.ReactNode {
  const params = useSearchParams();
  const documentId = params.get("session");
  const state = useSessionEditorDocument({ documentId, expectedFormat: "image" });
  if (!documentId || state.loading || state.error || !state.document) {
    return <SessionEditorLoadState format="image" state={state} />;
  }
  return <ImageViewer document={state.document} reload={state.reload} />;
}

function ImageViewer({
  document,
  reload,
}: {
  readonly document: NonNullable<ReturnType<typeof useSessionEditorDocument>["document"]>;
  readonly reload: () => Promise<void>;
}) {
  const { t } = useTranslator();
  const [source, setSource] = useState<ImageViewerSource | null>(null);
  const [objectUrl, setObjectUrl] = useState<string | null>(null);

  useEffect(() => {
    let cancelled = false;
    let nextUrl: string | null = null;
    void import("@/lib/viewers/image-normalize").then(async ({ normalizeImageForViewer }) => {
      const normalized = await normalizeImageForViewer(document.bytes, document.filename);
      if (cancelled) return;
      const blob = new Blob([arrayBufferFromBytes(normalized.displayBytes)], {
        type: normalized.browserRenderable ? normalized.mediaType : "image/svg+xml",
      });
      nextUrl = URL.createObjectURL(blob);
      setSource(normalized);
      setObjectUrl(nextUrl);
    });
    return () => {
      cancelled = true;
      if (nextUrl) URL.revokeObjectURL(nextUrl);
    };
  }, [document.bytes, document.filename]);

  const downloadHref = useMemo(() => {
    const blob = new Blob([arrayBufferFromBytes(document.bytes)], {
      type: source?.mediaType ?? "application/octet-stream",
    });
    return URL.createObjectURL(blob);
  }, [document.bytes, source?.mediaType]);

  useEffect(() => () => URL.revokeObjectURL(downloadHref), [downloadHref]);

  return (
    <main className="flex min-h-screen flex-col bg-background text-foreground">
      <header className="border-b border-divider bg-surface">
        <div className="mx-auto flex max-w-6xl flex-wrap items-center justify-between gap-3 px-6 py-4">
          <div className="min-w-0">
            <div className="flex items-center gap-2 text-xs uppercase text-tertiary">
              <FileImage size={14} />
              {t("imageViewer.title")}
            </div>
            <h1 className="mt-1 truncate text-xl font-semibold">{document.filename}</h1>
            <p className="mt-1 text-sm text-secondary">
              {source
                ? `${source.mediaType} · ${formatBytes(document.bytes.byteLength)}`
                : t("common.loading")}
            </p>
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
            <a href={downloadHref} download={document.filename}>
              <Button variant="ghost" size="sm">
                <Download size={14} className="mr-1.5" />
                {t("common.export")}
              </Button>
            </a>
            <Button variant="ghost" size="sm" onClick={() => void reload()}>
              <RefreshCw size={14} className="mr-1.5" />
              {t("home.refresh")}
            </Button>
          </div>
        </div>
      </header>

      <section className="mx-auto flex w-full max-w-6xl flex-1 flex-col gap-4 px-6 py-6">
        <div className="flex min-h-[62vh] items-center justify-center overflow-hidden rounded-lg border border-divider bg-surface p-4">
          {objectUrl && source ? (
            <img src={objectUrl} alt={document.filename} className="max-h-[72vh] max-w-full object-contain" />
          ) : (
            <div className="flex items-center gap-2 text-sm text-secondary">
              <Loader2 size={14} className="animate-spin" />
              {t("imageViewer.normalizing")}
            </div>
          )}
        </div>
        {source?.diagnostics.length ? (
          <div className="rounded-lg border border-divider bg-surface px-4 py-3 text-sm text-secondary">
            {source.diagnostics.map((diagnostic) => (
              <div key={diagnostic}>{diagnostic}</div>
            ))}
          </div>
        ) : null}
      </section>
    </main>
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

export default function ImageViewerPage(): React.ReactNode {
  return (
    <Suspense fallback={<LoadingScreen variant="splash" product="pdf" show />}>
      <ImageViewerPageInner />
    </Suspense>
  );
}
