"use client";

import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from "react";
import { BookOpen, FolderOpen, MousePointer2 } from "lucide-react";
import { Button } from "@officeai/ui";
import { CommentsSidebar } from "@officeai/ui";
import { loadDocument, type PdfEngineDocument } from "@officeai/pdf-engine";
import { PdfAgent } from "@officeai/pdf/agent";
import type {
  AddAnnotationPayload as AddAnnotationInput,
  PdfRect,
  PdfRotation,
  PdfSnapshot,
} from "@officeai/pdf";
import { I18nProvider, useTranslator, type Locale } from "@/lib/i18n";
import { handleUndoRedo } from "@/lib/undo-redo";
import { useShortcutsDialog } from "@/lib/shortcuts/useShortcutsDialog";
import {
  EditorShell,
  ZoomControl,
  buildPaletteFromCatalogue,
  createToastId,
  type ExportFormat,
  type ExportOptionValues,
  type FindAdapter,
  type FindMatch,
  type PaletteCommand,
  type PaletteRunners,
  type ProductAdapter,
  type SaveState,
  type ToastItem,
} from "@/lib/shell";
import { pdfActions } from "@officeai/pdf";
import { downloadBlob, openFile, saveFile } from "@/lib/files/file-service";
import {
  PresenceSlot,
  RemotePresenceList,
  readExplicitRoomFromUrl,
  roomIdForSource,
  useCommandBroadcast,
  usePublishPresence,
  useRealtimeRoom,
  useStableTabId,
  type PresenceCursor,
} from "@/lib/realtime";
import { PdfToolbar, type PdfAnnotationTool } from "./PdfToolbar";
import { PdfCanvas, type PdfHighlight, type PdfViewMode } from "./PdfCanvas";
import { PdfSidebar, type PdfSidebarTab } from "./PdfSidebar";
import { usePdfShortcuts } from "./usePdfShortcuts";
import { usePdfFormatProvider } from "./usePdfFormatProvider";
import { usePdfCommentsProvider } from "./usePdfCommentsProvider";

/** PDF MIME / extension descriptor — `PRODUCT_FILE_TYPES` in the
 * shared file-service still tracks only the OOXML formats, so we
 * inline the PDF descriptor here (the values agree with what the
 * file-service consumes — one source-of-truth would land in the
 * follow-up that promotes pdf into the shared map). */
const PDF_MIME = "application/pdf";
const PDF_FILE_DESC = {
  description: "PDF document",
  mimeToExt: { [PDF_MIME]: [".pdf"] } as Record<string, string[]>,
  accept: ".pdf,application/pdf",
};

const MIN_ZOOM = 0.25;
const MAX_ZOOM = 6;
const ZOOM_STEP = 1.25;

const clampZoom = (z: number): number => Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, z));

function stripPdfExtension(name: string): string {
  return name.replace(/\.pdf$/i, "");
}

/** Lazy one-shot for `pdfjs-dist`'s worker hookup. The default PDF.js
 * backend in `@officeai/pdf-engine` requires `workerSrc` to be set
 * before the first parse; we point it at the bundled worker via Vite-
 * style `new URL(...)` resolution so Next can fingerprint and serve
 * it from `_next/static/`. The promise is cached so concurrent loads
 * never race the assignment.
 *
 * Embedding hosts (e.g. `@officeai/react-editors` consumers) can opt
 * out of the bundler-resolved path by setting
 * `globalThis.__OFFICEAI_PDFJS_WORKER_SRC__` BEFORE the editor
 * loads — useful when `import.meta.url` resolves to a chunk URL the
 * host's static server doesn't serve (Vite hosts can't reliably
 * resolve a bare-specifier `new URL("pdfjs-dist/...", ...)` against
 * a pre-bundled chunk). The react-editors bundle wires this to a
 * versioned CDN URL automatically. */
type PdfjsHostOverrides = {
  __OFFICEAI_PDFJS_WORKER_SRC__?: string;
  __OFFICEAI_PDFJS_ASSETS_BASE__?: string;
};

let pdfjsWorkerSetup: Promise<void> | null = null;
async function ensurePdfjsWorker(): Promise<void> {
  if (pdfjsWorkerSetup) return pdfjsWorkerSetup;
  pdfjsWorkerSetup = (async (): Promise<void> => {
    const pdfjs = await import("pdfjs-dist/legacy/build/pdf.mjs");
    const override = (globalThis as unknown as PdfjsHostOverrides).__OFFICEAI_PDFJS_WORKER_SRC__;
    if (override) {
      pdfjs.GlobalWorkerOptions.workerSrc = override;
      return;
    }
    const url = new URL("pdfjs-dist/build/pdf.worker.min.mjs", import.meta.url);
    pdfjs.GlobalWorkerOptions.workerSrc = url.toString();
  })();
  return pdfjsWorkerSetup;
}

/**
 * Public asset base for PDF.js's CMap and standard-font payloads.
 * Mirrors the destination of `apps/web/scripts/copy-pdfjs-assets.mjs`.
 * Both the engine document and the headless agent's parser need
 * this — without it CJK PDFs fail to extract selectable text and
 * any PDF that references one of the 14 base fonts without
 * embedding it falls back to a hard-coded sans-serif (visible drift
 * in the text layer at every zoom).
 *
 * Embedding hosts override the default via
 * `globalThis.__OFFICEAI_PDFJS_ASSETS_BASE__` so they can point at a
 * CDN (or skip the assets entirely if they only ever open
 * Latin-script PDFs with embedded fonts).
 */
const DEFAULT_PDFJS_ASSETS_BASE = "/pdfjs/";
function resolvePdfjsAssetsBase(): string {
  const override = (globalThis as unknown as PdfjsHostOverrides).__OFFICEAI_PDFJS_ASSETS_BASE__;
  return override ?? DEFAULT_PDFJS_ASSETS_BASE;
}

export interface PdfEditorProps {
  /** Page-level splash listens on this — see `pdf-viewer/page.tsx`.
   * Stays `false` until the agent + engine document are both
   * mounted, then `true`. */
  readonly onBootstrapReady?: (ready: boolean) => void;
  /** Optional pre-loaded PDF (sample-files entry on the home page).
   * Fetched as bytes; `name` becomes the working filename. */
  readonly initialSource?: { readonly url: string; readonly name: string };
  /** Bootstrap with a one-page blank PDF (the rare "New PDF"
   * action). Ignored when `initialSource` is set. */
  readonly initialBlank?: boolean;
  /** Optional pre-loaded PDF bytes. When set, takes priority over
   * `initialSource` and `initialBlank` so embedding hosts can stream
   * a `Uint8Array` straight into the viewer. */
  readonly initialBytes?: Uint8Array;
  /** Filename to display + use on Save when `initialBytes` is set. */
  readonly initialFilename?: string;
  /** Host save handler. */
  readonly onSave?: (bytes: Uint8Array, mime: string, filename: string) => Promise<void>;
  /** Host close handler. */
  readonly onClose?: () => void;
  /** Override the i18n locale; mounts a self-contained
   * `<I18nProvider initialLocale={locale}>`. */
  readonly locale?: Locale;
  /** Theme override placeholder; wired in Phase 1. */
  readonly theme?: "light" | "dark";
  /** Realtime presence identity (host-supplied). When set, replaces
   * the default anonymous identity on the awareness payload so
   * presence chips show the authenticated user's real name. */
  readonly presenceUser?: { readonly id: string; readonly name: string; readonly color?: string };
  /** Explicit realtime room id (host-supplied). Pin two browsers
   * viewing the same PDF into the same room without coordinating
   * URLs. Pass `null` to disable realtime. */
  readonly room?: string | null;
  /** Hide the 📁 Open toolbar affordance. Set by embedded hosts that
   * own their document corpus — see
   * `EmbeddedEditorProps.hideLocalFileOpen` in
   * `@officeai/react-editors/contract`. */
  readonly hideLocalFileOpen?: boolean;
}

export function PdfEditor(props: PdfEditorProps = {}): ReactNode {
  const { locale } = props;
  if (locale !== undefined) {
    return (
      <I18nProvider initialLocale={locale}>
        <PdfEditorInner {...props} />
      </I18nProvider>
    );
  }
  return <PdfEditorInner {...props} />;
}

function PdfEditorInner({
  onBootstrapReady,
  initialSource,
  initialBlank,
  initialBytes,
  initialFilename,
  onSave: onSaveProp,
  onClose: onCloseProp,
  presenceUser,
  room: roomOverride,
  hideLocalFileOpen,
}: PdfEditorProps = {}): ReactNode {
  const { t } = useTranslator();
  const [agent, setAgent] = useState<PdfAgent | null>(null);
  const agentRef = useRef<PdfAgent | null>(null);
  const [engineDoc, setEngineDoc] = useState<PdfEngineDocument | null>(null);
  const engineDocRef = useRef<PdfEngineDocument | null>(null);
  const [ready, setReady] = useState(false);
  const [tick, setTick] = useState(0);
  const [toasts, setToasts] = useState<ToastItem[]>([]);
  const [docName, setDocName] = useState(
    initialFilename ?? initialSource?.name ?? (initialBlank || initialBytes ? "Untitled.pdf" : "document.pdf")
  );
  const [fileHandle, setFileHandle] = useState<FileSystemFileHandle | undefined>(undefined);
  const [saveState, setSaveState] = useState<SaveState>("saved");
  const [currentPage, setCurrentPage] = useState(1);
  const [jumpNonce, setJumpNonce] = useState(0);
  const [zoom, setZoom] = useState(0.95);
  const [viewMode, setViewMode] = useState<PdfViewMode>("continuous");
  // Default to fit-page on every fresh document load. We don't
  // know the actual fit ratio until the canvas reports back via
  // `onZoomMetricsChange` (it depends on container × page size),
  // so we set this flag in `mountAgent` and consume it on the
  // first metrics callback after load.
  const pendingFitOnLoadRef = useRef(true);
  const [viewportRotation, setViewportRotation] = useState<PdfRotation>(0);
  const [sidebarTab, setSidebarTab] = useState<PdfSidebarTab>("thumbnails");
  const [highlight, setHighlight] = useState<PdfHighlight | null>(null);
  const shortcutsDialog = useShortcutsDialog();

  // The find-replace panel renders matches incrementally — keeping
  // the latest results around lets `gotoMatch` flash the right
  // highlight rect without re-running the search.
  const lastSearchResultsRef = useRef<Map<string, { pageNumber: number; rects: ReadonlyArray<PdfRect> }>>(
    new Map()
  );

  useEffect(() => {
    onBootstrapReady?.(ready);
  }, [ready, onBootstrapReady]);

  const pushToast = useCallback((kind: ToastItem["kind"], text: string) => {
    const id = createToastId("pdf");
    setToasts((prev) => [...prev, { id, kind, text }]);
  }, []);
  const dismissToast = useCallback((id: string) => {
    setToasts((prev) => prev.filter((toast) => toast.id !== id));
  }, []);
  const pushToastRef = useRef(pushToast);
  useEffect(() => {
    pushToastRef.current = pushToast;
  }, [pushToast]);

  const onError = useCallback(
    (err: unknown) => pushToast("error", err instanceof Error ? err.message : String(err)),
    [pushToast]
  );

  const teardownEngineDoc = useCallback(async () => {
    const prev = engineDocRef.current;
    engineDocRef.current = null;
    if (prev) {
      try {
        await prev.destroy();
      } catch {
        // Engine teardown errors are non-fatal — the next mount
        // will spin up a new document handle either way.
      }
    }
  }, []);

  const mountAgent = useCallback(
    async (buf: Uint8Array): Promise<void> => {
      await ensurePdfjsWorker();
      // Two parallel loads from the same buffer — one builds the
      // headless agent (parser + command bus), the other spins up
      // the rendering engine. We clone before each consumer so the
      // PDF.js worker can transfer the underlying ArrayBuffer
      // without detaching the pristine copy our agent / engine
      // each retain.
      const agentBuf = new Uint8Array(buf.byteLength);
      agentBuf.set(buf);
      const engineBuf = new Uint8Array(buf.byteLength);
      engineBuf.set(buf);
      const assetsBase = resolvePdfjsAssetsBase();
      const [nextAgent, nextEngine] = await Promise.all([
        PdfAgent.fromBuffer(agentBuf, { assetsBase }),
        loadDocument(engineBuf, { assetsBase }),
      ]);
      await teardownEngineDoc();
      agentRef.current = nextAgent;
      engineDocRef.current = nextEngine;
      setAgent(nextAgent);
      setEngineDoc(nextEngine);
      setCurrentPage(1);
      // Provisional zoom — gets overwritten by fit-page once the
      // canvas reports its first layout metrics (see ref above).
      setZoom(1);
      pendingFitOnLoadRef.current = true;
      setViewportRotation(0);
      setHighlight(null);
      setSaveState("saved");
      setTick((n) => n + 1);
      setReady(true);
      nextAgent.subscribe((_snap, mutation) => {
        setTick((n) => n + 1);
        if (mutation.status === "rejected" && mutation.rejection?.code === "rebase-failed") {
          pushToastRef.current(
            "warn",
            `An agent suggestion couldn't be re-applied after the last edit (${mutation.rejection.message})`
          );
          return;
        }
        setSaveState("modified");
      });
    },
    [teardownEngineDoc]
  );

  // Bootstrap: fetch initial source / build blank / leave empty. We
  // intentionally don't auto-build a "welcome PDF" the way PPTX
  // does — direct navigation lands on the EmptyState, matching the
  // home page's "PDF is open-only" stance.
  useEffect(() => {
    let cancelled = false;
    void (async () => {
      try {
        if (initialBytes) {
          const copy = new Uint8Array(initialBytes.byteLength);
          copy.set(initialBytes);
          if (!cancelled) await mountAgent(copy);
        } else if (initialSource) {
          const res = await fetch(initialSource.url);
          if (!res.ok) {
            throw new Error(`Failed to load ${initialSource.name} (${res.status})`);
          }
          const buf = new Uint8Array(await res.arrayBuffer());
          if (!cancelled) await mountAgent(buf);
        } else if (initialBlank) {
          const blank = await PdfAgent.empty();
          const bytes = await blank.exportFile();
          if (!cancelled) await mountAgent(bytes);
        } else {
          // No bootstrap → empty state. We still flip `ready` so
          // the splash retracts and the user sees the EmptyState
          // CTA instead of an indefinite spinner.
          if (!cancelled) setReady(true);
        }
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
        if (!cancelled) setReady(true);
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [initialBlank, initialBytes, initialSource, mountAgent, pushToast]);

  useEffect(() => {
    return () => {
      void teardownEngineDoc();
    };
  }, [teardownEngineDoc]);

  // Cross-product undo/redo: Cmd/Ctrl+Z and Cmd/Ctrl+Shift+Z route
  // through the shared chord detector so the keystroke means the
  // same thing inside any office editor in the suite.
  useEffect(() => {
    const onKey = (e: KeyboardEvent): void => {
      handleUndoRedo(e, agentRef.current);
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);

  const snap: PdfSnapshot | null = agent?.getSnapshot() ?? null;
  void tick;
  const pages = snap?.root.pages ?? [];
  const totalPages = pages.length;

  // Keep the active page clamped — if a page-delete command shrinks
  // the document under our feet, this saves the toolbar from
  // showing "Page 47 / 12" for a frame.
  useEffect(() => {
    if (totalPages === 0) {
      if (currentPage !== 1) setCurrentPage(1);
      return;
    }
    if (currentPage > totalPages) setCurrentPage(totalPages);
    if (currentPage < 1) setCurrentPage(1);
  }, [currentPage, totalPages]);

  // ── File ops ──────────────────────────────────────────────────────
  const handleFile = useCallback(
    async (file: File, handle?: FileSystemFileHandle) => {
      try {
        const buf = new Uint8Array(await file.arrayBuffer());
        setDocName(file.name);
        setFileHandle(handle);
        await mountAgent(buf);
        pushToast("info", `Opened ${file.name}`);
      } catch (err) {
        onError(err);
      }
    },
    [mountAgent, onError, pushToast]
  );

  const handleOpenFile = useCallback(async () => {
    try {
      const opened = await openFile({
        description: PDF_FILE_DESC.description,
        mimeToExt: PDF_FILE_DESC.mimeToExt,
        accept: PDF_FILE_DESC.accept,
      });
      if (!opened) return;
      const file = new File([opened.bytes as BlobPart], opened.name, {
        type: PDF_MIME,
      });
      await handleFile(file, opened.handle);
    } catch (err) {
      onError(err);
    }
  }, [handleFile, onError]);

  const handleSave = useCallback(async () => {
    const a = agentRef.current;
    if (!a) return;
    try {
      setSaveState("saving");
      const buf = await a.exportFile();
      const bytes = new Uint8Array(buf);
      if (onSaveProp) {
        await onSaveProp(bytes, PDF_MIME, docName);
        setSaveState("saved");
        pushToast("success", `Saved ${docName}`);
        return;
      }
      const inPlace = await saveFile(bytes, docName, PDF_MIME, fileHandle);
      setSaveState("saved");
      pushToast("success", inPlace ? `Saved ${docName}` : `Downloaded ${docName}`);
    } catch (err) {
      setSaveState("error");
      onError(err);
    }
  }, [docName, fileHandle, onError, onSaveProp, pushToast]);

  const exportFormats = usePdfFormatProvider();

  const handleExport = useCallback(
    async (format: ExportFormat, _options?: ExportOptionValues) => {
      const a = agentRef.current;
      if (!a) return;
      const baseName = stripPdfExtension(docName);
      const downloadName = `${baseName}.${format.extension}`;
      try {
        switch (format.id) {
          case "pdf": {
            const bytes = await a.exportFile();
            await saveFile(new Uint8Array(bytes), downloadName, format.mime, undefined);
            break;
          }
          case "markdown": {
            const md = a.toMarkdown();
            downloadBlob(new Blob([md], { type: format.mime }), downloadName);
            break;
          }
          case "text": {
            const text = a
              .getSnapshot()
              .root.pages.map((p) => p.text)
              .join("\n\n");
            downloadBlob(new Blob([text], { type: format.mime }), downloadName);
            break;
          }
          default:
            throw new Error(`Unsupported export format: ${format.id}`);
        }
        pushToast("success", `Exported ${downloadName}`);
      } catch (err) {
        onError(err);
        throw err;
      }
    },
    [docName, onError, pushToast]
  );

  // ── Toolbar callbacks ─────────────────────────────────────────────
  const goToPage = useCallback(
    (n: number) => {
      if (totalPages === 0) return;
      const clamped = Math.max(1, Math.min(totalPages, n));
      setCurrentPage(clamped);
      setJumpNonce((nonce) => nonce + 1);
    },
    [totalPages]
  );
  const onPrevPage = useCallback(() => goToPage(currentPage - 1), [currentPage, goToPage]);
  const onNextPage = useCallback(() => goToPage(currentPage + 1), [currentPage, goToPage]);
  const onFirstPage = useCallback(() => goToPage(1), [goToPage]);
  const onLastPage = useCallback(() => goToPage(totalPages || 1), [goToPage, totalPages]);

  const zoomMetricsRef = useRef<{ actualSizeZoom: number; fitPageZoom: number }>({
    actualSizeZoom: 1,
    fitPageZoom: 0.95,
  });
  const onZoomMetricsChange = useCallback((metrics: { actualSizeZoom: number; fitPageZoom: number }) => {
    zoomMetricsRef.current = metrics;
    if (pendingFitOnLoadRef.current) {
      pendingFitOnLoadRef.current = false;
      setZoom(clampZoom(metrics.fitPageZoom));
    }
  }, []);

  const onZoomIn = useCallback(() => setZoom((z) => clampZoom(z * ZOOM_STEP)), []);
  const onZoomOut = useCallback(() => setZoom((z) => clampZoom(z / ZOOM_STEP)), []);
  const onFitWidth = useCallback(() => setZoom(1), []);
  const onFitPage = useCallback(() => setZoom(clampZoom(zoomMetricsRef.current.fitPageZoom)), []);
  const onActualSize = useCallback(() => setZoom(clampZoom(zoomMetricsRef.current.actualSizeZoom)), []);
  const onSetZoom = useCallback((next: number) => setZoom(clampZoom(next)), []);

  const onRotateClockwise = useCallback(
    () => setViewportRotation((r) => ((r + 90) % 360) as PdfRotation),
    []
  );
  const onRotateCounterClockwise = useCallback(
    () => setViewportRotation((r) => ((r + 270) % 360) as PdfRotation),
    []
  );

  const printIframeRef = useRef<HTMLIFrameElement | null>(null);
  const printObjectUrlRef = useRef<string | null>(null);
  const onPrint = useCallback(async () => {
    const a = agentRef.current;
    if (!a) return;
    try {
      const bytes = await a.exportFile();
      const blob = new Blob([bytes as BlobPart], { type: "application/pdf" });
      const previousUrl = printObjectUrlRef.current;
      const url = URL.createObjectURL(blob);
      printObjectUrlRef.current = url;
      let iframe = printIframeRef.current;
      if (!iframe) {
        iframe = document.createElement("iframe");
        iframe.style.position = "fixed";
        iframe.style.right = "0";
        iframe.style.bottom = "0";
        iframe.style.width = "0";
        iframe.style.height = "0";
        iframe.style.border = "0";
        iframe.setAttribute("aria-hidden", "true");
        iframe.setAttribute("data-testid", "pdf-print-iframe");
        document.body.appendChild(iframe);
        printIframeRef.current = iframe;
      }
      iframe.onload = () => {
        try {
          iframe?.contentWindow?.focus();
          iframe?.contentWindow?.print();
        } catch (err) {
          // Some browsers block print() on cross-origin or sandboxed
          // iframes — fall back to opening the blob in a new tab so
          // the user can hit Cmd+P themselves.
          window.open(url, "_blank", "noopener,noreferrer");
          if (err instanceof Error) onError(err);
        }
      };
      iframe.src = url;
      if (previousUrl) {
        // Defer revoking the previous URL until after the new one
        // has loaded so the iframe doesn't snap to about:blank
        // mid-transition.
        setTimeout(() => URL.revokeObjectURL(previousUrl), 60_000);
      }
    } catch (err) {
      onError(err);
    }
  }, [onError]);
  useEffect(() => {
    return () => {
      const url = printObjectUrlRef.current;
      if (url) URL.revokeObjectURL(url);
      const iframe = printIframeRef.current;
      if (iframe?.parentNode) iframe.parentNode.removeChild(iframe);
    };
  }, []);

  const [armedTool, setArmedTool] = useState<PdfAnnotationTool | null>(null);
  const onAnnotate = useCallback((tool: PdfAnnotationTool) => {
    setArmedTool((current) => (current === tool ? null : tool));
  }, []);
  const unarmTool = useCallback(() => setArmedTool(null), []);

  // Esc cancels any armed annotation tool. Cheap UX safety net so
  // the user can never get "stuck" in highlight or sticky mode.
  useEffect(() => {
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === "Escape") setArmedTool((current) => (current ? null : current));
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);

  // Cmd/Ctrl+P prints the current PDF (with session annotations).
  // We intercept the system print dialog so the bytes that go to
  // the printer are the freshly-serialised PDF — not a screenshot
  // of the canvas, which would lose vector quality and any text.
  useEffect(() => {
    const onKey = (e: KeyboardEvent): void => {
      if (e.key.toLowerCase() !== "p") return;
      if (!(e.metaKey || e.ctrlKey)) return;
      if (e.shiftKey || e.altKey) return;
      e.preventDefault();
      void onPrint();
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [onPrint]);

  const onAddAnnotation = useCallback(
    async (input: AddAnnotationInput) => {
      const a = agentRef.current;
      if (!a) return;
      try {
        await a.applyCommand({
          type: "pdf:add-annotation",
          payload: input,
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [onError]
  );

  const onRemoveAnnotation = useCallback(
    async (annotationId: string) => {
      const a = agentRef.current;
      if (!a) return;
      try {
        await a.applyCommand({
          type: "pdf:remove-annotation",
          payload: { annotationId },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [onError]
  );

  const onUpdateAnnotation = useCallback(
    async (annotationId: string, contents: string) => {
      const a = agentRef.current;
      if (!a) return;
      try {
        await a.applyCommand({
          type: "pdf:update-annotation",
          payload: { annotationId, contents },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [onError]
  );

  const onRotatePages = useCallback(async () => {
    const a = agentRef.current;
    if (!a || totalPages === 0) return;
    try {
      await a.applyCommand({
        type: "pdf:rotate-pages",
        payload: { pageNumbers: [currentPage], rotation: 90, mode: "delta" },
        source: "human",
      });
    } catch (err) {
      onError(err);
    }
  }, [currentPage, onError, totalPages]);

  const onDeletePages = useCallback(async () => {
    const a = agentRef.current;
    if (!a || totalPages <= 1) {
      pushToast("warn", "Cannot delete the last remaining page.");
      return;
    }
    try {
      await a.applyCommand({
        type: "pdf:delete-pages",
        payload: { pageNumbers: [currentPage] },
        source: "human",
      });
      setCurrentPage((p) => Math.max(1, Math.min(totalPages - 1, p)));
    } catch (err) {
      onError(err);
    }
  }, [currentPage, onError, pushToast, totalPages]);

  // ── Search / find-replace ─────────────────────────────────────────
  const findAdapter = useMemo<FindAdapter | undefined>(() => {
    if (!agent) return undefined;
    return {
      findAll(query, opts) {
        const a = agentRef.current;
        if (!a || query.length === 0) {
          lastSearchResultsRef.current = new Map();
          return [];
        }
        try {
          const results = a.search({
            query,
            ...(opts.caseSensitive ? { caseSensitive: true } : {}),
            ...(opts.regex ? { regex: true } : {}),
          });
          const out: FindMatch[] = [];
          const map = new Map<string, { pageNumber: number; rects: ReadonlyArray<PdfRect> }>();
          for (const r of results) {
            const id = `${r.pageNumber}:${r.start}:${r.end}`;
            map.set(id, { pageNumber: r.pageNumber, rects: r.rects });
            out.push({ id, preview: r.preview });
            if (out.length >= 5000) break;
          }
          lastSearchResultsRef.current = map;
          return out;
        } catch {
          // Invalid regex / unsupported pattern — degrade quietly so
          // the panel just shows "no matches" instead of a toast.
          return [];
        }
      },
      gotoMatch(match) {
        const meta = lastSearchResultsRef.current.get(match.id);
        if (!meta) return;
        goToPage(meta.pageNumber);
        setHighlight({
          pageNumber: meta.pageNumber,
          // Glyph-precise quads when the structured-text pass found
          // them; falls back to a soft top-of-page pulse for hits
          // that span paragraphs / scanned pages where there's no
          // structured layer to project onto.
          rect: [0.05, 0.05, 0.95, 0.12],
          ...(meta.rects.length > 0 ? { quads: meta.rects } : {}),
          nonce: Date.now(),
        });
      },
      replaceMatch() {
        // PDF replace is a downstream command (W8). We keep find
        // working in the meantime so the keyboard shortcut
        // continues to do something useful.
      },
      replaceAll() {
        return 0;
      },
    };
  }, [agent, goToPage]);

  // ── Sidebar callbacks ─────────────────────────────────────────────
  const onJumpToPage = useCallback((n: number) => goToPage(n), [goToPage]);

  const onJumpToComment = useCallback(
    (commentId: string) => {
      const a = agentRef.current;
      if (!a) return;
      const c = a.getSnapshot().root.comments.find((cm) => cm.id === commentId);
      if (!c) return;
      goToPage(c.pageNumber);
      setHighlight({ pageNumber: c.pageNumber, rect: c.normalizedRect, nonce: Date.now() });
    },
    [goToPage]
  );

  // ── Keyboard shortcuts ────────────────────────────────────────────
  usePdfShortcuts({
    nextPage: onNextPage,
    prevPage: onPrevPage,
    firstPage: onFirstPage,
    lastPage: onLastPage,
    zoomIn: onZoomIn,
    zoomOut: onZoomOut,
    fitWidth: onFitWidth,
    fitPage: onFitPage,
    actualSize: onActualSize,
    rotateClockwise: onRotateClockwise,
    rotateCounterClockwise: onRotateCounterClockwise,
  });

  // ── Comments provider ─────────────────────────────────────────────
  const commentsProvider = usePdfCommentsProvider({
    agent: agent as PdfAgent,
    currentPage,
    onScrollTo: onJumpToComment,
  });

  const renderCommentsPanel = useCallback((): ReactNode => {
    if (!agent) {
      return <div className="p-4 text-sm text-secondary">{t("pdf.addCommentHint")}</div>;
    }
    return (
      <div className="flex min-h-0 flex-1 flex-col overflow-y-auto">
        <CommentsSidebar provider={commentsProvider} onScrollTo={onJumpToComment} />
      </div>
    );
  }, [agent, commentsProvider, onJumpToComment, t]);

  // Comments badge for the top bar.
  const openCommentCount = useMemo(() => {
    if (!snap) return 0;
    return snap.root.comments.filter((c) => !c.parentId && !c.resolved).length;
  }, [snap]);

  // ── Selection / status bar ────────────────────────────────────────
  const selectionText = useMemo(() => {
    if (!ready || totalPages === 0) return "";
    return t("pdf.pageOf", { n: currentPage, total: totalPages });
  }, [currentPage, ready, t, totalPages]);

  // ── Palette commands ──────────────────────────────────────────────
  // Generated from the central pdf action catalogue (see
  // packages/pdf/src/actions/catalogue.ts). Labels/sections flow from
  // the catalogue; this map only carries the closure-bound side
  // effects + per-id `enabled` gating.
  const paletteCommands = useMemo<ReadonlyArray<PaletteCommand>>(() => {
    const runners: PaletteRunners = {
      "pdf.next-page": { run: onNextPage, enabled: currentPage < totalPages },
      "pdf.prev-page": { run: onPrevPage, enabled: currentPage > 1 },
      "pdf.first-page": { run: onFirstPage, enabled: totalPages > 0 },
      "pdf.last-page": { run: onLastPage, enabled: totalPages > 0 },
      "pdf.zoom-in": { run: onZoomIn },
      "pdf.zoom-out": { run: onZoomOut },
      "pdf.fit-width": { run: onFitWidth },
      "pdf.fit-page": { run: onFitPage },
      "pdf.actual-size": { run: onActualSize },
      "pdf.rotate-cw": { run: onRotateClockwise },
      "pdf.rotate-ccw": { run: onRotateCounterClockwise },
      "pdf.rotate-page": { run: () => void onRotatePages(), enabled: totalPages > 0 },
      "pdf.delete-page": { run: () => void onDeletePages(), enabled: totalPages > 1 },
      "pdf.print": { run: () => void onPrint(), enabled: totalPages > 0 },
    };
    return buildPaletteFromCatalogue(pdfActions, runners, t);
  }, [
    t,
    currentPage,
    onActualSize,
    onDeletePages,
    onFirstPage,
    onFitPage,
    onFitWidth,
    onLastPage,
    onNextPage,
    onPrevPage,
    onPrint,
    onRotateClockwise,
    onRotateCounterClockwise,
    onRotatePages,
    onZoomIn,
    onZoomOut,
    totalPages,
  ]);

  // ── Realtime ──────────────────────────────────────────────────────
  const tabFallback = useStableTabId("pdf");
  const realtimeRoomId = useMemo<string | null>(() => {
    if (!ready) return null;
    if (roomOverride === null) return null;
    if (typeof roomOverride === "string" && roomOverride.length > 0) {
      return `oai/pdf/host/${roomOverride}`;
    }
    if (!tabFallback && !initialSource) return null;
    return roomIdForSource({
      product: "pdf",
      src: initialSource?.url ?? null,
      tabFallback,
      explicitRoom: readExplicitRoomFromUrl(),
    });
  }, [ready, initialSource, tabFallback, roomOverride]);
  const realtimeRoom = useRealtimeRoom({
    roomId: realtimeRoomId,
    product: "pdf",
    ...(presenceUser
      ? {
          identity: {
            id: presenceUser.id,
            name: presenceUser.name,
            ...(presenceUser.color ? { color: presenceUser.color } : {}),
          },
        }
      : {}),
  });
  useCommandBroadcast({
    agent,
    room: realtimeRoom.room,
  });

  const presenceCursor = useMemo<PresenceCursor | null>(() => {
    if (!ready || totalPages === 0) return null;
    return {
      product: "pdf",
      pageNumber: currentPage,
    };
  }, [currentPage, ready, totalPages]);
  usePublishPresence({ room: realtimeRoom.room, cursor: presenceCursor });

  // ── Adapter ───────────────────────────────────────────────────────
  const adapter = useMemo<ProductAdapter>(
    () => ({
      product: "pdf",
      filename: docName,
      saveState,
      comments: { openCount: openCommentCount, resolvedCount: 0 },
      selectionSummary: { text: selectionText },
      canOpen: true,
      hideOpen: hideLocalFileOpen,
      canSave: ready && agent !== null,
      canExport: ready && agent !== null,
      exportFormats,
      onOpenFile: () => void handleOpenFile(),
      onSave: () => handleSave(),
      onExport: (format, options) => handleExport(format, options),
      canUndo: ready ? (agentRef.current?.canUndo() ?? false) : false,
      canRedo: ready ? (agentRef.current?.canRedo() ?? false) : false,
      onUndo: () => {
        const a = agentRef.current;
        if (!a || !a.canUndo()) return;
        a.undo();
      },
      onRedo: () => {
        const a = agentRef.current;
        if (!a || !a.canRedo()) return;
        a.redo();
      },
      onOpenShortcuts: () => shortcutsDialog.setOpen(true),
      paletteCommands,
      ...(findAdapter ? { findAdapter } : {}),
      renderCommentsPanel,
    }),
    [
      agent,
      docName,
      exportFormats,
      findAdapter,
      handleExport,
      handleOpenFile,
      handleSave,
      hideLocalFileOpen,
      openCommentCount,
      paletteCommands,
      ready,
      renderCommentsPanel,
      saveState,
      selectionText,
      shortcutsDialog,
      tick,
    ]
  );

  return (
    <>
      <RemotePresenceList peers={realtimeRoom.remotePeers} />
      <EditorShell
        adapter={adapter}
        onBack={onCloseProp}
        topBarExtras={<PresenceSlot state={realtimeRoom} />}
        toolbar={
          <PdfToolbar
            disabled={!ready || !agent}
            currentPage={currentPage}
            totalPages={totalPages}
            zoom={zoom}
            viewMode={viewMode}
            onPrevPage={onPrevPage}
            onNextPage={onNextPage}
            onJumpToPage={goToPage}
            onZoomIn={onZoomIn}
            onZoomOut={onZoomOut}
            onFitWidth={onFitWidth}
            onFitPage={onFitPage}
            onActualSize={onActualSize}
            onSetZoom={onSetZoom}
            onSetViewMode={setViewMode}
            onRotateClockwise={onRotateClockwise}
            onRotateCounterClockwise={onRotateCounterClockwise}
            onAnnotate={onAnnotate}
            armedTool={armedTool}
            onPrint={() => void onPrint()}
            onRotatePages={() => void onRotatePages()}
            onDeletePages={() => void onDeletePages()}
          />
        }
        statusBarLeft={<PdfStatusHint currentPage={currentPage} totalPages={totalPages} />}
        statusBarRight={
          <ZoomControl
            value={zoom}
            minZoom={MIN_ZOOM}
            maxZoom={MAX_ZOOM}
            onChange={(next) => setZoom(clampZoom(next))}
            disabled={!ready || !agent}
          />
        }
        body={
          <div className="pdf-editor flex min-h-0 flex-1">
            {!agent ? (
              <PdfEmptyState onOpen={() => void handleOpenFile()} />
            ) : (
              <div className="flex min-h-0 flex-1">
                <PdfSidebar
                  snapshot={snap}
                  engineDoc={engineDoc}
                  currentPage={currentPage}
                  viewportRotation={viewportRotation}
                  tab={sidebarTab}
                  onTabChange={setSidebarTab}
                  onJumpToPage={onJumpToPage}
                  onJumpToComment={onJumpToComment}
                />
                <section className="relative min-h-0 min-w-0 flex-1">
                  <PdfCanvas
                    engineDoc={engineDoc}
                    snapshot={snap}
                    currentPage={currentPage}
                    zoom={zoom}
                    viewMode={viewMode}
                    darkMode="off"
                    reflow={false}
                    viewportRotation={viewportRotation}
                    onCurrentPageChange={setCurrentPage}
                    onZoomMetricsChange={onZoomMetricsChange}
                    jumpNonce={jumpNonce}
                    highlight={highlight}
                    armedTool={armedTool}
                    onAddAnnotation={(input) => void onAddAnnotation(input)}
                    onAnnotationCreated={unarmTool}
                  />
                </section>
              </div>
            )}
          </div>
        }
        toasts={toasts}
        onDismissToast={dismissToast}
        onFileDrop={(file) => void handleFile(file)}
        dropExtension=".pdf"
        onRenameFilename={(next) => setDocName(next)}
      />
    </>
  );
}

interface PdfEmptyStateProps {
  readonly onOpen: () => void;
}

/**
 * Inline empty state for the PDF viewer.
 *
 * The shared `EmptyState` component currently only knows about the
 * OOXML editors (it builds its label / extension copy from a
 * three-entry record). Rather than reach across into shared shell
 * code from this scope, we render the same visual layout locally —
 * promoting it back into `EmptyState` is a follow-up.
 */
function PdfEmptyState({ onOpen }: PdfEmptyStateProps): ReactNode {
  const { t } = useTranslator();
  return (
    <div
      className="flex h-full w-full flex-col items-center justify-center gap-6 px-6 text-center"
      data-testid="pdf-empty-state"
    >
      <div className="flex flex-col items-center gap-2">
        <div className="inline-flex h-12 w-12 items-center justify-center rounded-full bg-hover text-secondary">
          <BookOpen size={22} />
        </div>
        <h2 className="text-lg font-semibold text-foreground">{t("pdf.title")}</h2>
        <p className="max-w-md text-sm text-secondary">{t("pdf.openSample")}</p>
      </div>
      <Button variant="primary" size="md" onClick={onOpen}>
        <FolderOpen size={14} />
        {t("common.open")}
      </Button>
      <div className="flex items-center gap-2 text-xs text-tertiary">
        <MousePointer2 size={12} />
        <span>{t("common.orDropHere", { ext: ".pdf" })}</span>
      </div>
    </div>
  );
}

interface PdfStatusHintProps {
  readonly currentPage: number;
  readonly totalPages: number;
}

function PdfStatusHint({ currentPage, totalPages }: PdfStatusHintProps): ReactNode {
  const { t } = useTranslator();
  if (totalPages === 0) return null;
  return (
    <span
      className="flex items-center gap-3 text-[11px] tabular-nums text-tertiary"
      data-testid="pdf-status-hint"
      aria-live="polite"
    >
      <span>{t("pdf.pageOf", { n: currentPage, total: totalPages })}</span>
    </span>
  );
}
