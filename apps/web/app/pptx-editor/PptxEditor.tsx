"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { CommentComposer, CommentsSidebar } from "@officeai/ui";
import { createPptxCommentsProvider } from "./pptxCommentsProvider";
import { PptxAgent } from "@officeai/pptx/agent";
import {
  SlideCanvas,
  SlidesSidebar,
  type PptxTextSelection,
  type SlideContextAction,
} from "@officeai/pptx/renderer/react";
import { MAX_ZOOM, MIN_ZOOM, clampZoom } from "@officeai/pptx/renderer";
import type { LayoutKindPayload, Picture, PptxSnapshot, Shape, ShapePreset, TextShape } from "@officeai/pptx";
import { buildSamplePptx } from "@/lib/sample-pptx";
import { PptxToolbar } from "./PptxToolbar";
import { ConnectorContextBar, type ConnectorStylePatch } from "./ConnectorContextBar";
import { PresentMode } from "./PresentMode";
import { AnimationsPanel } from "./AnimationsPanel";
import { computePptxActive, createPptxFormatProvider } from "./pptxFormatProvider";
import { useShortcutsDialog } from "@/lib/shortcuts/useShortcutsDialog";
import { KeyboardShortcutsDialog } from "@/lib/shortcuts/KeyboardShortcutsDialog";
import { usePptxShortcuts } from "@/lib/pptx-shortcuts";
import {
  EditorShell,
  EmptyState,
  LoadingScreen,
  ZoomControl,
  createToastId,
  type ExportFormat,
  type ExportOptionValues,
  type PaletteCommand,
  type ProductAdapter,
  type SaveState,
  type ToastItem,
} from "@/lib/shell";
import {
  PRODUCT_FILE_TYPES,
  downloadBlob,
  openFile,
  saveFile,
} from "@/lib/files/file-service";
import { convertViaServer } from "@/lib/files/convert-client";
import {
  parseSlideRange,
  snapshotToPngZip,
  snapshotToSlideJpeg,
  snapshotToSlidePng,
  snapshotToSlideSvg,
  snapshotToSvgZip,
} from "./lib/export-images";

const SCALE_OPTIONS = {
  type: "select" as const,
  defaultId: "2",
  options: [
    { id: "1", label: "1× (standard)" },
    { id: "2", label: "2× (retina)" },
    { id: "3", label: "3× (high-DPI)" },
  ],
};

const PPTX_EXPORT_FORMATS: ReadonlyArray<ExportFormat> = [
  // ── Whole deck ─────────────────────────────────────────────────
  // Deck-level exports lead the dropdown because "download the
  // .pptx / PDF the deck" is overwhelmingly the common ask. We
  // collapse what would otherwise be Native / PDF & web / Images
  // into a single bucket: in PPTX each holds only 1–2 entries, and
  // rendering them as separate sections reads as a chopped-up
  // dropdown. Order: native (round-trip-true) first, then PDF and
  // HTML, then the bundles (PNG before SVG to mirror the per-slide
  // ordering further down).
  {
    id: "pptx",
    label: "PowerPoint presentation (.pptx)",
    description: "Round-trip native OOXML — opens in PowerPoint, Keynote and LibreOffice Impress.",
    extension: "pptx",
    mime: PRODUCT_FILE_TYPES.pptx.primaryMime,
    kind: "instant",
    group: "deck",
    icon: "slides",
  },
  {
    id: "pdf",
    label: "PDF document (.pdf)",
    description: "Server-side conversion via LibreOffice. One slide per PDF page.",
    extension: "pdf",
    mime: "application/pdf",
    kind: "dialog",
    group: "deck",
    icon: "pdf",
    optionFields: [
      {
        id: "slideRange",
        label: "Slide range",
        control: { type: "text", placeholder: "All slides — try 1,3-5" },
        hint: "Leave blank for every slide. Examples: 1,3 — 2-5 — 1,4-7,10.",
      },
      {
        id: "pageSize",
        label: "Page size",
        control: {
          type: "select",
          defaultId: "source",
          options: [
            { id: "source", label: "Use slide size" },
            { id: "a4", label: "A4 landscape" },
            { id: "letter", label: "Letter landscape" },
          ],
        },
      },
    ],
  },
  {
    id: "html",
    label: "Web page (.html)",
    description: "Server-side HTML export. Ships an interactive viewer with each slide.",
    extension: "html",
    mime: "text/html",
    kind: "instant",
    group: "deck",
    icon: "code",
  },
  {
    id: "png-zip",
    label: "All slides — PNG (.zip)",
    description: "One PNG per slide, bundled as a zip. Pick a scale for retina exports.",
    extension: "zip",
    mime: "application/zip",
    kind: "dialog",
    group: "deck",
    icon: "image",
    optionFields: [
      { id: "scale", label: "Resolution", control: SCALE_OPTIONS },
      {
        id: "slideRange",
        label: "Slide range",
        control: { type: "text", placeholder: "All slides — try 1,3-5" },
        hint: "Leave blank for every slide.",
      },
    ],
  },
  {
    id: "svg-zip",
    label: "All slides — SVG (.zip)",
    description: "Resolution-independent SVG per slide.",
    extension: "zip",
    mime: "application/zip",
    kind: "instant",
    group: "deck",
    icon: "image",
  },
  // ── This slide ─────────────────────────────────────────────────
  // Special-case bucket at the bottom: "I just want this one slide
  // out". Order within the group is most-shareable → most-
  // specialised: PDF (universal single-page handout), PNG (default
  // raster for chat / docs), JPEG (smaller raster), SVG (vector
  // for designers).
  {
    id: "slide-pdf",
    label: "Current slide — PDF (.pdf)",
    description: "Single-page PDF of the current slide. Server-side via LibreOffice.",
    extension: "pdf",
    mime: "application/pdf",
    kind: "instant",
    group: "current",
    icon: "pdf",
  },
  {
    id: "slide-png",
    label: "Current slide — PNG (.png)",
    description: "Just the slide that's currently in view. Lossless, with a white matte.",
    extension: "png",
    mime: "image/png",
    kind: "dialog",
    group: "current",
    icon: "image",
    optionFields: [{ id: "scale", label: "Resolution", control: SCALE_OPTIONS }],
  },
  {
    id: "slide-jpeg",
    label: "Current slide — JPEG (.jpg)",
    description: "Smaller, share-friendly version of the current slide.",
    extension: "jpg",
    mime: "image/jpeg",
    kind: "dialog",
    group: "current",
    icon: "image",
    optionFields: [
      { id: "scale", label: "Resolution", control: SCALE_OPTIONS },
      {
        id: "quality",
        label: "Quality",
        control: {
          type: "select",
          defaultId: "high",
          options: [
            { id: "low", label: "Low (~60%)" },
            { id: "medium", label: "Medium (~75%)" },
            { id: "high", label: "High (~92%)" },
            { id: "max", label: "Maximum (~98%)" },
          ],
        },
      },
    ],
  },
  {
    id: "slide-svg",
    label: "Current slide — SVG (.svg)",
    description: "Resolution-independent vector of the current slide.",
    extension: "svg",
    mime: "image/svg+xml",
    kind: "instant",
    group: "current",
    icon: "image",
  },
];

function stripPptxExtension(name: string): string {
  return name.replace(/\.pptx$/i, "");
}

/** Coerce a dialog "scale" option (`"1" | "2" | "3"`) into the
 * narrowed numeric type the renderer accepts. Defaults to 2× to
 * match the export-zip default. */
function parseScale(raw: ExportOptionValues[string] | undefined): 1 | 2 | 3 {
  const s = typeof raw === "string" ? raw : "2";
  return s === "1" ? 1 : s === "3" ? 3 : 2;
}

/** Map the dialog's coarse quality picker to the 0-1 number canvas
 * wants. Mirrors the lossless-vs-share trade-off most users actually
 * care about; the in-between values come from PhotoShop's "Save for
 * web" presets. */
function parseJpegQuality(raw: ExportOptionValues[string] | undefined): number {
  switch (typeof raw === "string" ? raw : "high") {
    case "low":
      return 0.6;
    case "medium":
      return 0.75;
    case "max":
      return 0.98;
    case "high":
    default:
      return 0.92;
  }
}

const SUPPORTED_IMAGE_MIME: ReadonlySet<string> = new Set([
  "image/png",
  "image/jpeg",
  "image/jpg",
  "image/gif",
  "image/bmp",
  "image/webp",
  "image/svg+xml",
]);

export function PptxEditor(): React.ReactNode {
  const [agent, setAgent] = useState<PptxAgent | null>(null);
  const agentRef = useRef<PptxAgent | null>(null);
  const [ready, setReady] = useState(false);
  const [docName, setDocName] = useState("welcome.pptx");
  const [activeIndex, setActiveIndex] = useState(0);
  const [tick, setTick] = useState(0);
  const [toasts, setToasts] = useState<ToastItem[]>([]);
  const [zoom, setZoom] = useState(1);
  const [presenting, setPresenting] = useState(false);
  // Speaker-notes panel is hidden by default — matches PowerPoint's
  // "View ▸ Notes" opt-in. Users toggle it via the trailing toolbar
  // button; we deliberately don't persist this across sessions because
  // the panel takes meaningful vertical space and the default state is
  // what users expect when reopening a deck.
  const [notesOpen, setNotesOpen] = useState(false);
  const [selectedShapeIds, setSelectedShapeIds] = useState<ReadonlyArray<string>>([]);
  const selectedShapeId = selectedShapeIds[0] ?? null;
  const [textSelection, setTextSelection] = useState<PptxTextSelection | null>(null);
  // Connector tool mode: when set, the canvas surfaces ports on every
  // hovered shape and a press-drag gesture commits a brand-new
  // connector of the chosen type. PowerPoint / Google Slides model the
  // connector picker as a "tool" rather than an instant insert because
  // the user almost always wants to place the line themselves; the
  // toolbar item toggles this state, and `null` is the default
  // selection-mode behaviour. Cleared on commit, on Esc, and when the
  // user re-clicks the same toolbar item.
  const [connectorTool, setConnectorTool] = useState<{
    type: "straight" | "elbow" | "curved";
  } | null>(null);
  const [saveState, setSaveState] = useState<SaveState>("saved");
  const [fileHandle, setFileHandle] = useState<FileSystemFileHandle | undefined>(undefined);
  // Mutable mirrors used by the format provider so it can read live
  // state without triggering rerenders inside its own callbacks.
  const slideIndexRef = useRef(0);
  const textSelectionRef = useRef<PptxTextSelection | null>(null);
  const selectedShapeIdRef = useRef<string | null>(null);
  const pushToastRef = useRef<((kind: ToastItem["kind"], text: string) => void) | null>(null);
  const slideSurfaceRef = useRef<HTMLElement | null>(null);
  // Bumped each time the comments sidebar requests "scroll / focus on
  // this comment". The canvas keys its flash overlay off `nonce` so
  // re-clicking the same comment re-plays the animation.
  const [commentFlashTarget, setCommentFlashTarget] = useState<
    | { kind: "shape"; shapeId: string; nonce: number }
    | { kind: "pin"; xEmu: number; yEmu: number; nonce: number }
    | null
  >(null);
  const shortcutsDialog = useShortcutsDialog();

  const onZoomChange = useCallback((next: number) => {
    setZoom(clampZoom(next));
  }, []);

  const pushToast = useCallback((kind: ToastItem["kind"], text: string) => {
    const id = createToastId("pptx");
    setToasts((prev) => [...prev, { id, kind, text }]);
  }, []);
  const dismissToast = useCallback((id: string) => {
    setToasts((prev) => prev.filter((t) => t.id !== id));
  }, []);
  useEffect(() => {
    pushToastRef.current = pushToast;
  }, [pushToast]);
  // These refs back the format provider's call-time reads. We sync them
  // *during render* (rather than in `useEffect`) so the provider sees
  // the freshest state on the very first render after a state change —
  // otherwise the toolbar's `provider.hasSelection()` evaluates against
  // a stale ref and the controls render disabled for one frame and
  // never re-render to recover (refs don't trigger renders themselves).
  slideIndexRef.current = activeIndex;
  textSelectionRef.current = textSelection;
  selectedShapeIdRef.current = selectedShapeId;

  /* eslint-disable react-hooks/refs -- the provider stores refs for
     call-time reads; it never dereferences them during construction. */
  const [textFormatProvider] = useState(() =>
    createPptxFormatProvider({
      agentRef,
      slideIndexRef,
      selectionRef: textSelectionRef,
      selectedShapeIdRef,
      pushToast: (kind, text) => pushToastRef.current?.(kind, text),
    })
  );
  /* eslint-enable react-hooks/refs */

  const mountAgent = useCallback(async (buf: ArrayBuffer) => {
    const next = await PptxAgent.fromBuffer(buf);
    agentRef.current = next;
    setAgent(next);
    setActiveIndex(0);
    setSelectedShapeIds([]);
    setReady(true);
    setTick((t) => t + 1);
    setSaveState("saved");
    next.subscribe(() => {
      setTick((t) => t + 1);
      setSaveState("modified");
    });
  }, []);

  useEffect(() => {
    let cancelled = false;
    void (async () => {
      try {
        const buf = await buildSamplePptx();
        if (!cancelled) await mountAgent(buf);
      } catch (err) {
        pushToast("error", err instanceof Error ? err.message : String(err));
      }
    })();
    return () => {
      cancelled = true;
    };
  }, [mountAgent, pushToast]);

  const onError = useCallback(
    (err: unknown) => pushToast("error", err instanceof Error ? err.message : String(err)),
    [pushToast]
  );

  // Stash the latest duplicate handler in a ref so the keyboard
  // listener doesn't have to re-bind every time `selectedShapeIds`
  // changes (which would also create a temporal-dead-zone here, since
  // `duplicateSelectedShapes` is declared further down).
  const duplicateRef = useRef<() => void>(() => {});
  const groupRef = useRef<() => void>(() => {});
  const ungroupRef = useRef<() => void>(() => {});

  // D1 — Undo / Redo keyboard shortcuts (Cmd/Ctrl+Z, Cmd/Ctrl+Shift+Z,
  // Cmd/Ctrl+Y). We bind on document so they fire regardless of which
  // surface element is focused, and skip when the user is typing inside
  // a real text input/textarea/contenteditable so we don't fight the
  // browser's native undo for those fields.
  useEffect(() => {
    const isFormField = (target: EventTarget | null): boolean => {
      if (!(target instanceof HTMLElement)) return false;
      const tag = target.tagName;
      if (tag === "INPUT" || tag === "TEXTAREA") return true;
      if (target.isContentEditable) return true;
      return false;
    };
    const onKey = (e: KeyboardEvent) => {
      const mod = e.metaKey || e.ctrlKey;
      if (!mod) return;
      const key = e.key.toLowerCase();
      const isUndo = key === "z" && !e.shiftKey;
      const isRedo = (key === "z" && e.shiftKey) || key === "y";
      const isDuplicate = key === "d" && !e.shiftKey && !e.altKey;
      const isUngroup = key === "g" && e.shiftKey && e.altKey;
      const isGroup = key === "g" && e.shiftKey && !e.altKey;
      if (!isUndo && !isRedo && !isDuplicate && !isGroup && !isUngroup) return;
      if (isFormField(e.target)) return;
      const a = agentRef.current;
      if (!a) return;
      e.preventDefault();
      if (isUndo) {
        if (a.canUndo()) a.undo();
      } else if (isRedo) {
        if (a.canRedo()) a.redo();
      } else if (isDuplicate) {
        duplicateRef.current();
      } else if (isUngroup) {
        ungroupRef.current();
      } else if (isGroup) {
        groupRef.current();
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, []);

  // D3 — Tab cycles selection across the slide's top-level shapes. With
  // Shift it cycles backward. Skipped when typing inside a real form
  // field so the browser's native tab-to-focus still works.
  useEffect(() => {
    const isFormField = (target: EventTarget | null): boolean => {
      if (!(target instanceof HTMLElement)) return false;
      const tag = target.tagName;
      if (tag === "INPUT" || tag === "TEXTAREA" || tag === "SELECT") return true;
      if (target.isContentEditable) return true;
      return false;
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key !== "Tab") return;
      if (e.metaKey || e.ctrlKey || e.altKey) return;
      if (isFormField(e.target)) return;
      const a = agentRef.current;
      if (!a) return;
      const slide = a.getSnapshot().root.slides[activeIndex];
      if (!slide || slide.shapes.length === 0) return;
      e.preventDefault();
      const shapes = slide.shapes;
      const currentId = selectedShapeIds[selectedShapeIds.length - 1] ?? null;
      const currentIdx = currentId ? shapes.findIndex((s) => s.id === currentId) : -1;
      const dir = e.shiftKey ? -1 : 1;
      const len = shapes.length;
      const startIdx = currentIdx === -1 ? (e.shiftKey ? len - 1 : 0) : (currentIdx + dir + len) % len;
      setSelectedShapeIds([shapes[startIdx]!.id]);
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [activeIndex, selectedShapeIds]);

  const handleFile = useCallback(
    async (file: File, handle?: FileSystemFileHandle) => {
      const buf = await file.arrayBuffer();
      setDocName(file.name);
      setFileHandle(handle);
      try {
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
        description: PRODUCT_FILE_TYPES.pptx.description,
        mimeToExt: PRODUCT_FILE_TYPES.pptx.mimeToExt,
        accept: PRODUCT_FILE_TYPES.pptx.accept,
      });
      if (!opened) return;
      const file = new File([opened.bytes as BlobPart], opened.name, {
        type: PRODUCT_FILE_TYPES.pptx.primaryMime,
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
      const inPlace = await saveFile(
        new Uint8Array(buf),
        docName,
        PRODUCT_FILE_TYPES.pptx.primaryMime,
        fileHandle
      );
      setSaveState("saved");
      pushToast("success", inPlace ? `Saved ${docName}` : `Downloaded ${docName}`);
    } catch (err) {
      setSaveState("error");
      onError(err);
    }
  }, [docName, fileHandle, onError, pushToast]);

  const handleExport = useCallback(
    async (format: ExportFormat, options?: ExportOptionValues) => {
      const a = agentRef.current;
      if (!a) return;
      const baseName = stripPptxExtension(docName);
      // Single-slide exports get a `-slideN` suffix so a "current
      // slide" PNG can't silently overwrite a previous deck-level
      // export sitting in the user's Downloads folder.
      const slideIdx = slideIndexRef.current;
      const isCurrentSlideFormat = format.group === "current";
      const slideSuffix = isCurrentSlideFormat ? `-slide${slideIdx + 1}` : "";
      const downloadName = `${baseName}${slideSuffix}.${format.extension}`;
      try {
        switch (format.id) {
          case "pptx": {
            const buf = await a.exportFile();
            await saveFile(
              new Uint8Array(buf),
              downloadName,
              format.mime,
              undefined
            );
            break;
          }
          case "pdf":
          case "html": {
            const buf = await a.exportFile();
            const out = await convertViaServer({
              bytes: new Uint8Array(buf),
              sourceExt: "pptx",
              targetExt: format.id,
              filename: baseName,
            });
            downloadBlob(out, downloadName);
            break;
          }
          case "slide-pdf": {
            // Server-converted single-slide PDF. We send the whole
            // pptx and let LibreOffice's `PageRange` filter cut it
            // down — that keeps the renderer a single code path
            // (vs. building a one-slide pptx in the browser) and
            // matches what PowerPoint's "Print → Selection" does
            // under the hood.
            const buf = await a.exportFile();
            const out = await convertViaServer({
              bytes: new Uint8Array(buf),
              sourceExt: "pptx",
              targetExt: "pdf",
              filename: `${baseName}-slide${slideIdx + 1}`,
              pageRange: String(slideIdx + 1),
            });
            downloadBlob(out, downloadName);
            break;
          }
          case "slide-png": {
            const snap = a.getSnapshot();
            const scale = parseScale(options?.scale);
            const blob = await snapshotToSlidePng(snap, slideIdx, { scale });
            downloadBlob(blob, downloadName);
            break;
          }
          case "slide-jpeg": {
            const snap = a.getSnapshot();
            const scale = parseScale(options?.scale);
            const quality = parseJpegQuality(options?.quality);
            const blob = await snapshotToSlideJpeg(snap, slideIdx, { scale, quality });
            downloadBlob(blob, downloadName);
            break;
          }
          case "slide-svg": {
            const snap = a.getSnapshot();
            const svg = snapshotToSlideSvg(snap, slideIdx);
            downloadBlob(new Blob([svg], { type: format.mime }), downloadName);
            break;
          }
          case "png-zip": {
            const snap = a.getSnapshot();
            const total = snap.root.slides.length;
            const rangeRaw = typeof options?.slideRange === "string" ? options.slideRange : "";
            const indices = parseSlideRange(rangeRaw, total);
            if (indices.length === 0) {
              throw new Error("Slide range matched no slides.");
            }
            const scale = parseScale(options?.scale);
            const blob = await snapshotToPngZip(snap, { scale, indices });
            downloadBlob(blob, downloadName);
            break;
          }
          case "svg-zip": {
            const snap = a.getSnapshot();
            const blob = await snapshotToSvgZip(snap);
            downloadBlob(blob, downloadName);
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

  const snap = agent?.getSnapshot() ?? null;
  void tick;
  const slides = snap?.root.slides ?? [];
  const slideSize = snap?.root.slideSize ?? { cxEmu: 9144000, cyEmu: 6858000 };
  const themeDefault = snap?.root.themeDefault;
  const slide = slides[activeIndex];

  const selectedShape = useMemo<Shape | null>(() => {
    if (!slide || !selectedShapeId) return null;
    return findShape(slide.shapes, selectedShapeId);
  }, [slide, selectedShapeId]);

  // Shape ids that have an unresolved comment thread on the active
  // slide. Reply comments and resolved threads are filtered out so the
  // yellow indicator reads "open conversation".
  const commentedShapeIds = useMemo<ReadonlyArray<string>>(() => {
    if (!snap || !slide || !slide.commentsPartPath) return [];
    const part = snap.root.commentsByPart.get(slide.commentsPartPath);
    if (!part) return [];
    const ids = new Set<string>();
    for (const c of part.comments) {
      if (c.parentId) continue;
      if (c.resolved) continue;
      if (!c.shapeId) continue;
      ids.add(c.shapeId);
    }
    return [...ids];
  }, [snap, slide]);

  const currentFill = useMemo(() => {
    if (!selectedShape) return null;
    if (selectedShape.kind !== "text" && selectedShape.kind !== "pic") return null;
    return readSolidFill(selectedShape) ?? null;
  }, [selectedShape]);

  // ActiveTextFormat for the shared TextFormatBar. Recomputed on every
  // render that touches the snapshot tick or text selection so the
  // toolbar reflects the current run-level formatting at the caret.
  const textFormatActive = useMemo(() => {
    void tick;
    return computePptxActive(agent, activeIndex, textSelection, selectedShapeId);
  }, [agent, activeIndex, textSelection, selectedShapeId, tick]);

  const addSlide = useCallback(async () => {
    const a = agentRef.current;
    if (!a) return;
    try {
      await a.applyCommand({ type: "pptx:add-slide", payload: {}, source: "human" });
      setActiveIndex(a.getSnapshot().root.slides.length - 1);
    } catch (err) {
      onError(err);
    }
  }, [onError]);

  const addSlideWithLayout = useCallback(
    async (kind: LayoutKindPayload) => {
      const a = agentRef.current;
      if (!a) return;
      try {
        await a.applyCommand({
          type: "pptx:add-slide",
          payload: { layoutKind: kind },
          source: "human",
        });
        setActiveIndex(a.getSnapshot().root.slides.length - 1);
      } catch (err) {
        onError(err);
      }
    },
    [onError]
  );

  const setSlideLayout = useCallback(
    async (kind: LayoutKindPayload) => {
      const a = agentRef.current;
      if (!a) return;
      try {
        await a.applyCommand({
          type: "pptx:set-slide-layout",
          payload: { slideIndex: activeIndex, layoutKind: kind },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError]
  );

  const deleteSlide = useCallback(async () => {
    const a = agentRef.current;
    if (!a) return;
    if (a.getSnapshot().root.slides.length <= 1) {
      pushToast("warn", "Cannot delete the last slide.");
      return;
    }
    try {
      await a.applyCommand({
        type: "pptx:delete-slide",
        payload: { slideIndex: activeIndex },
        source: "human",
      });
      setActiveIndex((i) => Math.max(0, i - 1));
    } catch (err) {
      onError(err);
    }
  }, [activeIndex, onError, pushToast]);

  const duplicateSlide = useCallback(async () => {
    const a = agentRef.current;
    if (!a) return;
    try {
      await a.applyCommand({
        type: "pptx:duplicate-slide",
        payload: { slideIndex: activeIndex },
        source: "human",
      });
      setActiveIndex(activeIndex + 1);
    } catch (err) {
      onError(err);
    }
  }, [activeIndex, onError]);

  // D7 — slide thumbnail drag-reorder. The sidebar reports
  // `(from, to)` indices already adjusted for splice semantics so we
  // can dispatch `pptx:move-slide` directly.
  const moveSlide = useCallback(
    async (from: number, to: number) => {
      const a = agentRef.current;
      if (!a) return;
      if (from === to) return;
      try {
        await a.applyCommand({
          type: "pptx:move-slide",
          payload: { from, to },
          source: "human",
        });
        // Track the moved slide so it stays selected after the
        // reorder. `setActiveIndex` is otherwise sticky to the old
        // index, which would visually jump to the wrong slide.
        setActiveIndex(to);
      } catch (err) {
        onError(err);
      }
    },
    [onError]
  );

  // D7 — slide thumbnail context menu. Handles every entry the sidebar
  // can surface (`SlideContextAction`), reusing the existing
  // add/duplicate/delete/move commands. Insert-before / insert-after
  // dispatch `pptx:add-slide` with the explicit `at` slot rather than
  // the toolbar's append-to-end semantics.
  const handleSlideContextAction = useCallback(
    async (slideIndex: number, action: SlideContextAction) => {
      const a = agentRef.current;
      if (!a) return;
      const total = a.getSnapshot().root.slides.length;
      try {
        switch (action) {
          case "insert-before": {
            await a.applyCommand({
              type: "pptx:add-slide",
              payload: { at: slideIndex },
              source: "human",
            });
            setActiveIndex(slideIndex);
            break;
          }
          case "insert-after": {
            await a.applyCommand({
              type: "pptx:add-slide",
              payload: { at: slideIndex + 1 },
              source: "human",
            });
            setActiveIndex(slideIndex + 1);
            break;
          }
          case "duplicate": {
            await a.applyCommand({
              type: "pptx:duplicate-slide",
              payload: { slideIndex },
              source: "human",
            });
            setActiveIndex(slideIndex + 1);
            break;
          }
          case "delete": {
            if (total <= 1) {
              pushToast("warn", "Cannot delete the last slide.");
              return;
            }
            await a.applyCommand({
              type: "pptx:delete-slide",
              payload: { slideIndex },
              source: "human",
            });
            setActiveIndex(Math.max(0, Math.min(slideIndex, total - 2)));
            break;
          }
          case "move-up": {
            if (slideIndex === 0) return;
            await a.applyCommand({
              type: "pptx:move-slide",
              payload: { from: slideIndex, to: slideIndex - 1 },
              source: "human",
            });
            setActiveIndex(slideIndex - 1);
            break;
          }
          case "move-down": {
            if (slideIndex >= total - 1) return;
            await a.applyCommand({
              type: "pptx:move-slide",
              payload: { from: slideIndex, to: slideIndex + 1 },
              source: "human",
            });
            setActiveIndex(slideIndex + 1);
            break;
          }
          case "move-to-start": {
            if (slideIndex === 0) return;
            await a.applyCommand({
              type: "pptx:move-slide",
              payload: { from: slideIndex, to: 0 },
              source: "human",
            });
            setActiveIndex(0);
            break;
          }
          case "move-to-end": {
            if (slideIndex >= total - 1) return;
            await a.applyCommand({
              type: "pptx:move-slide",
              payload: { from: slideIndex, to: total - 1 },
              source: "human",
            });
            setActiveIndex(total - 1);
            break;
          }
          default: {
            const exhaust: never = action;
            throw new Error(`unhandled slide action: ${String(exhaust)}`);
          }
        }
      } catch (err) {
        onError(err);
      }
    },
    [onError, pushToast]
  );

  // Each "Insert" command drops the new shape near the slide centre at a
  // sensible default size. We avoid stacking duplicates by nudging by an
  // extra (count * step) so successive inserts don't pile up at the exact
  // same coordinates.
  const insertOffset = useCallback((): { x: number; y: number } => {
    const a = agentRef.current;
    const s = a?.getSnapshot().root.slides[activeIndex];
    const count = s?.shapes.length ?? 0;
    const step = 200_000;
    return {
      x: 1_000_000 + (count % 8) * step,
      y: 1_000_000 + (count % 8) * step,
    };
  }, [activeIndex]);

  const addTextBox = useCallback(async () => {
    const a = agentRef.current;
    if (!a) return;
    try {
      const off = insertOffset();
      await a.applyCommand({
        type: "pptx:add-text-box",
        payload: {
          slideIndex: activeIndex,
          text: "New text box",
          x: off.x,
          y: off.y,
          width: 4_000_000,
          height: 800_000,
        },
        source: "human",
      });
      const s = a.getSnapshot().root.slides[activeIndex];
      setSelectedShapeIds([s.shapes[s.shapes.length - 1]!.id]);
    } catch (err) {
      onError(err);
    }
  }, [activeIndex, insertOffset, onError]);

  const addShape = useCallback(
    async (preset: ShapePreset) => {
      const a = agentRef.current;
      if (!a) return;
      try {
        const off = insertOffset();
        const isLine = preset === "line";
        await a.applyCommand({
          type: "pptx:add-shape",
          payload: {
            slideIndex: activeIndex,
            preset,
            x: off.x,
            y: off.y,
            width: isLine ? 3_000_000 : 2_500_000,
            height: isLine ? 0 : preset === "ellipse" ? 1_500_000 : 1_500_000,
          },
          source: "human",
        });
        const s = a.getSnapshot().root.slides[activeIndex];
        setSelectedShapeIds([s.shapes[s.shapes.length - 1]!.id]);
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, insertOffset, onError]
  );

  // Toggle connector tool mode for the requested type. Re-clicking
  // the same type while it's already armed exits the mode (matches
  // Figma / Slides). The actual `pptx:add-connector` command runs
  // from inside the canvas once the user finishes their drag — that's
  // what makes the experience feel like a real drawing tool instead
  // of a "click button → guess where the line landed" form.
  const startConnectorTool = useCallback(
    (connectorType: "straight" | "elbow" | "curved") => {
      setConnectorTool((prev) => (prev?.type === connectorType ? null : { type: connectorType }));
    },
    []
  );
  const exitConnectorTool = useCallback(() => setConnectorTool(null), []);

  // Apply a partial style patch from the floating connector mini-bar.
  // The bar passes only the fields the user actually touched; we relay
  // them straight through to `pptx:set-connector-style` so undo/redo
  // sees a single command per click instead of separate commands per
  // field. The empty-patch guard mirrors the command's own validation
  // so a no-op interaction (e.g. picker closed without a change) costs
  // zero round-trips.
  const applyConnectorStylePatch = useCallback(
    async (shapeId: string, patch: ConnectorStylePatch) => {
      const a = agentRef.current;
      if (!a) return;
      const hasField =
        patch.connectorType !== undefined ||
        patch.strokeColor !== undefined ||
        patch.strokeWidthEmu !== undefined ||
        patch.strokeDash !== undefined ||
        patch.headEnd !== undefined ||
        patch.tailEnd !== undefined;
      if (!hasField) return;
      try {
        await a.applyCommand({
          type: "pptx:set-connector-style",
          source: "human",
          payload: {
            slideIndex: activeIndex,
            shapeId,
            ...patch,
          },
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError]
  );

  const insertImage = useCallback(
    async (file: File) => {
      const a = agentRef.current;
      if (!a) return;
      const mime = (file.type || "").toLowerCase();
      if (!SUPPORTED_IMAGE_MIME.has(mime)) {
        pushToast(
          "error",
          `Unsupported image type "${mime || "unknown"}". Use PNG, JPEG, GIF, BMP, or WEBP.`
        );
        return;
      }
      try {
        const buf = await file.arrayBuffer();
        const bytes = new Uint8Array(buf);
        const intrinsic = await readIntrinsicSize(bytes, mime);
        // Cap the displayed width at ~half a 16:9 slide; preserves aspect.
        const maxWidthEmu = Math.floor(slideSize.cxEmu / 2);
        let widthEmu = intrinsic.width * 9525; // 1 px ≈ 9525 EMU @ 96 DPI
        let heightEmu = intrinsic.height * 9525;
        if (widthEmu > maxWidthEmu) {
          const r = maxWidthEmu / widthEmu;
          widthEmu = Math.round(widthEmu * r);
          heightEmu = Math.round(heightEmu * r);
        }
        if (widthEmu <= 0 || heightEmu <= 0) {
          widthEmu = 2_000_000;
          heightEmu = 2_000_000;
        }
        const off = insertOffset();
        await a.applyCommand({
          type: "pptx:insert-image",
          payload: {
            slideIndex: activeIndex,
            data: bytes,
            mimeType: mime === "image/jpg" ? "image/jpeg" : mime,
            x: off.x,
            y: off.y,
            width: widthEmu,
            height: heightEmu,
            altText: file.name,
            name: file.name,
          },
          source: "human",
        });
        const s = a.getSnapshot().root.slides[activeIndex];
        setSelectedShapeIds([s.shapes[s.shapes.length - 1]!.id]);
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, insertOffset, onError, pushToast, slideSize.cxEmu]
  );

  /**
   * D9 — replace the bitmap behind the currently selected `Picture`
   * with a user-picked file. Position, size, alt-text, and any
   * `spPrTail` styling are preserved by `pptx:replace-picture-media`.
   * No-ops when the selection isn't a picture.
   */
  const replaceSelectedPicture = useCallback(
    async (file: File) => {
      const a = agentRef.current;
      if (!a || !selectedShapeId) return;
      const sh = slide ? findShape(slide.shapes, selectedShapeId) : null;
      if (!sh || sh.kind !== "pic") {
        pushToast("error", "Replace image only works on the selected picture.");
        return;
      }
      const mime = (file.type || "").toLowerCase();
      if (!SUPPORTED_IMAGE_MIME.has(mime)) {
        pushToast(
          "error",
          `Unsupported image type "${mime || "unknown"}". Use PNG, JPEG, GIF, BMP, or WEBP.`
        );
        return;
      }
      try {
        const buf = await file.arrayBuffer();
        const bytes = new Uint8Array(buf);
        await a.applyCommand({
          type: "pptx:replace-picture-media",
          payload: {
            slideIndex: activeIndex,
            shapeId: sh.id,
            data: bytes,
            mimeType: mime === "image/jpg" ? "image/jpeg" : mime,
            altText: file.name,
          },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError, pushToast, selectedShapeId, slide]
  );

  /**
   * Fired by the canvas when the user activates an empty placeholder
   * (double-click, or PowerPoint-style "click again to enter") whose
   * type isn't text. For `pic` we open a file picker, drop the chosen
   * image into the slot the placeholder was occupying — same x / y /
   * cx / cy — and then remove the placeholder. The two commands are
   * issued sequentially so each sees the previous mutation; if the
   * insert fails we leave the placeholder untouched, which preserves
   * the slot for a retry.
   *
   * Other placeholder types (chart / tbl / dgm / media) trigger a
   * "not yet implemented" toast for now — the wiring scaffolding is
   * here for when those wizards exist.
   */
  const handlePlaceholderActivate = useCallback(
    async (info: { shapeId: string; placeholder: { type?: string; idx?: number } }) => {
      const a = agentRef.current;
      if (!a) return;
      const type = info.placeholder.type ?? "";
      if (type !== "pic") {
        pushToast(
          "info",
          `Filling "${type}" placeholders from the canvas isn't supported yet — use the toolbar.`
        );
        return;
      }
      const currentSlide = a.getSnapshot().root.slides[activeIndex];
      const placeholderShape = currentSlide ? findShape(currentSlide.shapes, info.shapeId) : null;
      if (!placeholderShape || !placeholderShape.position || !placeholderShape.size) {
        return;
      }
      const slot = {
        x: placeholderShape.position.xEmu,
        y: placeholderShape.position.yEmu,
        cx: placeholderShape.size.cxEmu,
        cy: placeholderShape.size.cyEmu,
      };
      const file = await pickImageFile();
      if (!file) return;
      const mime = (file.type || "").toLowerCase();
      if (!SUPPORTED_IMAGE_MIME.has(mime)) {
        pushToast(
          "error",
          `Unsupported image type "${mime || "unknown"}". Use PNG, JPEG, GIF, BMP, or WEBP.`
        );
        return;
      }
      try {
        const buf = await file.arrayBuffer();
        const bytes = new Uint8Array(buf);
        await a.applyCommand({
          type: "pptx:insert-image",
          payload: {
            slideIndex: activeIndex,
            data: bytes,
            mimeType: mime === "image/jpg" ? "image/jpeg" : mime,
            x: slot.x,
            y: slot.y,
            width: slot.cx,
            height: slot.cy,
            altText: file.name,
            name: file.name,
          },
          source: "human",
        });
        await a.applyCommand({
          type: "pptx:delete-shape",
          payload: { slideIndex: activeIndex, shapeId: info.shapeId },
          source: "human",
        });
        const updated = a.getSnapshot().root.slides[activeIndex];
        const last = updated?.shapes[updated.shapes.length - 1];
        if (last) setSelectedShapeIds([last.id]);
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError, pushToast]
  );

  const deleteSelectedShape = useCallback(async () => {
    const a = agentRef.current;
    if (!a || selectedShapeIds.length === 0) return;
    try {
      // Delete each selected shape independently. We dispatch sequentially
      // so each command sees the previous mutation's snapshot — the
      // command bus rejects "unknown shape" otherwise when later ids
      // happen to share group ancestry.
      for (const id of selectedShapeIds) {
        await a.applyCommand({
          type: "pptx:delete-shape",
          payload: { slideIndex: activeIndex, shapeId: id },
          source: "human",
        });
      }
      setSelectedShapeIds([]);
    } catch (err) {
      onError(err);
    }
  }, [activeIndex, onError, selectedShapeIds]);

  // D3 — Cmd+D duplicate. Clones every selected shape on the active
  // slide using the duplicate-shape command (¼" diagonal nudge), then
  // selects the freshly-minted clones so the user can immediately
  // drag/move them. Skipped when nothing is selected.
  // D3 — Cmd+Shift+G group / Cmd+Shift+Alt+G ungroup. Uses the
  // group-shapes / ungroup-shape commands. We refuse silently if the
  // selection doesn't satisfy the command's preconditions (the command
  // bus would reject anyway, but a toast on every failed Tab cycle is
  // noisy).
  const groupSelectedShapes = useCallback(async () => {
    const a = agentRef.current;
    if (!a || selectedShapeIds.length < 2) return;
    try {
      await a.applyCommand({
        type: "pptx:group-shapes",
        payload: { slideIndex: activeIndex, shapeIds: selectedShapeIds },
        source: "human",
      });
      const slide = a.getSnapshot().root.slides[activeIndex];
      const last = slide.shapes[slide.shapes.length - 1];
      if (last && last.kind === "group") setSelectedShapeIds([last.id]);
    } catch (err) {
      onError(err);
    }
  }, [activeIndex, onError, selectedShapeIds]);

  const ungroupSelectedShape = useCallback(async () => {
    const a = agentRef.current;
    if (!a || selectedShapeIds.length !== 1) return;
    const slide = a.getSnapshot().root.slides[activeIndex];
    const target = slide.shapes.find((s) => s.id === selectedShapeIds[0]);
    if (!target || target.kind !== "group") return;
    const childIds = target.children.map((c) => c.id);
    try {
      await a.applyCommand({
        type: "pptx:ungroup-shape",
        payload: { slideIndex: activeIndex, shapeId: target.id },
        source: "human",
      });
      setSelectedShapeIds(childIds);
    } catch (err) {
      onError(err);
    }
  }, [activeIndex, onError, selectedShapeIds]);

  const duplicateSelectedShapes: () => Promise<void> = useCallback(async () => {
    const a = agentRef.current;
    if (!a || selectedShapeIds.length === 0) return;
    try {
      const newIds: string[] = [];
      for (const id of selectedShapeIds) {
        await a.applyCommand({
          type: "pptx:duplicate-shape",
          payload: { slideIndex: activeIndex, shapeId: id },
          source: "human",
        });
        const slide = a.getSnapshot().root.slides[activeIndex];
        const last = slide.shapes[slide.shapes.length - 1];
        if (last) newIds.push(last.id);
      }
      if (newIds.length > 0) setSelectedShapeIds(newIds);
    } catch (err) {
      onError(err);
    }
  }, [activeIndex, onError, selectedShapeIds]);

  // Keep the refs in sync so the keyboard listener (bound once) always
  // calls the latest versions of the closures.
  useEffect(() => {
    duplicateRef.current = () => void duplicateSelectedShapes();
  }, [duplicateSelectedShapes]);
  useEffect(() => {
    groupRef.current = () => void groupSelectedShapes();
  }, [groupSelectedShapes]);
  useEffect(() => {
    ungroupRef.current = () => void ungroupSelectedShape();
  }, [ungroupSelectedShape]);

  const alignSelected = useCallback(
    async (
      mode: "left" | "center-h" | "right" | "top" | "middle-v" | "bottom",
      relativeTo: "selection" | "slide"
    ) => {
      const a = agentRef.current;
      if (!a) return;
      const minCount = relativeTo === "slide" ? 1 : 2;
      if (selectedShapeIds.length < minCount) return;
      try {
        await a.applyCommand({
          type: "pptx:align-shapes",
          payload: {
            slideIndex: activeIndex,
            shapeIds: selectedShapeIds,
            mode,
            relativeTo,
          },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError, selectedShapeIds]
  );

  const distributeSelected = useCallback(
    async (axis: "horizontal" | "vertical") => {
      const a = agentRef.current;
      if (!a || selectedShapeIds.length < 3) return;
      try {
        await a.applyCommand({
          type: "pptx:distribute-shapes",
          payload: { slideIndex: activeIndex, shapeIds: selectedShapeIds, axis },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError, selectedShapeIds]
  );

  // D5 — z-order. Operates on every selected shape sequentially so
  // multi-select "send to back" / "bring to front" produce the right
  // *relative* stacking even when shapes were originally at different
  // depths.
  const reorderSelected = useCallback(
    async (mode: import("@officeai/pptx").ReorderShapeMode) => {
      const a = agentRef.current;
      if (!a || selectedShapeIds.length === 0) return;
      const ids =
        mode === "to-front" || mode === "forward" ? selectedShapeIds : [...selectedShapeIds].reverse();
      try {
        for (const shapeId of ids) {
          await a.applyCommand({
            type: "pptx:reorder-shape",
            payload: { slideIndex: activeIndex, shapeId, mode },
            source: "human",
          });
        }
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError, selectedShapeIds]
  );

  /**
   * D11 — Animations panel command bridges. Each one is a thin
   * forwarder onto the existing `pptx:set-slide-transition`,
   * `pptx:add-shape-animation`, `pptx:remove-shape-animation`, and
   * `pptx:reorder-shape-animations` handlers. Add-animation requires
   * a selected shape because animations are anchored to a `cNvPrId`.
   */
  const setSlideTransition = useCallback(
    async (
      kind: import("@officeai/pptx").TransitionKind,
      speed: import("@officeai/pptx").TransitionSpeed | null
    ) => {
      const a = agentRef.current;
      if (!a) return;
      try {
        // The "unsupported" landing-pad kind is only ever emitted by
        // the parser; the panel never offers it.
        if (kind === "unsupported") return;
        await a.applyCommand({
          type: "pptx:set-slide-transition",
          payload: {
            slideIndex: activeIndex,
            kind,
            ...(speed ? { speed } : {}),
          },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError]
  );

  const addShapeAnimation = useCallback(
    async (effect: import("@officeai/pptx").EntranceEffect) => {
      const a = agentRef.current;
      if (!a || !selectedShapeId) return;
      try {
        await a.applyCommand({
          type: "pptx:add-shape-animation",
          payload: { slideIndex: activeIndex, shapeId: selectedShapeId, effect },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError, selectedShapeId]
  );

  const removeShapeAnimation = useCallback(
    async (animationId: string) => {
      const a = agentRef.current;
      if (!a) return;
      try {
        await a.applyCommand({
          type: "pptx:remove-shape-animation",
          payload: { slideIndex: activeIndex, animationId },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError]
  );

  const reorderShapeAnimations = useCallback(
    async (orderIds: ReadonlyArray<string>) => {
      const a = agentRef.current;
      if (!a) return;
      try {
        await a.applyCommand({
          type: "pptx:reorder-shape-animations",
          payload: { slideIndex: activeIndex, order: orderIds },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError]
  );

  const setSlideNotes = useCallback(
    async (text: string) => {
      const a = agentRef.current;
      if (!a) return;
      try {
        await a.applyCommand({
          type: "pptx:set-slide-notes",
          payload: { slideIndex: activeIndex, text },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError]
  );

  const changeFill = useCallback(
    async (hex: string | null) => {
      const a = agentRef.current;
      if (!a || !selectedShapeId) return;
      try {
        await a.applyCommand({
          type: "pptx:set-shape-fill",
          payload: { slideIndex: activeIndex, shapeId: selectedShapeId, fill: hex },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError, selectedShapeId]
  );

  // Toolbar "Add comment" entry point. The shared right-rail owns the
  // comments panel + composer; this just opens it and focuses the
  // textarea via a custom event the rail listens for.
  const focusCommentComposer = useCallback(() => {
    requestAnimationFrame(() => {
      const root = document.querySelector<HTMLElement>('[data-testid="rail-tab-comments"]');
      if (root) (root as HTMLButtonElement).click();
      requestAnimationFrame(() => {
        const textarea = document.querySelector<HTMLTextAreaElement>(
          'aside [data-testid="comment-composer"] textarea'
        );
        textarea?.focus();
      });
    });
  }, []);

  // "Click to locate" handler wired into the comments sidebar. Looks
  // up the comment in the active slide's comments part, switches the
  // canvas selection to the anchored shape (if any) and triggers a
  // yellow flash overlay so the user can find it. Falls back to a
  // pin-position flash for free-floating comments.
  const scrollToComment = useCallback(
    (commentId: string) => {
      const a = agentRef.current;
      if (!a) return;
      const snap = a.getSnapshot();
      const target = snap.root.slides[activeIndex];
      if (!target || !target.commentsPartPath) return;
      const part = snap.root.commentsByPart.get(target.commentsPartPath);
      if (!part) return;
      const c = part.comments.find((cm) => cm.id === commentId);
      if (!c) return;
      const surface = slideSurfaceRef.current;
      if (surface) surface.scrollIntoView({ block: "nearest", behavior: "smooth" });
      if (c.shapeId) {
        setSelectedShapeIds([c.shapeId]);
        setCommentFlashTarget((prev) => ({
          kind: "shape",
          shapeId: c.shapeId!,
          nonce: (prev?.nonce ?? 0) + 1,
        }));
      } else {
        setCommentFlashTarget((prev) => ({
          kind: "pin",
          xEmu: c.xEmu,
          yEmu: c.yEmu,
          nonce: (prev?.nonce ?? 0) + 1,
        }));
      }
    },
    [activeIndex]
  );

  // Auto-clear the flash overlay after the animation completes so we
  // don't keep stale SVG nodes in the tree.
  useEffect(() => {
    if (!commentFlashTarget) return;
    const handle = window.setTimeout(() => setCommentFlashTarget(null), 1700);
    return () => window.clearTimeout(handle);
  }, [commentFlashTarget]);

  usePptxShortcuts({
    surfaceRef: slideSurfaceRef,
    agentRef,
    activeIndex,
    slideCount: slides.length,
    selectedShape,
    selectedShapeIds,
    textFormatProvider,
    onAddSlide: () => void addSlide(),
    onDuplicateSlide: () => void duplicateSlide(),
    onDeleteShape: () => void deleteSelectedShape(),
    onChangeSlide: setActiveIndex,
    onError,
  });

  /**
   * D10 — global F5 / Shift+F5 to enter Present mode. F5 starts from
   * slide 1, Shift+F5 starts from the currently active slide. We
   * consume the event so the browser's native F5 (page reload) is
   * suppressed while a presentation is loaded — matching PowerPoint.
   * The listener is gated on `ready` so it never fires before a
   * snapshot is available.
   */
  const startPresenting = useCallback(
    (fromCurrent: boolean) => {
      if (!ready) return;
      if (!fromCurrent) setActiveIndex(0);
      setPresenting(true);
    },
    [ready]
  );
  useEffect(() => {
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === "F5") {
        const target = e.target as HTMLElement | null;
        if (
          target &&
          (target.tagName === "INPUT" || target.tagName === "TEXTAREA" || target.isContentEditable)
        ) {
          return;
        }
        e.preventDefault();
        startPresenting(e.shiftKey);
      }
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [startPresenting]);

  // Build object-URL map for every embedded media part so the renderer
  // can paint <Picture> shapes inserted via the toolbar. We rebuild on
  // every snapshot tick — cheap because URLs are deduplicated by part path.
  const mediaUrls = useMemo(() => {
    const map = new Map<string, string>();
    if (!snap) return map;
    for (const [path, part] of snap.root.media) {
      const blob = new Blob([part.bytes as BlobPart], { type: part.contentType });
      map.set(path, URL.createObjectURL(blob));
    }
    return map;
  }, [snap]);
  useEffect(() => {
    return () => {
      for (const url of mediaUrls.values()) URL.revokeObjectURL(url);
    };
  }, [mediaUrls]);

  // Comment-thread count for the right rail badge.
  const openCommentCount = useMemo(() => {
    if (!snap || !slide || !slide.commentsPartPath) return 0;
    const part = snap.root.commentsByPart.get(slide.commentsPartPath);
    if (!part) return 0;
    return part.comments.filter((c) => !c.parentId && !c.resolved).length;
  }, [snap, slide]);

  const renderCommentsPanel = useCallback((): React.ReactNode => {
    if (!agent || !snap || !slide) {
      return <div className="p-4 text-sm text-secondary">No slide selected.</div>;
    }
    return (
      <div className="flex min-h-0 flex-1 flex-col">
        <div className="min-h-0 flex-1 overflow-y-auto">
          <CommentsSidebar
            key={`comments-${activeIndex}-${tick}`}
            provider={createPptxCommentsProvider({
              agent,
              slideIndex: activeIndex,
              onScrollTo: scrollToComment,
            })}
            author="You"
            emptyHint="No comments on this slide yet. Select a shape and press Add comment, or comment on the slide as a whole."
            onScrollTo={scrollToComment}
          />
        </div>
        <div className="border-t border-divider p-2">
          <CommentComposer
            provider={createPptxCommentsProvider({ agent, slideIndex: activeIndex })}
            anchor={
              selectedShape && selectedShape.position
                ? {
                    kind: "pptx-pin",
                    slideIndex: activeIndex,
                    xEmu: selectedShape.position.xEmu + Math.round((selectedShape.size?.cxEmu ?? 0) / 2),
                    yEmu: selectedShape.position.yEmu + Math.round((selectedShape.size?.cyEmu ?? 0) / 2),
                    shapeId: selectedShape.id,
                  }
                : {
                    kind: "pptx-pin",
                    slideIndex: activeIndex,
                    xEmu: Math.round(slideSize.cxEmu / 2),
                    yEmu: Math.round(slideSize.cyEmu / 2),
                  }
            }
            placeholder={
              selectedShape
                ? `Comment on ${selectedShape.name || "selected shape"}…`
                : "Add a comment to this slide…"
            }
          />
        </div>
      </div>
    );
  }, [
    activeIndex,
    agent,
    scrollToComment,
    selectedShape,
    slide,
    slideSize.cxEmu,
    slideSize.cyEmu,
    snap,
    tick,
  ]);

  // Selection summary surfaced in the status bar.
  const selectionText = useMemo(() => {
    if (!ready) return "";
    if (selectedShapeIds.length === 0) {
      return `Slide ${activeIndex + 1} of ${slides.length}`;
    }
    if (selectedShapeIds.length === 1) {
      return `${selectedShape?.name ?? "Shape"} selected · Slide ${activeIndex + 1} of ${slides.length}`;
    }
    return `${selectedShapeIds.length} shapes selected · Slide ${activeIndex + 1} of ${slides.length}`;
  }, [activeIndex, ready, selectedShape?.name, selectedShapeIds.length, slides.length]);

  // Curated palette commands. The PPTX agent's typed command bus is
  // also indexable here; we surface only the user-facing entries.
  const paletteCommands = useMemo<ReadonlyArray<PaletteCommand>>(() => {
    return [
      { id: "pptx.add-slide", label: "Insert new slide", section: "Slide", run: () => void addSlide() },
      {
        id: "pptx.duplicate-slide",
        label: "Duplicate slide",
        section: "Slide",
        run: () => void duplicateSlide(),
      },
      {
        id: "pptx.delete-slide",
        label: "Delete slide",
        section: "Slide",
        run: () => void deleteSlide(),
        enabled: slides.length > 1,
      },
      { id: "pptx.add-text", label: "Insert text box", section: "Insert", run: () => void addTextBox() },
      { id: "pptx.add-rect", label: "Insert rectangle", section: "Insert", run: () => void addShape("rect") },
      {
        id: "pptx.add-ellipse",
        label: "Insert ellipse",
        section: "Insert",
        run: () => void addShape("ellipse"),
      },
      {
        id: "pptx.add-arrow",
        label: "Insert arrow",
        section: "Insert",
        run: () => void addShape("rightArrow"),
      },
      {
        id: "pptx.add-connector-elbow",
        label: "Draw elbow connector",
        section: "Insert",
        run: () => startConnectorTool("elbow"),
      },
      {
        id: "pptx.add-connector-straight",
        label: "Draw straight connector",
        section: "Insert",
        run: () => startConnectorTool("straight"),
      },
      {
        id: "pptx.add-connector-curved",
        label: "Draw curved connector",
        section: "Insert",
        run: () => startConnectorTool("curved"),
      },
      {
        id: "pptx.add-comment",
        label: "Add comment",
        section: "Collaboration",
        run: () => focusCommentComposer(),
      },
      {
        id: "pptx.delete-shape",
        label: "Delete selected shape(s)",
        section: "Edit",
        run: () => void deleteSelectedShape(),
        enabled: selectedShapeIds.length > 0,
      },
      {
        id: "pptx.duplicate-shape",
        label: "Duplicate selected shape(s)",
        section: "Edit",
        run: () => void duplicateSelectedShapes(),
        enabled: selectedShapeIds.length > 0,
      },
      {
        id: "pptx.group-shapes",
        label: "Group selected shapes",
        section: "Arrange",
        run: () => void groupSelectedShapes(),
        enabled: selectedShapeIds.length >= 2,
      },
      {
        id: "pptx.ungroup-shape",
        label: "Ungroup",
        section: "Arrange",
        run: () => void ungroupSelectedShape(),
        enabled: selectedShapeIds.length === 1,
      },
      { id: "pptx.zoom-reset", label: "Reset zoom to 100%", section: "View", run: () => setZoom(1) },
      {
        id: "pptx.present-from-start",
        label: "Start presentation from beginning (F5)",
        section: "View",
        run: () => startPresenting(false),
        enabled: ready && slides.length > 0,
      },
      {
        id: "pptx.present-from-current",
        label: "Start presentation from current slide (Shift+F5)",
        section: "View",
        run: () => startPresenting(true),
        enabled: ready && slides.length > 0,
      },
    ];
  }, [
    startConnectorTool,
    addShape,
    addSlide,
    addTextBox,
    deleteSelectedShape,
    deleteSlide,
    duplicateSelectedShapes,
    duplicateSlide,
    focusCommentComposer,
    groupSelectedShapes,
    ready,
    selectedShapeIds.length,
    slides.length,
    startPresenting,
    ungroupSelectedShape,
  ]);

  const adapter = useMemo<ProductAdapter>(
    () => ({
      product: "pptx",
      filename: docName,
      saveState,
      comments: { openCount: openCommentCount, resolvedCount: 0 },
      selectionSummary: { text: selectionText },
      canOpen: true,
      canSave: ready,
      canExport: ready,
      exportFormats: PPTX_EXPORT_FORMATS,
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
      renderCommentsPanel,
      renderAnimationsPanel: snap
        ? () => (
            <AnimationsPanel
              snapshot={snap}
              activeIndex={activeIndex}
              selectedShape={selectedShape}
              disabled={!ready}
              onSetTransition={(kind, speed) => void setSlideTransition(kind, speed)}
              onAddAnimation={(effect) => void addShapeAnimation(effect)}
              onRemoveAnimation={(id) => void removeShapeAnimation(id)}
              onReorderAnimations={(orderIds) => void reorderShapeAnimations(orderIds)}
            />
          )
        : undefined,
      onAddComment: focusCommentComposer,
    }),
    [
      activeIndex,
      addShapeAnimation,
      docName,
      focusCommentComposer,
      handleExport,
      handleOpenFile,
      handleSave,
      openCommentCount,
      paletteCommands,
      ready,
      removeShapeAnimation,
      renderCommentsPanel,
      reorderShapeAnimations,
      saveState,
      selectedShape,
      selectionText,
      setSlideTransition,
      shortcutsDialog,
      snap,
      tick,
    ]
  );

  return (
    <>
      <EditorShell
        adapter={adapter}
        toolbar={
          <PptxToolbar
            disabled={!ready}
            slideCount={slides.length}
            hasSelection={selectedShapeId != null}
            selectionCount={selectedShapeIds.length}
            currentFill={currentFill ? `#${currentFill}` : null}
            textFormatProvider={textFormatProvider}
            textFormatActive={textFormatActive}
            onAddSlide={() => void addSlide()}
            onAddSlideWithLayout={(k) => void addSlideWithLayout(k)}
            onSetSlideLayout={(k) => void setSlideLayout(k)}
            onDeleteSlide={() => void deleteSlide()}
            onDuplicateSlide={() => void duplicateSlide()}
            onAddTextBox={() => void addTextBox()}
            onAddShape={(p) => void addShape(p)}
            onAddConnector={(t) => startConnectorTool(t)}
            connectorToolType={connectorTool?.type ?? null}
            onInsertImage={(f) => void insertImage(f)}
            onReplacePicture={(f) => void replaceSelectedPicture(f)}
            selectedIsPicture={selectedShape?.kind === "pic"}
            onDeleteShape={() => void deleteSelectedShape()}
            onAlign={(mode, relativeTo) => void alignSelected(mode, relativeTo)}
            onDistribute={(axis) => void distributeSelected(axis)}
            onReorder={(mode) => void reorderSelected(mode)}
            onGroup={() => void groupSelectedShapes()}
            onUngroup={() => void ungroupSelectedShape()}
            onDuplicateShape={() => void duplicateSelectedShapes()}
            canGroup={selectedShapeIds.length >= 2}
            canUngroup={selectedShapeIds.length === 1 && selectedShape?.kind === "group"}
            onChangeFill={(h) => void changeFill(h)}
            onAddComment={focusCommentComposer}
            onPresent={() => startPresenting(true)}
            canPresent={ready && slides.length > 0}
            onToggleNotes={() => setNotesOpen((v) => !v)}
            notesOpen={notesOpen}
          />
        }
        statusBarRight={
          <ZoomControl
            value={zoom}
            minZoom={MIN_ZOOM}
            maxZoom={MAX_ZOOM}
            onChange={onZoomChange}
            disabled={!ready}
          />
        }
        body={
          <div className="pptx-editor flex min-h-0 flex-1 flex-col gap-2 p-3">
            {!agent ? (
              <EmptyState product="pptx" onOpen={() => void handleOpenFile()} />
            ) : (
              <>
                <div className="relative flex min-h-0 flex-1 gap-3">
                  <aside
                    data-testid="pptx-sidebar"
                    className="hidden w-[200px] shrink-0 overflow-y-auto rounded-md border border-divider bg-surface md:block"
                  >
                    {snap ? (
                      <SlidesSidebar
                        slides={slides}
                        slideSize={slideSize}
                        mediaUrls={mediaUrls}
                        theme={themeDefault}
                        activeIndex={activeIndex}
                        onSelect={setActiveIndex}
                        thumbnailWidth={170}
                        onReorder={(from, to) => void moveSlide(from, to)}
                        onContextAction={(idx, action) => void handleSlideContextAction(idx, action)}
                      />
                    ) : null}
                  </aside>
                  <section
                    ref={slideSurfaceRef as React.RefObject<HTMLElement>}
                    tabIndex={-1}
                    data-testid="pptx-slide-surface"
                    className="relative flex min-h-0 flex-1 justify-center overflow-auto rounded-md border border-divider bg-background p-4"
                  >
                    <div className="w-full max-w-[1100px]" style={{ alignSelf: "flex-start" }}>
                      <SlideCanvas
                        agent={agent}
                        slideIndex={activeIndex}
                        mediaUrls={mediaUrls}
                        onError={onError}
                        zoom={zoom}
                        onSelectionChange={setSelectedShapeIds}
                        onTextSelectionChange={setTextSelection}
                        selectedShapeIds={selectedShapeIds}
                        commentedShapeIds={commentedShapeIds}
                        commentFlashTarget={commentFlashTarget}
                        onPlaceholderActivate={(info) => void handlePlaceholderActivate(info)}
                        connectorTool={connectorTool}
                        onConnectorToolExit={exitConnectorTool}
                      />
                      {selectedShape && selectedShape.kind === "connector" && selectedShapeIds.length === 1 ? (
                        <div className="pointer-events-auto absolute left-1/2 top-2 z-20 -translate-x-1/2">
                          <ConnectorContextBar
                            connector={selectedShape}
                            onPatch={(patch) => void applyConnectorStylePatch(selectedShape.id, patch)}
                          />
                        </div>
                      ) : null}
                    </div>
                    {!ready ? <LoadingScreen variant="splash" product="pptx" /> : null}
                  </section>
                </div>
                {snap && notesOpen ? (
                  <NotesPanel
                    key={`notes-${activeIndex}-${tick}`}
                    notesText={readNotesText(snap, activeIndex)}
                    disabled={!ready}
                    onChange={(t) => void setSlideNotes(t)}
                  />
                ) : null}
              </>
            )}
          </div>
        }
        toasts={toasts}
        onDismissToast={dismissToast}
        onFileDrop={(file) => void handleFile(file)}
        dropExtension=".pptx"
        onRenameFilename={(next) => setDocName(next)}
      />
      <KeyboardShortcutsDialog
        product="pptx"
        open={shortcutsDialog.open}
        onClose={() => shortcutsDialog.setOpen(false)}
      />
      {presenting && snap ? (
        <PresentMode
          snapshot={snap}
          initialSlideIndex={activeIndex}
          mediaUrls={mediaUrls}
          charts={snap.root.charts}
          onClose={() => setPresenting(false)}
        />
      ) : null}
    </>
  );
}

function readNotesText(snap: PptxSnapshot, slideIndex: number): string {
  const slide = snap.root.slides[slideIndex];
  if (!slide || !slide.notesSlidePartPath) return "";
  const notes = snap.root.notesSlides.get(slide.notesSlidePartPath);
  if (!notes) return "";
  return notes.body.paragraphs.map((p) => p.runs.map((r) => r.text).join("")).join("\n");
}

interface NotesPanelProps {
  notesText: string;
  disabled: boolean;
  onChange: (text: string) => void;
}

function NotesPanel(props: NotesPanelProps): React.ReactNode {
  // Local mirror of the textarea so per-keystroke typing is fluid; we
  // commit to the snapshot on blur and on a 600 ms idle (debounced)
  // rather than dispatching a command on every character — the typed
  // notes part rebuild is cheap but the React re-render churn isn't.
  const [draft, setDraft] = useState(props.notesText);
  const [dirty, setDirty] = useState(false);
  const lastSyncedRef = useRef(props.notesText);

  useEffect(() => {
    // Pull the upstream value when the slide changes (parent passes a
    // fresh `key` on slide change, so this branch is mainly defensive).
    if (lastSyncedRef.current !== props.notesText && !dirty) {
      setDraft(props.notesText);
      lastSyncedRef.current = props.notesText;
    }
  }, [props.notesText, dirty]);

  useEffect(() => {
    if (!dirty) return;
    const handle = setTimeout(() => {
      props.onChange(draft);
      lastSyncedRef.current = draft;
      setDirty(false);
    }, 600);
    return () => clearTimeout(handle);
  }, [draft, dirty, props]);

  return (
    <section
      data-testid="pptx-notes-panel"
      className="flex shrink-0 flex-col gap-1 rounded-md border border-divider bg-surface p-2"
    >
      <label className="text-xs font-medium text-secondary">Speaker notes</label>
      <textarea
        className="min-h-[64px] resize-y rounded border border-divider bg-background p-2 text-sm focus:outline-none focus:ring-1 focus:ring-[var(--accent)]"
        placeholder="Add notes for this slide…"
        disabled={props.disabled}
        value={draft}
        onChange={(e) => {
          setDraft(e.target.value);
          setDirty(true);
        }}
        onBlur={() => {
          if (!dirty) return;
          props.onChange(draft);
          lastSyncedRef.current = draft;
          setDirty(false);
        }}
      />
    </section>
  );
}

function findShape(shapes: ReadonlyArray<Shape>, id: string): Shape | null {
  for (const s of shapes) {
    if (s.id === id) return s;
    if (s.kind === "group") {
      const inner = findShape(s.children, id);
      if (inner) return inner;
    }
  }
  return null;
}

/** Walks `spPrTail` for the first `<a:solidFill><a:srgbClr val="…"/></a:solidFill>`. */
function readSolidFill(shape: TextShape | Picture): string | null {
  for (const c of shape.spPrTail) {
    if (c.tag !== "a:solidFill") continue;
    for (const inner of c.subtree) {
      if (!inner || typeof inner !== "object" || Array.isArray(inner)) continue;
      const obj = inner as Record<string, unknown>;
      const attrs = obj[":@"] as Record<string, unknown> | undefined;
      const val = attrs && typeof attrs === "object" ? attrs["@_val"] : undefined;
      if (typeof val === "string" && /^[0-9a-fA-F]{6}$/.test(val)) {
        return val.toUpperCase();
      }
    }
  }
  return null;
}

/**
 * Open the OS file picker for raster images and resolve to the chosen
 * `File` (or `null` if the user cancelled). We deliberately go via a
 * one-shot detached `<input>` rather than stashing a hidden ref on
 * the component: placeholder activation is a rare, transient flow and
 * adding a permanent hidden input + ref to the editor surface for
 * each new affordance would noisily accumulate over time.
 *
 * The MIME allowlist matches the one used by the toolbar's "Insert
 * image" input so the canvas-driven path can't sneak in a file type
 * that `pptx:insert-image` would reject downstream.
 */
function pickImageFile(): Promise<File | null> {
  return new Promise((resolve) => {
    if (typeof document === "undefined") {
      resolve(null);
      return;
    }
    const input = document.createElement("input");
    input.type = "file";
    input.accept = "image/png,image/jpeg,image/gif,image/webp,image/bmp";
    let resolved = false;
    const settle = (file: File | null): void => {
      if (resolved) return;
      resolved = true;
      resolve(file);
    };
    input.onchange = () => settle(input.files?.[0] ?? null);
    // Some browsers fire `cancel`; for the rest, the focus-back-to-window
    // heuristic (next tick after a short delay) is the only way to
    // detect a cancelled picker. We schedule a "no file picked" fallback
    // so callers don't await forever.
    input.oncancel = () => settle(null);
    window.addEventListener(
      "focus",
      () => {
        window.setTimeout(() => settle(null), 500);
      },
      { once: true }
    );
    input.click();
  });
}

async function readIntrinsicSize(
  bytes: Uint8Array,
  mime: string
): Promise<{ width: number; height: number }> {
  if (typeof window === "undefined") return { width: 0, height: 0 };
  const blob = new Blob([bytes as BlobPart], { type: mime });
  if (typeof createImageBitmap === "function" && mime !== "image/svg+xml") {
    try {
      const bitmap = await createImageBitmap(blob);
      const w = bitmap.width;
      const h = bitmap.height;
      bitmap.close?.();
      return { width: w, height: h };
    } catch {
      // fall through to <img> path
    }
  }
  return await new Promise<{ width: number; height: number }>((resolve, reject) => {
    const url = URL.createObjectURL(blob);
    const img = new Image();
    img.onload = () => {
      const out = { width: img.naturalWidth, height: img.naturalHeight };
      URL.revokeObjectURL(url);
      resolve(out);
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("Could not decode image to read intrinsic size."));
    };
    img.src = url;
  });
}
