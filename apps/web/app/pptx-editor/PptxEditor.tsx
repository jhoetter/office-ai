"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Loader2 } from "lucide-react";
import { CommentComposer, CommentsSidebar, cn } from "@officeai/ui";
import { createPptxCommentsProvider } from "./pptxCommentsProvider";
import { PptxAgent } from "@officeai/pptx/agent";
import { SlideCanvas, SlidesSidebar, type PptxTextSelection } from "@officeai/pptx/renderer/react";
import { MAX_ZOOM, MIN_ZOOM, clampZoom } from "@officeai/pptx/renderer";
import type {
  LayoutKindPayload,
  PptxSnapshot,
  Shape,
  ShapePreset,
  TextShape,
} from "@officeai/pptx";
import { buildSamplePptx } from "@/lib/sample-pptx";
import { PptxToolbar } from "./PptxToolbar";
import { computePptxActive, createPptxFormatProvider } from "./pptxFormatProvider";
import { useShortcutsDialog } from "@/lib/shortcuts/useShortcutsDialog";
import { KeyboardShortcutsDialog } from "@/lib/shortcuts/KeyboardShortcutsDialog";
import { usePptxShortcuts } from "@/lib/pptx-shortcuts";

interface ToastMessage {
  id: number;
  kind: "info" | "warn" | "error";
  text: string;
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
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const [ready, setReady] = useState(false);
  const [docName, setDocName] = useState("welcome.pptx");
  const [activeIndex, setActiveIndex] = useState(0);
  const [tick, setTick] = useState(0);
  const [toasts, setToasts] = useState<ToastMessage[]>([]);
  const toastId = useRef(0);
  const [zoom, setZoom] = useState(1);
  const [selectedShapeIds, setSelectedShapeIds] = useState<ReadonlyArray<string>>([]);
  const selectedShapeId = selectedShapeIds[0] ?? null;
  const [textSelection, setTextSelection] = useState<PptxTextSelection | null>(null);
  // Mutable mirrors used by the format provider so it can read live
  // state without triggering rerenders inside its own callbacks.
  const slideIndexRef = useRef(0);
  const textSelectionRef = useRef<PptxTextSelection | null>(null);
  const selectedShapeIdRef = useRef<string | null>(null);
  const pushToastRef = useRef<((kind: ToastMessage["kind"], text: string) => void) | null>(null);
  const slideSurfaceRef = useRef<HTMLElement | null>(null);
  const shortcutsDialog = useShortcutsDialog();

  const onZoomChange = useCallback((next: number) => {
    setZoom(clampZoom(next));
  }, []);

  const pushToast = useCallback((kind: ToastMessage["kind"], text: string) => {
    const id = ++toastId.current;
    setToasts((prev) => [...prev, { id, kind, text }]);
    setTimeout(() => setToasts((prev) => prev.filter((t) => t.id !== id)), 4000);
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
    next.subscribe(() => {
      setTick((t) => t + 1);
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

  const handleFile = useCallback(
    async (file: File) => {
      const buf = await file.arrayBuffer();
      setDocName(file.name);
      try {
        await mountAgent(buf);
        pushToast("info", `Opened ${file.name}`);
      } catch (err) {
        onError(err);
      }
    },
    [mountAgent, onError, pushToast]
  );

  const handleExport = useCallback(async () => {
    const a = agentRef.current;
    if (!a) return;
    try {
      const buf = await a.exportFile();
      const blob = new Blob([buf], {
        type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      });
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = docName;
      link.click();
      URL.revokeObjectURL(url);
      pushToast("info", `Exported ${docName}`);
    } catch (err) {
      onError(err);
    }
  }, [docName, onError, pushToast]);

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

  const currentFill = useMemo(() => {
    if (!selectedShape || selectedShape.kind !== "text") return null;
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

  const addConnector = useCallback(
    async (connectorType: "straight" | "elbow" | "curved") => {
      const a = agentRef.current;
      if (!a) return;
      try {
        const off = insertOffset();
        await a.applyCommand({
          type: "pptx:add-connector",
          payload: {
            slideIndex: activeIndex,
            connectorType,
            start: { kind: "free", xEmu: off.x, yEmu: off.y },
            end: { kind: "free", xEmu: off.x + 3_000_000, yEmu: off.y + 1_000_000 },
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

  const alignSelected = useCallback(
    async (mode: "left" | "center-h" | "right" | "top" | "middle-v" | "bottom") => {
      const a = agentRef.current;
      if (!a || selectedShapeIds.length < 2) return;
      try {
        await a.applyCommand({
          type: "pptx:align-shapes",
          payload: { slideIndex: activeIndex, shapeIds: selectedShapeIds, mode },
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

  return (
    <div className="pptx-editor flex h-full min-h-0 flex-col gap-3">
      <PptxToolbar
        disabled={!ready}
        slideCount={slides.length}
        hasSelection={selectedShapeId != null}
        selectionCount={selectedShapeIds.length}
        currentFill={currentFill ? `#${currentFill}` : null}
        textFormatProvider={textFormatProvider}
        textFormatActive={textFormatActive}
        onOpenFile={() => fileInputRef.current?.click()}
        onExport={() => void handleExport()}
        onAddSlide={() => void addSlide()}
        onAddSlideWithLayout={(k) => void addSlideWithLayout(k)}
        onSetSlideLayout={(k) => void setSlideLayout(k)}
        onDeleteSlide={() => void deleteSlide()}
        onDuplicateSlide={() => void duplicateSlide()}
        onAddTextBox={() => void addTextBox()}
        onAddShape={(p) => void addShape(p)}
        onAddConnector={(t) => void addConnector(t)}
        onInsertImage={(f) => void insertImage(f)}
        onDeleteShape={() => void deleteSelectedShape()}
        onAlign={(mode) => void alignSelected(mode)}
        onDistribute={(axis) => void distributeSelected(axis)}
        onChangeFill={(h) => void changeFill(h)}
        zoom={zoom}
        minZoom={MIN_ZOOM}
        maxZoom={MAX_ZOOM}
        onZoomChange={onZoomChange}
        onZoomReset={() => setZoom(1)}
        onOpenShortcuts={() => shortcutsDialog.setOpen(true)}
      />
      <input
        ref={fileInputRef}
        type="file"
        accept=".pptx,application/vnd.openxmlformats-officedocument.presentationml.presentation"
        className="hidden"
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) void handleFile(f);
          e.target.value = "";
        }}
      />
      <div className="flex min-h-0 flex-1 gap-3">
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
            />
          ) : null}
        </aside>
        <section
          ref={slideSurfaceRef as React.RefObject<HTMLElement>}
          tabIndex={-1}
          data-testid="pptx-slide-surface"
          className="relative flex min-h-0 flex-1 justify-center overflow-auto rounded-md border border-divider bg-background p-4"
        >
          {agent ? (
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
              />
            </div>
          ) : null}
          {!ready ? (
            <div className="absolute inset-0 flex items-center justify-center text-sm text-secondary">
              <Loader2 className="mr-2 animate-spin" size={14} />
              Loading…
            </div>
          ) : null}
        </section>
        {snap && agent ? (
          <aside
            data-testid="pptx-comments-sidebar"
            className="hidden w-[260px] shrink-0 overflow-y-auto rounded-md border border-divider bg-surface p-2 lg:block"
          >
            <CommentsSidebar
              key={`comments-${activeIndex}-${tick}`}
              provider={createPptxCommentsProvider({ agent, slideIndex: activeIndex })}
              author="You"
              emptyHint="No comments on this slide yet."
            />
            <div className="mt-2">
              <CommentComposer
                provider={createPptxCommentsProvider({ agent, slideIndex: activeIndex })}
                anchor={{
                  kind: "pptx-pin",
                  slideIndex: activeIndex,
                  xEmu: Math.round(slideSize.cxEmu / 2),
                  yEmu: Math.round(slideSize.cyEmu / 2),
                }}
                placeholder="Add a comment to this slide…"
              />
            </div>
          </aside>
        ) : null}
      </div>
      {snap ? (
        <NotesPanel
          key={`notes-${activeIndex}-${tick}`}
          notesText={readNotesText(snap, activeIndex)}
          disabled={!ready}
          onChange={(t) => void setSlideNotes(t)}
        />
      ) : null}
      <div className="pointer-events-none fixed bottom-6 left-1/2 z-50 flex -translate-x-1/2 flex-col items-center gap-2">
        {toasts.map((t) => (
          <div
            key={t.id}
            role="status"
            className={cn(
              "pointer-events-auto rounded-md border px-3 py-1.5 text-xs shadow-sm",
              t.kind === "info" && "border-divider bg-surface text-foreground",
              t.kind === "warn" && "border-[var(--warning)] bg-[var(--warning)]/10 text-[var(--warning)]",
              t.kind === "error" &&
                "border-[var(--error-border)] bg-[var(--error-bg)] text-[var(--error-text)]"
            )}
          >
            {t.text}
          </div>
        ))}
      </div>
      <KeyboardShortcutsDialog
        product="pptx"
        open={shortcutsDialog.open}
        onClose={() => shortcutsDialog.setOpen(false)}
      />
    </div>
  );
}

function readNotesText(snap: PptxSnapshot, slideIndex: number): string {
  const slide = snap.root.slides[slideIndex];
  if (!slide || !slide.notesSlidePartPath) return "";
  const notes = snap.root.notesSlides.get(slide.notesSlidePartPath);
  if (!notes) return "";
  return notes.body.paragraphs
    .map((p) => p.runs.map((r) => r.text).join(""))
    .join("\n");
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
function readSolidFill(shape: TextShape): string | null {
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
