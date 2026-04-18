"use client";

import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import { Loader2 } from "lucide-react";
import { cn } from "@officeai/ui";
import { PptxAgent } from "@officeai/pptx/agent";
import { SlideCanvas, SlidesSidebar } from "@officeai/pptx/renderer/react";
import { MAX_ZOOM, MIN_ZOOM, clampZoom } from "@officeai/pptx/renderer";
import type { Mutation } from "@officeai/core";
import type { Shape, ShapePreset, TextShape } from "@officeai/pptx";
import { buildSamplePptx } from "@/lib/sample-pptx";
import { dispatchToLlmPptx } from "@/lib/llm-client-pptx";
import { PptxToolbar } from "./PptxToolbar";
import { PptxAgentPanel, type PptxAgentDispatch } from "./PptxAgentPanel";

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
  const [pending, setPending] = useState<Mutation[]>([]);
  const [tick, setTick] = useState(0);
  const [toasts, setToasts] = useState<ToastMessage[]>([]);
  const toastId = useRef(0);
  const [zoom, setZoom] = useState(1);
  const [selectedShapeId, setSelectedShapeId] = useState<string | null>(null);

  const onZoomChange = useCallback((next: number) => {
    setZoom(clampZoom(next));
  }, []);

  const pushToast = useCallback((kind: ToastMessage["kind"], text: string) => {
    const id = ++toastId.current;
    setToasts((prev) => [...prev, { id, kind, text }]);
    setTimeout(() => setToasts((prev) => prev.filter((t) => t.id !== id)), 4000);
  }, []);

  const mountAgent = useCallback(async (buf: ArrayBuffer) => {
    const next = await PptxAgent.fromBuffer(buf);
    agentRef.current = next;
    setAgent(next);
    setActiveIndex(0);
    setSelectedShapeId(null);
    setPending([...next.getPendingMutations()]);
    setReady(true);
    setTick((t) => t + 1);
    next.subscribe(() => {
      setPending([...next.getPendingMutations()]);
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

  const currentFontPt = useMemo(() => {
    if (!selectedShape || selectedShape.kind !== "text") return null;
    const r = selectedShape.txBody.paragraphs[0]?.runs.find((x) => !x.isLineBreak && x.text.length > 0);
    if (!r) return null;
    if (r.properties.fontSizeHundredths == null) return null;
    return r.properties.fontSizeHundredths / 100;
  }, [selectedShape]);

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
      setSelectedShapeId(s.shapes[s.shapes.length - 1]!.id);
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
        setSelectedShapeId(s.shapes[s.shapes.length - 1]!.id);
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
        setSelectedShapeId(s.shapes[s.shapes.length - 1]!.id);
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, insertOffset, onError, pushToast, slideSize.cxEmu]
  );

  const deleteSelectedShape = useCallback(async () => {
    const a = agentRef.current;
    if (!a || !selectedShapeId) return;
    try {
      await a.applyCommand({
        type: "pptx:delete-shape",
        payload: { slideIndex: activeIndex, shapeId: selectedShapeId },
        source: "human",
      });
      setSelectedShapeId(null);
    } catch (err) {
      onError(err);
    }
  }, [activeIndex, onError, selectedShapeId]);

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

  // Picks the text shape to format: prefer the selected shape (if it is
  // text); otherwise fall back to the first non-empty text shape on the
  // slide. Allows formatting to apply to a freshly-added shape that has
  // no body runs yet (the format goes to the empty paragraph).
  const pickFormattingTarget = useCallback((): TextShape | null => {
    const a = agentRef.current;
    if (!a) return null;
    const s = a.getSnapshot().root.slides[activeIndex];
    if (!s) return null;
    const isText = (sh: Shape | null | undefined): sh is TextShape => sh?.kind === "text";
    if (selectedShapeId) {
      const sel = findShape(s.shapes, selectedShapeId);
      if (isText(sel)) return sel;
    }
    return s.shapes.find((sh): sh is TextShape => sh.kind === "text") ?? null;
  }, [activeIndex, selectedShapeId]);

  const applyFormat = useCallback(
    async (format: {
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
      color?: string;
      fontSizeHundredths?: number;
    }) => {
      const a = agentRef.current;
      if (!a) return;
      const ts = pickFormattingTarget();
      if (!ts) {
        pushToast("info", "Select a text shape first (or add one with the Text box button).");
        return;
      }
      const p = ts.txBody.paragraphs[0];
      const flatLen = p.runs.reduce((acc, r) => acc + (r.isLineBreak ? 0 : r.text.length), 0);
      try {
        await a.applyCommand({
          type: "pptx:format-text",
          payload: {
            slideIndex: activeIndex,
            shapeId: ts.id,
            range: { paragraph: 0, start: 0, end: Math.max(flatLen, 0) },
            format,
          },
          source: "human",
        });
      } catch (err) {
        onError(err);
      }
    },
    [activeIndex, onError, pickFormattingTarget, pushToast]
  );

  const toggleMark = useCallback(
    (mark: "bold" | "italic" | "underline") => applyFormat({ [mark]: true }),
    [applyFormat]
  );

  const changeTextColor = useCallback(
    (hex: string) => applyFormat({ color: hex.replace(/^#/, "").toUpperCase() }),
    [applyFormat]
  );

  const changeFontSize = useCallback(
    (pt: number) => applyFormat({ fontSizeHundredths: Math.round(pt * 100) }),
    [applyFormat]
  );

  // Routes prompts through the shared `/api/llm` bridge with `format: "pptx"`.
  // When no `OPENAI_API_KEY` is configured server-side the helper transparently
  // falls back to the in-process intent parser so the editor stays usable
  // without any env vars.
  const promptDispatch: PptxAgentDispatch = useCallback(
    async (text: string) => {
      const a = agentRef.current;
      if (!a) return;
      const result = await dispatchToLlmPptx(text, a, activeIndex);
      if (result.note) pushToast("warn", result.note);
      if (result.commands.length === 0) {
        pushToast("info", result.rationale || "No commands proposed.");
        return;
      }
      await a.applyCommands(result.commands);
      pushToast("info", `Applied ${result.commands.length} command(s).`);
    },
    [activeIndex, pushToast]
  );

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
        currentFill={currentFill ? `#${currentFill}` : null}
        currentFontPt={currentFontPt}
        onOpenFile={() => fileInputRef.current?.click()}
        onExport={() => void handleExport()}
        onAddSlide={() => void addSlide()}
        onDeleteSlide={() => void deleteSlide()}
        onDuplicateSlide={() => void duplicateSlide()}
        onAddTextBox={() => void addTextBox()}
        onAddShape={(p) => void addShape(p)}
        onInsertImage={(f) => void insertImage(f)}
        onDeleteShape={() => void deleteSelectedShape()}
        onToggleBold={() => void toggleMark("bold")}
        onToggleItalic={() => void toggleMark("italic")}
        onToggleUnderline={() => void toggleMark("underline")}
        onChangeFill={(h) => void changeFill(h)}
        onChangeTextColor={(h) => void changeTextColor(h)}
        onChangeFontSize={(pt) => void changeFontSize(pt)}
        zoom={zoom}
        minZoom={MIN_ZOOM}
        maxZoom={MAX_ZOOM}
        onZoomChange={onZoomChange}
        onZoomReset={() => setZoom(1)}
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
        <section className="relative flex min-h-0 flex-1 justify-center overflow-auto rounded-md border border-divider bg-background p-4">
          {agent ? (
            <div className="w-full max-w-[1100px]" style={{ alignSelf: "flex-start" }}>
              <SlideCanvas
                agent={agent}
                slideIndex={activeIndex}
                mediaUrls={mediaUrls}
                onError={onError}
                zoom={zoom}
                onSelectionChange={setSelectedShapeId}
                selectedShapeId={selectedShapeId}
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
        <aside
          data-testid="pptx-agent-panel"
          className="hidden w-[280px] shrink-0 flex-col gap-4 border-l border-divider pl-3 lg:flex"
        >
          <PptxAgentPanel
            ready={ready}
            pending={pending}
            onApprove={(id) => agentRef.current?.approveMutation(id)}
            onReject={(id) => agentRef.current?.rejectMutation(id)}
            onApproveAll={() => pending.forEach((m) => agentRef.current?.approveMutation(m.id))}
            onRejectAll={() => pending.forEach((m) => agentRef.current?.rejectMutation(m.id))}
            dispatch={promptDispatch}
          />
        </aside>
      </div>
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
    </div>
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
