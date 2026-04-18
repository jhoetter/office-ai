/* eslint-disable react/no-danger */
import * as React from "react";
import type { PptxAgent } from "../../agent/agent.js";
import type { Shape, Slide, SlideSize, TextShape } from "../../model/types.js";
import { shapeBoundingBox, type BoundingBox } from "../layout/shape.js";
import { slideAspectRatio, slideViewBox } from "../layout/slide.js";
import { DEFAULT_DPI, clampZoom } from "../layout/units.js";
import type { SvgRenderCtx } from "../svg/shapes.js";
import { shapeToSvg } from "../svg/shapes.js";
import { useAgentSnapshot } from "./use-agent-snapshot.js";

export interface SlideCanvasProps {
  readonly agent: PptxAgent;
  readonly slideIndex: number;
  readonly mediaUrls?: ReadonlyMap<string, string>;
  readonly onError?: (err: Error) => void;
  /** Zoom multiplier; 1 = fit-to-container. Clamped to [0.25, 3]. */
  readonly zoom?: number;
  /** DPI used for converting EMU/font sizes to CSS pixels in the HTML overlay. */
  readonly dpi?: number;
}

interface DragState {
  readonly shapeId: string;
  readonly mode: "move" | "resize-se";
  readonly startX: number;
  readonly startY: number;
  readonly origin: BoundingBox;
  readonly emuPerPx: number;
}

export function SlideCanvas(props: SlideCanvasProps): React.ReactElement | null {
  const snap = useAgentSnapshot(props.agent);
  const slide: Slide | undefined = snap.root.slides[props.slideIndex];
  const slideSize: SlideSize = snap.root.slideSize;
  const containerRef = React.useRef<HTMLDivElement>(null);
  const [drag, setDrag] = React.useState<DragState | null>(null);
  const [editingId, setEditingId] = React.useState<string | null>(null);
  const [selectedId, setSelectedId] = React.useState<string | null>(null);

  const themeDefault = snap.root.themeDefault;
  const ctx: SvgRenderCtx = React.useMemo(
    () => ({ slideSize, mediaUrls: props.mediaUrls, theme: themeDefault }),
    [slideSize, props.mediaUrls, themeDefault]
  );

  const svgInner = React.useMemo(() => {
    if (!slide) return "";
    return slide.shapes
      .filter((s) => s.id !== editingId)
      .map((s) => shapeToSvg(s, ctx))
      .join("");
  }, [slide, ctx, editingId]);

  const onPointerDown = React.useCallback(
    (e: React.PointerEvent<HTMLDivElement>) => {
      if (!slide || !containerRef.current) return;
      const target = e.target as Element | null;
      const handleEl = target?.closest("[data-handle]") as HTMLElement | null;
      const shapeEl = target?.closest("[data-shape-id]") as SVGGElement | null;
      const shapeId =
        handleEl?.dataset.shapeId ?? shapeEl?.dataset.shapeId ?? null;
      if (!shapeId) {
        setSelectedId(null);
        return;
      }
      setSelectedId(shapeId);
      const shape = findShape(slide.shapes, shapeId);
      if (!shape) return;
      const box = shapeBoundingBox(shape);
      if (!box) return;
      const rect = containerRef.current.getBoundingClientRect();
      const emuPerPx = slideSize.cxEmu / rect.width;
      const mode: DragState["mode"] = handleEl?.dataset.handle === "resize-se" ? "resize-se" : "move";
      setDrag({ shapeId, mode, startX: e.clientX, startY: e.clientY, origin: box, emuPerPx });
      (e.target as Element).setPointerCapture?.(e.pointerId);
    },
    [slide, slideSize]
  );

  const onPointerMove = React.useCallback(
    (e: React.PointerEvent<HTMLDivElement>) => {
      if (!drag) return;
      // Visual feedback only — actual command dispatched on pointerup.
      // (We could mirror SVG transform here for live preview; deferred.)
      void e;
    },
    [drag]
  );

  const onPointerUp = React.useCallback(
    async (e: React.PointerEvent<HTMLDivElement>) => {
      if (!drag) return;
      const dxPx = e.clientX - drag.startX;
      const dyPx = e.clientY - drag.startY;
      const dxEmu = Math.round(dxPx * drag.emuPerPx);
      const dyEmu = Math.round(dyPx * drag.emuPerPx);
      try {
        if (drag.mode === "move") {
          if (dxEmu !== 0 || dyEmu !== 0) {
            await props.agent.applyCommand({
              type: "pptx:set-position",
              source: "human",
              payload: {
                slideIndex: props.slideIndex,
                shapeId: drag.shapeId,
                x: drag.origin.x + dxEmu,
                y: drag.origin.y + dyEmu,
              },
            });
          }
        } else if (drag.mode === "resize-se") {
          const newW = Math.max(100000, drag.origin.cx + dxEmu);
          const newH = Math.max(100000, drag.origin.cy + dyEmu);
          if (newW !== drag.origin.cx || newH !== drag.origin.cy) {
            await props.agent.applyCommand({
              type: "pptx:set-size",
              source: "human",
              payload: {
                slideIndex: props.slideIndex,
                shapeId: drag.shapeId,
                width: newW,
                height: newH,
              },
            });
          }
        }
      } catch (err) {
        props.onError?.(err as Error);
      } finally {
        setDrag(null);
      }
    },
    [drag, props]
  );

  const startEditing = React.useCallback((shapeId: string) => {
    setEditingId(shapeId);
    setSelectedId(shapeId);
  }, []);

  const finishEditing = React.useCallback(
    async (shape: TextShape, newText: string) => {
      setEditingId(null);
      const original = textShapePlain(shape);
      if (original === newText) return;
      try {
        await props.agent.applyCommand({
          type: "pptx:set-text",
          source: "human",
          payload: { slideIndex: props.slideIndex, shapeId: shape.id, text: newText },
        });
      } catch (err) {
        props.onError?.(err as Error);
      }
    },
    [props]
  );

  if (!slide) {
    return null;
  }
  const aspect = slideAspectRatio(slideSize);
  const zoom = clampZoom(props.zoom ?? 1);
  const dpi = props.dpi ?? DEFAULT_DPI;

  return (
    <div
      ref={containerRef}
      data-testid="pptx-slide-canvas"
      data-zoom={zoom.toFixed(2)}
      data-dpi={dpi}
      className="officeai-pptx-canvas"
      style={{
        position: "relative",
        width: `${zoom * 100}%`,
        aspectRatio: String(aspect),
        background: "white",
        boxShadow: "0 1px 3px rgba(0,0,0,0.08)",
        userSelect: "none",
      }}
      onPointerDown={onPointerDown}
      onPointerMove={onPointerMove}
      onPointerUp={onPointerUp}
      onDoubleClick={(e) => {
        const t = e.target as Element | null;
        const shapeEl = t?.closest("[data-shape-id]") as SVGGElement | null;
        const id = shapeEl?.dataset.shapeId;
        if (!id) return;
        const sh = findShape(slide.shapes, id);
        if (sh?.kind === "text") startEditing(id);
      }}
    >
      <svg
        xmlns="http://www.w3.org/2000/svg"
        viewBox={slideViewBox(slideSize)}
        preserveAspectRatio="xMidYMid meet"
        style={{ width: "100%", height: "100%", display: "block" }}
        dangerouslySetInnerHTML={{
          __html: `<rect width="100%" height="100%" fill="white"/>${svgInner}${selectionOverlaySvg(slide, selectedId)}`,
        }}
      />
      {editingId
        ? renderEditingOverlay(slide, editingId, slideSize, dpi, finishEditing)
        : null}
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

function selectionOverlaySvg(slide: Slide, selectedId: string | null): string {
  if (!selectedId) return "";
  const shape = findShape(slide.shapes, selectedId);
  if (!shape) return "";
  const box = shapeBoundingBox(shape);
  if (!box) return "";
  const handleSize = 80000;
  return [
    `<g class="selection" pointer-events="none">`,
    `<rect x="${box.x}" y="${box.y}" width="${box.cx}" height="${box.cy}" fill="none" stroke="#7c3aed" stroke-width="20000" stroke-dasharray="40000,20000"/>`,
    `</g>`,
    // SE resize handle (interactive — hit area in SVG; transformed via data-handle)
    `<rect data-shape-id="${escAttr(selectedId)}" data-handle="resize-se" x="${box.x + box.cx - handleSize / 2}" y="${box.y + box.cy - handleSize / 2}" width="${handleSize}" height="${handleSize}" fill="#7c3aed" pointer-events="all" style="cursor:nwse-resize"/>`,
  ].join("");
}

function escAttr(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/"/g, "&quot;");
}

function renderEditingOverlay(
  slide: Slide,
  shapeId: string,
  slideSize: SlideSize,
  dpi: number,
  onCommit: (shape: TextShape, text: string) => void
): React.ReactElement | null {
  const shape = findShape(slide.shapes, shapeId);
  if (!shape || shape.kind !== "text") return null;
  const box = shapeBoundingBox(shape);
  if (!box) return null;
  const initial = textShapePlain(shape);
  // Position the overlay relative to the SVG's viewBox by percentages.
  const leftPct = (box.x / slideSize.cxEmu) * 100;
  const topPct = (box.y / slideSize.cyEmu) * 100;
  const widthPct = (box.cx / slideSize.cxEmu) * 100;
  const heightPct = (box.cy / slideSize.cyEmu) * 100;
  const fontPx = estimateFontPx(shape, dpi);
  return (
    <TextEditOverlay
      key={shapeId}
      initial={initial}
      style={{
        position: "absolute",
        left: `${leftPct}%`,
        top: `${topPct}%`,
        width: `${widthPct}%`,
        height: `${heightPct}%`,
        padding: 4,
        boxSizing: "border-box",
        outline: "2px solid #7c3aed",
        background: "rgba(255,255,255,0.95)",
        fontSize: fontPx,
        fontFamily: "sans-serif",
        whiteSpace: "pre-wrap",
        overflow: "auto",
      }}
      onCommit={(t) => onCommit(shape as TextShape, t)}
    />
  );
}

interface TextEditOverlayProps {
  readonly initial: string;
  readonly style: React.CSSProperties;
  readonly onCommit: (text: string) => void;
}

function TextEditOverlay({ initial, style, onCommit }: TextEditOverlayProps): React.ReactElement {
  const ref = React.useRef<HTMLDivElement>(null);
  React.useEffect(() => {
    if (ref.current) {
      ref.current.innerText = initial;
      ref.current.focus();
      const sel = window.getSelection?.();
      if (sel && ref.current.firstChild) {
        const range = document.createRange();
        range.selectNodeContents(ref.current);
        range.collapse(false);
        sel.removeAllRanges();
        sel.addRange(range);
      }
    }
  }, [initial]);
  return (
    <div
      ref={ref}
      contentEditable
      suppressContentEditableWarning
      style={style}
      onBlur={() => onCommit(ref.current?.innerText ?? "")}
      onKeyDown={(e) => {
        if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) {
          e.preventDefault();
          ref.current?.blur();
        } else if (e.key === "Escape") {
          e.preventDefault();
          if (ref.current) ref.current.innerText = initial;
          ref.current?.blur();
        }
      }}
    />
  );
}

function textShapePlain(shape: TextShape): string {
  return shape.txBody.paragraphs
    .map((p) =>
      p.runs
        .map((r) => (r.isLineBreak ? "\n" : r.text))
        .join("")
    )
    .join("\n");
}

function estimateFontPx(shape: TextShape, dpi: number = DEFAULT_DPI): number {
  const r = shape.txBody.paragraphs[0]?.runs.find((x) => !x.isLineBreak);
  if (r?.properties.fontSizeHundredths !== undefined) {
    const pt = r.properties.fontSizeHundredths / 100;
    return (pt * dpi) / 72;
  }
  return (18 * dpi) / 72; // default 18pt
}
