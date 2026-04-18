import * as React from "react";
import type { PptxAgent } from "../../agent/agent.js";
import type { Shape, Slide, SlideSize, TextShape } from "../../model/types.js";
import { DEFAULT_THEME } from "../layout/color.js";
import { shapeBoundingBox, type BoundingBox } from "../layout/shape.js";
import { slideAspectRatio, slideViewBox } from "../layout/slide.js";
import { DEFAULT_DPI, EMU_PER_PX_AT_96DPI, clampZoom } from "../layout/units.js";
import type { SvgRenderCtx } from "../svg/shapes.js";
import { shapeToSvg } from "../svg/shapes.js";
import { resolveSlideBackgroundColor } from "../svg/slide.js";
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
  /**
   * Notified whenever the user selects (or deselects) a shape on the canvas.
   * Lets the parent route toolbar actions (bold/italic/underline) to the
   * selected shape instead of guessing the first text shape on the slide.
   */
  readonly onSelectionChange?: (shapeId: string | null) => void;
  /**
   * Optional controlled selection. When provided, the canvas keeps its
   * internal selection in sync with this value — useful when the parent
   * wants to auto-select a freshly-inserted shape from the toolbar so the
   * user can immediately drag, format, or delete it without an extra click.
   */
  readonly selectedShapeId?: string | null;
}

type ResizeHandle = "n" | "s" | "e" | "w" | "nw" | "ne" | "sw" | "se";
type DragMode = "move" | { resize: ResizeHandle };

interface DragState {
  readonly shapeId: string;
  readonly mode: DragMode;
  readonly startX: number;
  readonly startY: number;
  readonly origin: BoundingBox;
  readonly emuPerPx: number;
}

interface DragPreview {
  readonly box: BoundingBox;
  readonly dx: number;
  readonly dy: number;
  readonly dw: number;
  readonly dh: number;
}

export function SlideCanvas(props: SlideCanvasProps): React.ReactElement | null {
  const snap = useAgentSnapshot(props.agent);
  const slide: Slide | undefined = snap.root.slides[props.slideIndex];
  const slideSize: SlideSize = snap.root.slideSize;
  const containerRef = React.useRef<HTMLDivElement>(null);
  const [drag, setDrag] = React.useState<DragState | null>(null);
  const [preview, setPreview] = React.useState<DragPreview | null>(null);
  const [editingId, setEditingId] = React.useState<string | null>(null);
  const [selectedId, setSelectedIdState] = React.useState<string | null>(null);
  const onSelectionChange = props.onSelectionChange;
  const setSelectedId = React.useCallback(
    (next: string | null) => {
      setSelectedIdState(next);
      onSelectionChange?.(next);
    },
    [onSelectionChange]
  );

  React.useEffect(() => {
    setSelectedIdState(null);
    onSelectionChange?.(null);
  }, [props.slideIndex, onSelectionChange]);

  // Mirror the controlled prop into local state so a parent that just
  // dispatched an "insert shape" command can programmatically select it.
  // We don't echo back through `onSelectionChange` here — the parent is
  // already the source of truth — to avoid a render loop.
  const controlledSelectedId = props.selectedShapeId;
  React.useEffect(() => {
    if (controlledSelectedId === undefined) return;
    setSelectedIdState((prev) => (prev === controlledSelectedId ? prev : controlledSelectedId));
  }, [controlledSelectedId]);

  const themeDefault = snap.root.themeDefault ?? DEFAULT_THEME;
  const charts = snap.root.charts;
  const ctx: SvgRenderCtx = React.useMemo(
    () => ({ slideSize, mediaUrls: props.mediaUrls, theme: themeDefault, charts }),
    [slideSize, props.mediaUrls, themeDefault, charts]
  );

  // Hide the dragged shape from the static SVG layer while a drag is in
  // flight — the React-managed ghost mirrors it at the live position so
  // we don't have to invalidate the whole SVG string on every pointer move.
  const hiddenIds = React.useMemo(() => {
    const set = new Set<string>();
    if (editingId) set.add(editingId);
    if (drag) set.add(drag.shapeId);
    return set;
  }, [editingId, drag]);

  const svgInner = React.useMemo(() => {
    if (!slide) return "";
    return slide.shapes
      .filter((s) => !hiddenIds.has(s.id))
      .map((s) => shapeToSvg(s, ctx))
      .join("");
  }, [slide, ctx, hiddenIds]);

  // The SVG snapshot of the dragged shape, captured once at drag start so
  // we don't recompute it on every pointermove. The transform is applied
  // via the wrapping <g> so we still get sub-pixel updates without any
  // string churn.
  const dragGhostSvg = React.useMemo(() => {
    if (!drag || !slide) return "";
    const sh = findShape(slide.shapes, drag.shapeId);
    if (!sh) return "";
    return shapeToSvg(sh, ctx);
  }, [drag, slide, ctx]);

  const onPointerDown = React.useCallback(
    (e: React.PointerEvent<HTMLDivElement>) => {
      if (!slide || !containerRef.current) return;
      if (e.button !== 0) return;
      const target = e.target as Element | null;
      const handleEl = target?.closest("[data-handle]") as HTMLElement | null;
      const shapeEl = target?.closest("[data-shape-id]") as SVGGElement | null;
      const shapeId = handleEl?.dataset.shapeId ?? shapeEl?.dataset.shapeId ?? null;
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
      const handle = handleEl?.dataset.handle as ResizeHandle | "move" | undefined;
      const mode: DragMode = handle && handle !== "move" ? { resize: handle as ResizeHandle } : "move";
      setDrag({ shapeId, mode, startX: e.clientX, startY: e.clientY, origin: box, emuPerPx });
      setPreview({ box, dx: 0, dy: 0, dw: 0, dh: 0 });
      // Capture on the container — handles get unmounted/replaced when the
      // preview state updates, which would otherwise lose pointer capture
      // mid-drag and freeze the gesture.
      containerRef.current.setPointerCapture?.(e.pointerId);
      e.preventDefault();
    },
    [slide, slideSize, setSelectedId]
  );

  const onPointerMove = React.useCallback(
    (e: React.PointerEvent<HTMLDivElement>) => {
      if (!drag) return;
      const dxPx = e.clientX - drag.startX;
      const dyPx = e.clientY - drag.startY;
      const dxEmu = Math.round(dxPx * drag.emuPerPx);
      const dyEmu = Math.round(dyPx * drag.emuPerPx);
      setPreview(computePreview(drag, dxEmu, dyEmu));
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
      const final = computePreview(drag, dxEmu, dyEmu);
      const noChange = final.dx === 0 && final.dy === 0 && final.dw === 0 && final.dh === 0;
      const shapeId = drag.shapeId;
      // Clear drag/preview FIRST so the React-managed ghost vanishes the
      // moment we hand off to the snapshot — otherwise there's a flicker
      // between command apply and snapshot re-render.
      setDrag(null);
      setPreview(null);
      containerRef.current?.releasePointerCapture?.(e.pointerId);
      if (noChange) return;
      try {
        if (drag.mode === "move") {
          await props.agent.applyCommand({
            type: "pptx:set-position",
            source: "human",
            payload: {
              slideIndex: props.slideIndex,
              shapeId,
              x: final.box.x,
              y: final.box.y,
            },
          });
        } else {
          // For resize-from-anywhere we may also need to update position
          // (resizing from N/W/NW/etc shifts the origin). Issue both.
          if (final.dx !== 0 || final.dy !== 0) {
            await props.agent.applyCommand({
              type: "pptx:set-position",
              source: "human",
              payload: {
                slideIndex: props.slideIndex,
                shapeId,
                x: final.box.x,
                y: final.box.y,
              },
            });
          }
          if (final.box.cx !== drag.origin.cx || final.box.cy !== drag.origin.cy) {
            await props.agent.applyCommand({
              type: "pptx:set-size",
              source: "human",
              payload: {
                slideIndex: props.slideIndex,
                shapeId,
                width: final.box.cx,
                height: final.box.cy,
              },
            });
          }
        }
      } catch (err) {
        props.onError?.(err as Error);
      }
    },
    [drag, props]
  );

  const startEditing = React.useCallback(
    (shapeId: string) => {
      setEditingId(shapeId);
      setSelectedId(shapeId);
    },
    [setSelectedId]
  );

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
  const previewBox = preview?.box ?? null;

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
        cursor: drag ? cursorForDrag(drag.mode) : "default",
        touchAction: "none",
      }}
      onPointerDown={onPointerDown}
      onPointerMove={onPointerMove}
      onPointerUp={onPointerUp}
      onPointerCancel={onPointerUp}
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
          __html: `<rect width="100%" height="100%" fill="${slideBackgroundFillAttr(slide, themeDefault)}"/>${svgInner}${animationBadgesSvg(slide, hiddenIds)}`,
        }}
      />
      {drag && previewBox ? (
        <DragGhostSvg
          slideSize={slideSize}
          ghostSvg={dragGhostSvg}
          previewBox={previewBox}
          originBox={drag.origin}
        />
      ) : null}
      <SelectionOverlaySvg
        slide={slide}
        slideSize={slideSize}
        selectedId={selectedId}
        previewBox={previewBox}
        previewTargetId={drag?.shapeId ?? null}
      />
      {editingId ? renderEditingOverlay(slide, editingId, slideSize, dpi, finishEditing) : null}
    </div>
  );
}

function slideBackgroundFillAttr(slide: Slide, theme: typeof DEFAULT_THEME): string {
  const bg = resolveSlideBackgroundColor(slide.cSldHead, theme);
  return bg ? `#${bg}` : "white";
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

/**
 * Compute the preview box from the drag's origin + accumulated pointer
 * delta. Move drags translate; resize drags adjust the corresponding
 * edge(s) and shift the origin so the opposite edge stays anchored.
 * A small minimum size (250k EMU ≈ 26 px @ 96 DPI) keeps the shape
 * grabbable after the drag ends.
 */
function computePreview(drag: DragState, dxEmu: number, dyEmu: number): DragPreview {
  const MIN = 250_000;
  const o = drag.origin;
  if (drag.mode === "move") {
    return {
      box: { x: o.x + dxEmu, y: o.y + dyEmu, cx: o.cx, cy: o.cy },
      dx: dxEmu,
      dy: dyEmu,
      dw: 0,
      dh: 0,
    };
  }
  const h = drag.mode.resize;
  let nx = o.x;
  let ny = o.y;
  let nw = o.cx;
  let nh = o.cy;
  if (h.includes("e")) nw = Math.max(MIN, o.cx + dxEmu);
  if (h.includes("s")) nh = Math.max(MIN, o.cy + dyEmu);
  if (h.includes("w")) {
    const newCx = Math.max(MIN, o.cx - dxEmu);
    nx = o.x + (o.cx - newCx);
    nw = newCx;
  }
  if (h.includes("n")) {
    const newCy = Math.max(MIN, o.cy - dyEmu);
    ny = o.y + (o.cy - newCy);
    nh = newCy;
  }
  return {
    box: { x: nx, y: ny, cx: nw, cy: nh },
    dx: nx - o.x,
    dy: ny - o.y,
    dw: nw - o.cx,
    dh: nh - o.cy,
  };
}

function cursorForDrag(mode: DragMode): string {
  if (mode === "move") return "grabbing";
  switch (mode.resize) {
    case "n":
    case "s":
      return "ns-resize";
    case "e":
    case "w":
      return "ew-resize";
    case "ne":
    case "sw":
      return "nesw-resize";
    case "nw":
    case "se":
      return "nwse-resize";
  }
}

interface DragGhostProps {
  readonly slideSize: SlideSize;
  readonly ghostSvg: string;
  readonly previewBox: BoundingBox;
  readonly originBox: BoundingBox;
}

/**
 * Renders the dragged shape at its live position by reusing its baked
 * SVG output and applying a translate+scale via the wrapping <g>. This
 * avoids touching the static SVG string on every pointer move while
 * still giving the user pixel-perfect feedback.
 */
function DragGhostSvg({ slideSize, ghostSvg, previewBox, originBox }: DragGhostProps) {
  const tx = px(previewBox.x - originBox.x);
  const ty = px(previewBox.y - originBox.y);
  const sx = originBox.cx > 0 ? previewBox.cx / originBox.cx : 1;
  const sy = originBox.cy > 0 ? previewBox.cy / originBox.cy : 1;
  // Scale around the origin's top-left so translate+scale compose cleanly.
  const ox = px(originBox.x);
  const oy = px(originBox.y);
  // `transform` order matters: translate to the new top-left, then
  // re-anchor to the origin's top-left, scale, then anchor back. This
  // produces the visual the user expects when grabbing a side handle.
  const transform = `translate(${tx} ${ty}) translate(${ox} ${oy}) scale(${sx} ${sy}) translate(${-ox} ${-oy})`;
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      viewBox={slideViewBox(slideSize)}
      preserveAspectRatio="xMidYMid meet"
      pointerEvents="none"
      style={{
        position: "absolute",
        inset: 0,
        width: "100%",
        height: "100%",
        opacity: 0.9,
      }}
      dangerouslySetInnerHTML={{ __html: `<g transform="${transform}">${ghostSvg}</g>` }}
    />
  );
}

interface SelectionOverlayProps {
  readonly slide: Slide;
  readonly slideSize: SlideSize;
  readonly selectedId: string | null;
  readonly previewBox: BoundingBox | null;
  readonly previewTargetId: string | null;
}

/**
 * Renders the selection chrome (dashed outline + 8 resize handles) in
 * its own SVG layer so handle hit-testing is independent of the static
 * shape SVG. Handles are sized in CSS pixels so they stay clickable
 * regardless of zoom; we accomplish this by reading the container's
 * displayed width via a `vector-effect: non-scaling-stroke` on the
 * outline and an absolute pixel size on the rects.
 */
function SelectionOverlaySvg({
  slide,
  slideSize,
  selectedId,
  previewBox,
  previewTargetId,
}: SelectionOverlayProps): React.ReactElement | null {
  if (!selectedId) return null;
  const shape = findShape(slide.shapes, selectedId);
  if (!shape) return null;
  const baseBox = shapeBoundingBox(shape);
  if (!baseBox) return null;
  const box = previewTargetId === selectedId && previewBox ? previewBox : baseBox;
  const x = px(box.x);
  const y = px(box.y);
  const cx = px(box.cx);
  const cy = px(box.cy);
  const handleSize = 10; // CSS pixels — applied via `vectorEffect` below.
  const sid = escAttr(selectedId);

  const handles: ReadonlyArray<{ readonly h: ResizeHandle; readonly hx: number; readonly hy: number }> = [
    { h: "nw", hx: x, hy: y },
    { h: "n", hx: x + cx / 2, hy: y },
    { h: "ne", hx: x + cx, hy: y },
    { h: "e", hx: x + cx, hy: y + cy / 2 },
    { h: "se", hx: x + cx, hy: y + cy },
    { h: "s", hx: x + cx / 2, hy: y + cy },
    { h: "sw", hx: x, hy: y + cy },
    { h: "w", hx: x, hy: y + cy / 2 },
  ];

  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      viewBox={slideViewBox(slideSize)}
      preserveAspectRatio="xMidYMid meet"
      style={{ position: "absolute", inset: 0, width: "100%", height: "100%" }}
    >
      <rect
        x={x}
        y={y}
        width={cx}
        height={cy}
        fill="none"
        stroke="#7c3aed"
        strokeWidth={1.5}
        strokeDasharray="4 2"
        vectorEffect="non-scaling-stroke"
        pointerEvents="none"
      />
      {/* Move handle: an invisible rectangle covering the body of the
          shape so any drag inside the selection is treated as a "move",
          not as a pointerdown on the unrelated shape behind it. */}
      <rect
        data-shape-id={sid}
        data-handle="move"
        x={x}
        y={y}
        width={cx}
        height={cy}
        fill="transparent"
        style={{ cursor: "move" }}
      />
      {handles.map((it) => (
        <rect
          key={it.h}
          data-shape-id={sid}
          data-handle={it.h}
          x={it.hx}
          y={it.hy}
          width={handleSize}
          height={handleSize}
          transform={`translate(${-handleSize / 2} ${-handleSize / 2})`}
          fill="#ffffff"
          stroke="#7c3aed"
          strokeWidth={1.5}
          vectorEffect="non-scaling-stroke"
          style={{ cursor: cursorForHandle(it.h) }}
        />
      ))}
    </svg>
  );
}

function cursorForHandle(h: ResizeHandle): string {
  switch (h) {
    case "n":
    case "s":
      return "ns-resize";
    case "e":
    case "w":
      return "ew-resize";
    case "ne":
    case "sw":
      return "nesw-resize";
    case "nw":
    case "se":
      return "nwse-resize";
  }
}

/**
 * F4: render a small badge near each shape that has at least one typed
 * entrance animation. The badge shows the 1-based animation order so the
 * user can see the entrance sequence at a glance. Drawn inside the SVG
 * so it scales with the canvas, but with `pointer-events="none"` so it
 * never steals clicks from the underlying shape.
 */
function animationBadgesSvg(slide: Slide, hiddenIds: ReadonlySet<string>): string {
  if (slide.animations.length === 0) return "";
  const byCNvPrId = new Map<number, Shape>();
  collectShapesByCNvPrId(slide.shapes, byCNvPrId);
  const parts: string[] = [];
  for (const a of slide.animations) {
    const shape = byCNvPrId.get(a.targetCNvPrId);
    if (!shape) continue;
    if (hiddenIds.has(shape.id)) continue;
    const box = shapeBoundingBox(shape);
    if (!box) continue;
    const r = px(90000);
    const cx = px(box.x) + r;
    const cy = px(box.y) + r;
    const order = a.order + 1;
    parts.push(
      `<g class="anim-badge" pointer-events="none">`,
      `<circle cx="${cx}" cy="${cy}" r="${r}" fill="#facc15" stroke="#1f2937" stroke-width="${px(12000)}"/>`,
      `<text x="${cx}" y="${cy + px(36000)}" text-anchor="middle" font-size="${px(100000)}" font-family="sans-serif" font-weight="700" fill="#1f2937">${order}</text>`,
      `</g>`
    );
  }
  return parts.join("");
}

/** EMU → SVG user units (matches `shapes.ts#u`); see `slideViewBox` rationale. */
function px(emu: number): number {
  return Math.round((emu / EMU_PER_PX_AT_96DPI) * 100) / 100;
}

function collectShapesByCNvPrId(shapes: ReadonlyArray<Shape>, out: Map<number, Shape>): void {
  for (const s of shapes) {
    if (s.cNvPrId > 0 && !out.has(s.cNvPrId)) out.set(s.cNvPrId, s);
    if (s.kind === "group") collectShapesByCNvPrId(s.children, out);
  }
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
    .map((p) => p.runs.map((r) => (r.isLineBreak ? "\n" : r.text)).join(""))
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
