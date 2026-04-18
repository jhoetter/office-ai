import * as React from "react";
import type { PptxAgent } from "../../agent/agent.js";
import type { Shape, Slide, SlideSize, TextShape } from "../../model/types.js";
import { DEFAULT_THEME } from "../layout/color.js";
import { boxesIntersect, shapeBoundingBox, type BoundingBox } from "../layout/shape.js";
import { slideAspectRatio, slideViewBox } from "../layout/slide.js";
import { computeSnap, type SnapGuide } from "../layout/snap.js";
import { snapToAnchor, type ShapeAnchor } from "../layout/anchors.js";
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
   * Notified whenever the user changes the on-canvas selection. Multi-
   * selection is communicated as an ordered array — the first id is the
   * "primary" selection (used by the toolbar to drive formatting & fill
   * actions). An empty array means the user clicked an empty area.
   */
  readonly onSelectionChange?: (ids: ReadonlyArray<string>) => void;
  /**
   * Optional controlled selection. When provided, the canvas keeps its
   * internal selection in sync with this value — useful when the parent
   * wants to auto-select a freshly-inserted shape from the toolbar so the
   * user can immediately drag, format, or delete it without an extra click.
   */
  readonly selectedShapeIds?: ReadonlyArray<string>;
}

type ResizeHandle = "n" | "s" | "e" | "w" | "nw" | "ne" | "sw" | "se";
type DragMode = "move" | { resize: ResizeHandle };

interface DragState {
  readonly mode: DragMode;
  /** Shapes participating in the drag with their pre-drag bounding boxes. */
  readonly targets: ReadonlyArray<{ readonly id: string; readonly origin: BoundingBox }>;
  readonly startX: number;
  readonly startY: number;
  readonly emuPerPx: number;
  /**
   * True when the primary target is a line shape — enables endpoint
   * anchor snapping during corner-handle resizes.
   */
  readonly primaryIsLine: boolean;
}

interface DragPreview {
  /** Per-shape preview boxes keyed by shape id. */
  readonly boxes: ReadonlyMap<string, BoundingBox>;
  readonly dx: number;
  readonly dy: number;
  readonly dw: number;
  readonly dh: number;
  /** Smart-guide lines to draw alongside the dragged shape(s). */
  readonly guides: ReadonlyArray<SnapGuide>;
  /**
   * Connection anchors near the active line endpoint while resizing a
   * line shape. The renderer draws every candidate as a faded ring and
   * `anchorSnap` (if set) as a solid filled dot.
   */
  readonly anchorCandidates: ReadonlyArray<ShapeAnchor>;
  readonly anchorSnap: ShapeAnchor | null;
}

interface MarqueeState {
  readonly startX: number;
  readonly startY: number;
  readonly currentX: number;
  readonly currentY: number;
  readonly emuPerPx: number;
  readonly startEmuX: number;
  readonly startEmuY: number;
}

export function SlideCanvas(props: SlideCanvasProps): React.ReactElement | null {
  const snap = useAgentSnapshot(props.agent);
  const slide: Slide | undefined = snap.root.slides[props.slideIndex];
  const slideSize: SlideSize = snap.root.slideSize;
  const containerRef = React.useRef<HTMLDivElement>(null);
  const [drag, setDrag] = React.useState<DragState | null>(null);
  const [preview, setPreview] = React.useState<DragPreview | null>(null);
  const [marquee, setMarquee] = React.useState<MarqueeState | null>(null);
  const [editingId, setEditingId] = React.useState<string | null>(null);
  const [selectedIds, setSelectedIdsState] = React.useState<ReadonlyArray<string>>([]);
  const onSelectionChange = props.onSelectionChange;
  const setSelectedIds = React.useCallback(
    (next: ReadonlyArray<string>) => {
      setSelectedIdsState(next);
      onSelectionChange?.(next);
    },
    [onSelectionChange]
  );

  React.useEffect(() => {
    setSelectedIdsState([]);
    onSelectionChange?.([]);
  }, [props.slideIndex, onSelectionChange]);

  // Mirror the controlled prop so a parent that just dispatched an
  // "insert shape" command can programmatically select the new shape.
  // We don't echo back through `onSelectionChange` here — the parent is
  // the source of truth — to avoid a render loop.
  const controlledSelectedIds = props.selectedShapeIds;
  React.useEffect(() => {
    if (controlledSelectedIds === undefined) return;
    setSelectedIdsState((prev) => (sameIds(prev, controlledSelectedIds) ? prev : controlledSelectedIds));
  }, [controlledSelectedIds]);

  const themeDefault = snap.root.themeDefault ?? DEFAULT_THEME;
  const charts = snap.root.charts;
  const ctx: SvgRenderCtx = React.useMemo(
    () => ({ slideSize, mediaUrls: props.mediaUrls, theme: themeDefault, charts }),
    [slideSize, props.mediaUrls, themeDefault, charts]
  );

  // Hide every shape currently being dragged or text-edited from the
  // static SVG layer — the React-managed ghost layer mirrors them at
  // their live positions so we don't have to rebuild the SVG string on
  // every pointermove.
  const hiddenIds = React.useMemo(() => {
    const set = new Set<string>();
    if (editingId) set.add(editingId);
    if (drag) for (const t of drag.targets) set.add(t.id);
    return set;
  }, [editingId, drag]);

  const svgInner = React.useMemo(() => {
    if (!slide) return "";
    return slide.shapes
      .filter((s) => !hiddenIds.has(s.id))
      .map((s) => shapeToSvg(s, ctx))
      .join("");
  }, [slide, ctx, hiddenIds]);

  // Capture the SVG snapshot of every dragged shape at drag start so we
  // don't recompute it on every pointermove. The transform is applied
  // via the wrapping <g> for sub-pixel updates without string churn.
  const dragGhosts = React.useMemo(() => {
    if (!drag || !slide) return [] as ReadonlyArray<{ id: string; svg: string; origin: BoundingBox }>;
    const out: { id: string; svg: string; origin: BoundingBox }[] = [];
    for (const t of drag.targets) {
      const sh = findShape(slide.shapes, t.id);
      if (!sh) continue;
      out.push({ id: t.id, svg: shapeToSvg(sh, ctx), origin: t.origin });
    }
    return out;
  }, [drag, slide, ctx]);

  const onPointerDown = React.useCallback(
    (e: React.PointerEvent<HTMLDivElement>) => {
      if (!slide || !containerRef.current) return;
      if (e.button !== 0) return;
      const target = e.target as Element | null;
      const handleEl = target?.closest("[data-handle]") as HTMLElement | null;
      const shapeEl = target?.closest("[data-shape-id]") as SVGGElement | null;
      const shapeId = handleEl?.dataset.shapeId ?? shapeEl?.dataset.shapeId ?? null;
      const rect = containerRef.current.getBoundingClientRect();
      const emuPerPx = slideSize.cxEmu / rect.width;
      const shiftHeld = e.shiftKey || e.metaKey || e.ctrlKey;

      if (!shapeId) {
        // Empty-area click: clear selection (unless shift) and start a
        // marquee so the user can rubber-band-select multiple shapes.
        if (!shiftHeld) setSelectedIds([]);
        setMarquee({
          startX: e.clientX,
          startY: e.clientY,
          currentX: e.clientX,
          currentY: e.clientY,
          emuPerPx,
          startEmuX: (e.clientX - rect.left) * emuPerPx,
          startEmuY: (e.clientY - rect.top) * emuPerPx,
        });
        containerRef.current.setPointerCapture?.(e.pointerId);
        e.preventDefault();
        return;
      }

      // Compute the next selection set up-front so the drag uses it.
      let nextSelection: ReadonlyArray<string>;
      if (shiftHeld) {
        const has = selectedIds.includes(shapeId);
        nextSelection = has ? selectedIds.filter((id) => id !== shapeId) : [...selectedIds, shapeId];
      } else if (selectedIds.includes(shapeId) && selectedIds.length > 1) {
        // Clicking a shape already in a multi-selection keeps the group
        // intact so the user can drag the whole thing; reorder so the
        // clicked shape becomes the primary.
        nextSelection = [shapeId, ...selectedIds.filter((id) => id !== shapeId)];
      } else {
        nextSelection = [shapeId];
      }
      setSelectedIds(nextSelection);

      // Resize handles only ever apply to the primary shape; group
      // resize is not supported in this iteration. Move handles drag
      // every shape in the (post-shift) selection together.
      const handle = handleEl?.dataset.handle as ResizeHandle | "move" | undefined;
      const isResize = !!handle && handle !== "move";
      const dragIds = isResize ? [shapeId] : nextSelection;
      const targets: { id: string; origin: BoundingBox }[] = [];
      for (const id of dragIds) {
        const sh = findShape(slide.shapes, id);
        if (!sh) continue;
        const box = shapeBoundingBox(sh);
        if (!box) continue;
        targets.push({ id, origin: box });
      }
      if (targets.length === 0) return;

      const mode: DragMode = isResize ? { resize: handle as ResizeHandle } : "move";
      const primaryShape = findShape(slide.shapes, targets[0].id);
      const next: DragState = {
        mode,
        targets,
        startX: e.clientX,
        startY: e.clientY,
        emuPerPx,
        primaryIsLine: primaryShape ? isLineShape(primaryShape) : false,
      };
      setDrag(next);
      setPreview(computePreview(next, 0, 0, null));
      // Capture on the container — handles get unmounted/replaced when
      // the preview state updates, which would otherwise lose pointer
      // capture mid-drag and freeze the gesture.
      containerRef.current.setPointerCapture?.(e.pointerId);
      e.preventDefault();
    },
    [slide, slideSize, setSelectedIds, selectedIds]
  );

  const onPointerMove = React.useCallback(
    (e: React.PointerEvent<HTMLDivElement>) => {
      if (drag) {
        const dxEmu = Math.round((e.clientX - drag.startX) * drag.emuPerPx);
        const dyEmu = Math.round((e.clientY - drag.startY) * drag.emuPerPx);
        const snapCtx = makeSnapContext(drag, slide, slideSize);
        setPreview(computePreview(drag, dxEmu, dyEmu, snapCtx));
        return;
      }
      if (marquee) {
        setMarquee({ ...marquee, currentX: e.clientX, currentY: e.clientY });
        return;
      }
    },
    [drag, marquee, slide, slideSize]
  );

  const onPointerUp = React.useCallback(
    async (e: React.PointerEvent<HTMLDivElement>) => {
      containerRef.current?.releasePointerCapture?.(e.pointerId);
      if (marquee && slide) {
        const rect = containerRef.current?.getBoundingClientRect();
        if (rect) {
          const x0 = Math.min(marquee.startX, e.clientX) - rect.left;
          const y0 = Math.min(marquee.startY, e.clientY) - rect.top;
          const x1 = Math.max(marquee.startX, e.clientX) - rect.left;
          const y1 = Math.max(marquee.startY, e.clientY) - rect.top;
          const region: BoundingBox = {
            x: x0 * marquee.emuPerPx,
            y: y0 * marquee.emuPerPx,
            cx: (x1 - x0) * marquee.emuPerPx,
            cy: (y1 - y0) * marquee.emuPerPx,
          };
          // Tiny marquees (< 4 px) are treated as plain clicks; we already
          // cleared the selection on pointerdown.
          if (region.cx > 4 * marquee.emuPerPx || region.cy > 4 * marquee.emuPerPx) {
            const hits: string[] = [];
            for (const sh of slide.shapes) {
              const b = shapeBoundingBox(sh);
              if (!b) continue;
              if (boxesIntersect(b, region)) hits.push(sh.id);
            }
            const shiftHeld = e.shiftKey || e.metaKey || e.ctrlKey;
            const next = shiftHeld ? unionIds(selectedIds, hits) : hits;
            setSelectedIds(next);
          }
        }
        setMarquee(null);
        return;
      }
      if (!drag) return;
      const dxEmu = Math.round((e.clientX - drag.startX) * drag.emuPerPx);
      const dyEmu = Math.round((e.clientY - drag.startY) * drag.emuPerPx);
      const snapCtx = makeSnapContext(drag, slide, slideSize);
      const final = computePreview(drag, dxEmu, dyEmu, snapCtx);
      const noChange = final.dx === 0 && final.dy === 0 && final.dw === 0 && final.dh === 0;
      const targets = drag.targets;
      const mode = drag.mode;
      // Clear drag/preview FIRST so the React-managed ghost vanishes the
      // moment we hand off to the snapshot — otherwise there's a flicker
      // between command apply and snapshot re-render.
      setDrag(null);
      setPreview(null);
      if (noChange) return;
      try {
        for (const t of targets) {
          const box = final.boxes.get(t.id);
          if (!box) continue;
          if (mode === "move") {
            await props.agent.applyCommand({
              type: "pptx:set-position",
              source: "human",
              payload: {
                slideIndex: props.slideIndex,
                shapeId: t.id,
                x: box.x,
                y: box.y,
              },
            });
          } else {
            // Resize-from-anywhere may also need a position update (n/w/nw
            // edges shift the origin). Issue both when needed.
            if (box.x !== t.origin.x || box.y !== t.origin.y) {
              await props.agent.applyCommand({
                type: "pptx:set-position",
                source: "human",
                payload: { slideIndex: props.slideIndex, shapeId: t.id, x: box.x, y: box.y },
              });
            }
            if (box.cx !== t.origin.cx || box.cy !== t.origin.cy) {
              await props.agent.applyCommand({
                type: "pptx:set-size",
                source: "human",
                payload: {
                  slideIndex: props.slideIndex,
                  shapeId: t.id,
                  width: box.cx,
                  height: box.cy,
                },
              });
            }
          }
        }
      } catch (err) {
        props.onError?.(err as Error);
      }
    },
    [drag, marquee, props, selectedIds, setSelectedIds, slide, slideSize]
  );

  const startEditing = React.useCallback(
    (shapeId: string) => {
      setEditingId(shapeId);
      setSelectedIds([shapeId]);
    },
    [setSelectedIds]
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
      {drag && preview ? (
        <DragGhostLayer slideSize={slideSize} ghosts={dragGhosts} preview={preview} />
      ) : null}
      {drag && preview && preview.guides.length > 0 ? (
        <SmartGuidesOverlay slideSize={slideSize} guides={preview.guides} />
      ) : null}
      {drag && preview && preview.anchorCandidates.length > 0 ? (
        <AnchorOverlay
          slideSize={slideSize}
          candidates={preview.anchorCandidates}
          snapped={preview.anchorSnap}
        />
      ) : null}
      <SelectionOverlaySvg
        slide={slide}
        slideSize={slideSize}
        selectedIds={selectedIds}
        previewBoxes={preview?.boxes ?? null}
        dragMode={drag?.mode ?? null}
      />
      {marquee && containerRef.current ? (
        <MarqueeOverlay marquee={marquee} containerRect={containerRef.current.getBoundingClientRect()} />
      ) : null}
      {editingId ? renderEditingOverlay(slide, editingId, slideSize, dpi, finishEditing) : null}
    </div>
  );
}

function slideBackgroundFillAttr(slide: Slide, theme: typeof DEFAULT_THEME): string {
  const bg = resolveSlideBackgroundColor(slide.cSldHead, theme);
  return bg ? `#${bg}` : "white";
}

/**
 * Collect every top-level shape's bounding box, skipping the ones
 * currently being dragged (we don't want a shape to snap to itself).
 * Group children are intentionally excluded — the smart-snap heuristic
 * only considers slide-level peers, which keeps perceived alignment
 * obvious for the user.
 */
function collectOtherBoxes(
  shapes: ReadonlyArray<Shape>,
  excludeIds: ReadonlySet<string>
): { id: string; box: BoundingBox }[] {
  const out: { id: string; box: BoundingBox }[] = [];
  for (const sh of shapes) {
    if (excludeIds.has(sh.id)) continue;
    const b = shapeBoundingBox(sh);
    if (!b) continue;
    out.push({ id: sh.id, box: b });
  }
  return out;
}

/**
 * Detect whether a shape is a `prstGeom` line (or arrow-style line).
 * We peek at its `spPrTail` for the `<a:prstGeom prst="line">` marker
 * since the model layer doesn't expose preset as a typed field —
 * everything decorative lives in the opaque tail to keep byte-roundtrip
 * cheap.
 */
function isLineShape(shape: Shape): boolean {
  if (shape.kind !== "text") return false;
  for (const c of shape.spPrTail) {
    if (c.tag !== "a:prstGeom") continue;
    const prst = c.attrs?.prst ?? c.rawAttrs?.["@_prst"];
    if (prst === "line") return true;
  }
  return false;
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

function sameIds(a: ReadonlyArray<string>, b: ReadonlyArray<string>): boolean {
  if (a.length !== b.length) return false;
  for (let i = 0; i < a.length; i++) if (a[i] !== b[i]) return false;
  return true;
}

function unionIds(a: ReadonlyArray<string>, b: ReadonlyArray<string>): string[] {
  const seen = new Set(a);
  const out = [...a];
  for (const id of b) {
    if (!seen.has(id)) {
      seen.add(id);
      out.push(id);
    }
  }
  return out;
}

interface SnapContext {
  readonly slideSize: SlideSize;
  readonly others: ReadonlyArray<{ id: string; box: BoundingBox }>;
  readonly thresholdEmu: number;
  /**
   * Threshold for connection-anchor snapping (used when resizing a
   * line endpoint). Slightly larger than `thresholdEmu` because
   * snapping a tiny line endpoint to a fixed dot benefits from a more
   * forgiving target.
   */
  readonly anchorThresholdEmu: number;
}

function makeSnapContext(
  drag: DragState,
  slide: Slide | undefined,
  slideSize: SlideSize
): SnapContext | null {
  if (!slide) return null;
  // We always build a SnapContext when a slide is present; computePreview
  // decides per-mode whether to consume the alignment guides vs the
  // anchor candidates. Threshold ≈ 6 CSS pixels at the current zoom —
  // matches Figma's "feel" of a magnetic gutter while staying off the
  // way when the user is intentionally placing a shape. Anchors get a
  // slightly larger 10 px threshold because they're discrete targets.
  return {
    slideSize,
    others: collectOtherBoxes(slide.shapes, new Set(drag.targets.map((t) => t.id))),
    thresholdEmu: Math.round(6 * drag.emuPerPx),
    anchorThresholdEmu: Math.round(10 * drag.emuPerPx),
  };
}

/**
 * Compute per-shape preview boxes for the current drag. Move drags
 * translate every target by the same delta and consult the smart-snap
 * helper to nudge the delta toward neighbouring shape edges/centres.
 * Resize drags only apply to the primary (single) target — group
 * resize is intentionally out-of-scope for this iteration. A small
 * minimum size (250k EMU ≈ 26 px @ 96 DPI) keeps shapes grabbable
 * after the gesture ends.
 */
function computePreview(
  drag: DragState,
  dxEmu: number,
  dyEmu: number,
  snap: SnapContext | null
): DragPreview {
  const MIN = 250_000;
  const boxes = new Map<string, BoundingBox>();
  if (drag.mode === "move") {
    // For multi-shape drags we snap the UNION box and apply the same
    // delta to every shape. This matches Figma's "the group as a whole
    // aligns with this neighbour" behaviour.
    const unionOrigin = unionOf(drag.targets.map((t) => t.origin));
    const proposedUnion: BoundingBox = {
      x: unionOrigin.x + dxEmu,
      y: unionOrigin.y + dyEmu,
      cx: unionOrigin.cx,
      cy: unionOrigin.cy,
    };
    let snapDx = 0;
    let snapDy = 0;
    let guides: ReadonlyArray<SnapGuide> = [];
    if (snap) {
      const r = computeSnap(proposedUnion, snap.others, snap.slideSize, snap.thresholdEmu);
      snapDx = r.snapDx;
      snapDy = r.snapDy;
      guides = r.guides;
    }
    const totalDx = dxEmu + snapDx;
    const totalDy = dyEmu + snapDy;
    for (const t of drag.targets) {
      boxes.set(t.id, {
        x: t.origin.x + totalDx,
        y: t.origin.y + totalDy,
        cx: t.origin.cx,
        cy: t.origin.cy,
      });
    }
    return {
      boxes,
      dx: totalDx,
      dy: totalDy,
      dw: 0,
      dh: 0,
      guides,
      anchorCandidates: [],
      anchorSnap: null,
    };
  }
  const t = drag.targets[0];
  const o = t.origin;
  const h = drag.mode.resize;
  // Lines have a single useful dimension (cx OR cy); the standard MIN
  // would force a minimum cy on a horizontal line, which would visually
  // "thicken" it. Use 0 as the floor for lines and let the user drag
  // freely.
  const minSize = drag.primaryIsLine ? 0 : MIN;
  let nx = o.x;
  let ny = o.y;
  let nw = o.cx;
  let nh = o.cy;
  if (h.includes("e")) nw = Math.max(minSize, o.cx + dxEmu);
  if (h.includes("s")) nh = Math.max(minSize, o.cy + dyEmu);
  if (h.includes("w")) {
    const newCx = Math.max(minSize, o.cx - dxEmu);
    nx = o.x + (o.cx - newCx);
    nw = newCx;
  }
  if (h.includes("n")) {
    const newCy = Math.max(minSize, o.cy - dyEmu);
    ny = o.y + (o.cy - newCy);
    nh = newCy;
  }

  // Anchor snap: lines have endpoints at the NW (start) and SE (end)
  // corners of their bounding box (the parser/serializer pair handles
  // flipH/flipV transparently). When the user drags one of those
  // corners we try to glue it onto a nearby shape's anchor.
  let anchorCandidates: ReadonlyArray<ShapeAnchor> = [];
  let anchorSnap: ShapeAnchor | null = null;
  if (drag.primaryIsLine && snap && (h === "nw" || h === "se" || h === "ne" || h === "sw")) {
    const endpoint = endpointForHandle(h, { x: nx, y: ny, cx: nw, cy: nh });
    const r = snapToAnchor(endpoint, snap.others, snap.anchorThresholdEmu);
    anchorCandidates = r.nearby;
    anchorSnap = r.anchor;
    if (r.anchor) {
      const adj = applyAnchorDelta(h, { x: nx, y: ny, cx: nw, cy: nh }, r.dx, r.dy);
      nx = adj.x;
      ny = adj.y;
      nw = adj.cx;
      nh = adj.cy;
    }
  }

  boxes.set(t.id, { x: nx, y: ny, cx: nw, cy: nh });
  return {
    boxes,
    dx: nx - o.x,
    dy: ny - o.y,
    dw: nw - o.cx,
    dh: nh - o.cy,
    guides: [],
    anchorCandidates,
    anchorSnap,
  };
}

function endpointForHandle(
  h: ResizeHandle,
  box: BoundingBox
): { x: number; y: number } {
  switch (h) {
    case "nw":
      return { x: box.x, y: box.y };
    case "ne":
      return { x: box.x + box.cx, y: box.y };
    case "sw":
      return { x: box.x, y: box.y + box.cy };
    case "se":
      return { x: box.x + box.cx, y: box.y + box.cy };
    case "n":
    case "s":
    case "e":
    case "w":
      return { x: box.x + box.cx / 2, y: box.y + box.cy / 2 };
  }
}

function applyAnchorDelta(
  h: ResizeHandle,
  box: BoundingBox,
  dx: number,
  dy: number
): BoundingBox {
  // Translating a corner means changing both that corner's coordinate
  // AND the bounding box dimension on that axis (the opposite corner
  // stays anchored).
  let { x, y, cx, cy } = box;
  if (h === "nw") {
    x += dx;
    y += dy;
    cx -= dx;
    cy -= dy;
  } else if (h === "ne") {
    y += dy;
    cx += dx;
    cy -= dy;
  } else if (h === "sw") {
    x += dx;
    cx -= dx;
    cy += dy;
  } else if (h === "se") {
    cx += dx;
    cy += dy;
  }
  return { x, y, cx, cy };
}

function unionOf(boxes: ReadonlyArray<BoundingBox>): BoundingBox {
  let x = boxes[0].x;
  let y = boxes[0].y;
  let r = boxes[0].x + boxes[0].cx;
  let b = boxes[0].y + boxes[0].cy;
  for (let i = 1; i < boxes.length; i++) {
    if (boxes[i].x < x) x = boxes[i].x;
    if (boxes[i].y < y) y = boxes[i].y;
    if (boxes[i].x + boxes[i].cx > r) r = boxes[i].x + boxes[i].cx;
    if (boxes[i].y + boxes[i].cy > b) b = boxes[i].y + boxes[i].cy;
  }
  return { x, y, cx: r - x, cy: b - y };
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

interface DragGhostLayerProps {
  readonly slideSize: SlideSize;
  readonly ghosts: ReadonlyArray<{ id: string; svg: string; origin: BoundingBox }>;
  readonly preview: DragPreview;
}

/**
 * Renders every dragged shape at its live position by reusing each
 * shape's baked SVG output and applying a translate (+ scale, when
 * resizing) via a wrapping <g>. This avoids re-emitting the static SVG
 * string on every pointer move while still giving the user pixel-
 * perfect feedback even for multi-shape group drags.
 */
function DragGhostLayer({ slideSize, ghosts, preview }: DragGhostLayerProps): React.ReactElement {
  const inner = ghosts
    .map((g) => {
      const box = preview.boxes.get(g.id);
      if (!box) return "";
      const tx = px(box.x - g.origin.x);
      const ty = px(box.y - g.origin.y);
      const sx = g.origin.cx > 0 ? box.cx / g.origin.cx : 1;
      const sy = g.origin.cy > 0 ? box.cy / g.origin.cy : 1;
      const ox = px(g.origin.x);
      const oy = px(g.origin.y);
      const transform = `translate(${tx} ${ty}) translate(${ox} ${oy}) scale(${sx} ${sy}) translate(${-ox} ${-oy})`;
      return `<g transform="${transform}">${g.svg}</g>`;
    })
    .join("");
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
      dangerouslySetInnerHTML={{ __html: inner }}
    />
  );
}

interface SelectionOverlayProps {
  readonly slide: Slide;
  readonly slideSize: SlideSize;
  readonly selectedIds: ReadonlyArray<string>;
  readonly previewBoxes: ReadonlyMap<string, BoundingBox> | null;
  readonly dragMode: DragMode | null;
}

/**
 * Renders selection chrome in its own SVG layer so handle hit-testing
 * is independent of the static shape SVG. For a single selection we
 * draw the dashed outline + 8 resize handles. For multi-selection we
 * draw a per-shape outline plus a union outline; each shape gets a
 * "move" hit zone so dragging anywhere inside the group moves the
 * whole thing. Resize handles are intentionally suppressed for multi-
 * selections — group resize is out-of-scope for this iteration.
 */
function SelectionOverlaySvg({
  slide,
  slideSize,
  selectedIds,
  previewBoxes,
  dragMode,
}: SelectionOverlayProps): React.ReactElement | null {
  if (selectedIds.length === 0) return null;
  const entries: { id: string; box: BoundingBox }[] = [];
  for (const id of selectedIds) {
    const sh = findShape(slide.shapes, id);
    if (!sh) continue;
    const base = shapeBoundingBox(sh);
    if (!base) continue;
    const previewBox = previewBoxes?.get(id) ?? null;
    entries.push({ id, box: previewBox ?? base });
  }
  if (entries.length === 0) return null;
  const isMulti = entries.length > 1;
  const isResizing = dragMode !== null && dragMode !== "move";
  // For single-shape resizes the union/handles should track the live
  // preview box; for moves they follow the single shape too.
  const primary = entries[0];
  const handleSize = 10;

  const handles: { readonly h: ResizeHandle; readonly hx: number; readonly hy: number }[] = [];
  if (!isMulti) {
    const x = px(primary.box.x);
    const y = px(primary.box.y);
    const cx = px(primary.box.cx);
    const cy = px(primary.box.cy);
    handles.push(
      { h: "nw", hx: x, hy: y },
      { h: "n", hx: x + cx / 2, hy: y },
      { h: "ne", hx: x + cx, hy: y },
      { h: "e", hx: x + cx, hy: y + cy / 2 },
      { h: "se", hx: x + cx, hy: y + cy },
      { h: "s", hx: x + cx / 2, hy: y + cy },
      { h: "sw", hx: x, hy: y + cy },
      { h: "w", hx: x, hy: y + cy / 2 }
    );
  }

  const unionBox = entries.reduce<BoundingBox>(
    (acc, e, i) => {
      if (i === 0) return e.box;
      const x = Math.min(acc.x, e.box.x);
      const y = Math.min(acc.y, e.box.y);
      const right = Math.max(acc.x + acc.cx, e.box.x + e.box.cx);
      const bottom = Math.max(acc.y + acc.cy, e.box.y + e.box.cy);
      return { x, y, cx: right - x, cy: bottom - y };
    },
    entries[0].box
  );
  const ux = px(unionBox.x);
  const uy = px(unionBox.y);
  const ucx = px(unionBox.cx);
  const ucy = px(unionBox.cy);

  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      viewBox={slideViewBox(slideSize)}
      preserveAspectRatio="xMidYMid meet"
      style={{ position: "absolute", inset: 0, width: "100%", height: "100%" }}
    >
      {/* Per-shape outlines (visible in multi-select; redundant in single
          select but kept consistent for clarity). */}
      {isMulti
        ? entries.map((e) => (
            <rect
              key={`outline-${e.id}`}
              x={px(e.box.x)}
              y={px(e.box.y)}
              width={px(e.box.cx)}
              height={px(e.box.cy)}
              fill="none"
              stroke="#7c3aed"
              strokeOpacity={0.5}
              strokeWidth={1}
              strokeDasharray="3 2"
              vectorEffect="non-scaling-stroke"
              pointerEvents="none"
            />
          ))
        : null}
      {/* Union / single-shape outline */}
      <rect
        x={ux}
        y={uy}
        width={ucx}
        height={ucy}
        fill="none"
        stroke="#7c3aed"
        strokeWidth={1.5}
        strokeDasharray={isMulti ? "6 3" : "4 2"}
        vectorEffect="non-scaling-stroke"
        pointerEvents="none"
      />
      {/* Move hit-zones — one per selected shape so a click on any of
          them moves the entire group. We deliberately use per-shape
          rects (not the union) so the user can click between two
          selected shapes without grabbing the gap. */}
      {entries.map((e) => (
        <rect
          key={`move-${e.id}`}
          data-shape-id={escAttr(e.id)}
          data-handle="move"
          x={px(e.box.x)}
          y={px(e.box.y)}
          width={px(e.box.cx)}
          height={px(e.box.cy)}
          fill="transparent"
          style={{ cursor: "move" }}
        />
      ))}
      {/* Resize handles (single-select only). During a resize gesture we
          keep them rendered so the user sees the corner they're pulling. */}
      {handles.map((it) => (
        <rect
          key={it.h}
          data-shape-id={escAttr(primary.id)}
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
          opacity={isResizing ? 0.6 : 1}
        />
      ))}
    </svg>
  );
}

interface SmartGuidesOverlayProps {
  readonly slideSize: SlideSize;
  readonly guides: ReadonlyArray<SnapGuide>;
}

/**
 * Renders Figma-style alignment guides while a shape is being dragged.
 * Vertical guides span between the topmost and bottommost involved
 * shape edges (so the user sees exactly which two shapes are being
 * tied together); slide-derived guides span the full slide. We use a
 * non-scaling stroke so the line stays 1px regardless of zoom.
 */
function SmartGuidesOverlay({ slideSize, guides }: SmartGuidesOverlayProps): React.ReactElement {
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      viewBox={slideViewBox(slideSize)}
      preserveAspectRatio="xMidYMid meet"
      pointerEvents="none"
      style={{ position: "absolute", inset: 0, width: "100%", height: "100%" }}
    >
      {guides.map((g, i) => {
        const colour = g.kind === "slide" ? "#0ea5e9" : g.kind === "center" ? "#f43f5e" : "#22c55e";
        if (g.axis === "vertical") {
          const y1 = isFinite(g.spanStart) ? px(g.spanStart) : 0;
          const y2 = isFinite(g.spanEnd) ? px(g.spanEnd) : px(slideSize.cyEmu);
          return (
            <line
              key={`v-${i}`}
              x1={px(g.value)}
              y1={y1}
              x2={px(g.value)}
              y2={y2}
              stroke={colour}
              strokeWidth={1}
              strokeDasharray="3 2"
              vectorEffect="non-scaling-stroke"
            />
          );
        }
        const x1 = isFinite(g.spanStart) ? px(g.spanStart) : 0;
        const x2 = isFinite(g.spanEnd) ? px(g.spanEnd) : px(slideSize.cxEmu);
        return (
          <line
            key={`h-${i}`}
            x1={x1}
            y1={px(g.value)}
            x2={x2}
            y2={px(g.value)}
            stroke={colour}
            strokeWidth={1}
            strokeDasharray="3 2"
            vectorEffect="non-scaling-stroke"
          />
        );
      })}
    </svg>
  );
}

interface AnchorOverlayProps {
  readonly slideSize: SlideSize;
  readonly candidates: ReadonlyArray<ShapeAnchor>;
  readonly snapped: ShapeAnchor | null;
}

/**
 * Renders connection-anchor dots while a line endpoint is being
 * dragged near other shapes. Candidate anchors appear as small hollow
 * rings; the one we'd snap to (if the user releases now) renders as a
 * solid filled dot. Sky-blue matches our slide-snap guides for visual
 * consistency.
 */
function AnchorOverlay({ slideSize, candidates, snapped }: AnchorOverlayProps): React.ReactElement {
  const r = 80_000; // ≈ 8.4 px @ 96 DPI; readable but not invasive
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      viewBox={slideViewBox(slideSize)}
      preserveAspectRatio="xMidYMid meet"
      pointerEvents="none"
      style={{ position: "absolute", inset: 0, width: "100%", height: "100%" }}
    >
      {candidates.map((a, i) => {
        const isSnapped = snapped !== null && a.shapeId === snapped.shapeId && a.side === snapped.side;
        return (
          <circle
            key={`${a.shapeId}-${a.side}-${i}`}
            cx={px(a.x)}
            cy={px(a.y)}
            r={px(r)}
            fill={isSnapped ? "#0ea5e9" : "white"}
            stroke="#0ea5e9"
            strokeWidth={1.5}
            vectorEffect="non-scaling-stroke"
          />
        );
      })}
    </svg>
  );
}

interface MarqueeOverlayProps {
  readonly marquee: MarqueeState;
  readonly containerRect: DOMRect;
}

/** Light translucent rectangle drawn while the user rubber-bands. */
function MarqueeOverlay({ marquee, containerRect }: MarqueeOverlayProps): React.ReactElement {
  const x0 = Math.min(marquee.startX, marquee.currentX) - containerRect.left;
  const y0 = Math.min(marquee.startY, marquee.currentY) - containerRect.top;
  const x1 = Math.max(marquee.startX, marquee.currentX) - containerRect.left;
  const y1 = Math.max(marquee.startY, marquee.currentY) - containerRect.top;
  return (
    <div
      style={{
        position: "absolute",
        left: x0,
        top: y0,
        width: x1 - x0,
        height: y1 - y0,
        background: "rgba(124,58,237,0.12)",
        border: "1px solid rgba(124,58,237,0.6)",
        pointerEvents: "none",
      }}
    />
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
