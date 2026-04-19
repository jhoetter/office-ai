import * as React from "react";
import type { PptxAgent } from "../../agent/agent.js";
import type { Shape, Slide, SlideSize, TextShape } from "../../model/types.js";
import { DEFAULT_THEME } from "../layout/color.js";
import { boxesIntersect, shapeBoundingBox, type BoundingBox } from "../layout/shape.js";
import { slideAspectRatio, slideViewBox } from "../layout/slide.js";
import { computeSnap, type SnapGuide } from "../layout/snap.js";
import { anchorsFor, snapToAnchor, type AnchorSide, type ShapeAnchor } from "../layout/anchors.js";
import { DEFAULT_DPI, EMU_PER_PX_AT_96DPI, clampZoom } from "../layout/units.js";
import type { SvgRenderCtx } from "../svg/shapes.js";
import { shapeToSvg } from "../svg/shapes.js";
import { buildShapesByCNvPrId, resolveSlideBackgroundColor } from "../svg/slide.js";
import { useAgentSnapshot } from "./use-agent-snapshot.js";

/**
 * Live caret/selection inside the text edit overlay.
 *
 * - `shapeId` identifies the text shape currently being edited.
 * - `paragraph` is the 0-based paragraph index inside that shape.
 * - `start` / `end` are character offsets into the paragraph's flat
 *   text (sum of all run text + line-break placeholder lengths).
 *   `start === end` means the caret is collapsed.
 *
 * The shared `pptxFormatProvider` reads this every render to compute
 * `ActiveTextFormat` and to dispatch `pptx:format-text` patches with
 * the correct range. Cleared (`null`) when the overlay isn't open or
 * the user blurs to a non-toolbar element.
 */
export interface PptxTextSelection {
  readonly shapeId: string;
  readonly paragraph: number;
  readonly start: number;
  readonly end: number;
}

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
  /**
   * Notified whenever the live text-editing selection changes (caret
   * moves, selection expands, overlay closes). `null` when no shape is
   * being edited. Used by the shared text-format toolbar to drive
   * MIXED-state pickers and the `pptx:format-text` dispatch range.
   */
  readonly onTextSelectionChange?: (sel: PptxTextSelection | null) => void;
  /**
   * Shape ids that carry an unresolved comment thread. The canvas
   * paints a soft yellow outline + corner badge over each matching
   * shape so the user can spot anchored conversations at a glance.
   * The parent owns the resolved/unresolved logic; the canvas just
   * draws what it's told.
   */
  readonly commentedShapeIds?: ReadonlyArray<string>;
  /**
   * One-shot "look here" pulse triggered by the comments sidebar's
   * "click to locate" affordance. The canvas paints a yellow flash
   * over the requested shape (or pin position when the comment isn't
   * shape-anchored), keyed off `nonce` so re-clicking the same
   * comment re-triggers the animation. Cleared by the parent or by
   * the canvas's own timer.
   */
  readonly commentFlashTarget?:
    | { readonly kind: "shape"; readonly shapeId: string; readonly nonce: number }
    | { readonly kind: "pin"; readonly xEmu: number; readonly yEmu: number; readonly nonce: number }
    | null;
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
  /**
   * True when the pointerdown landed on a text shape that was already
   * the sole selection before this click. PointerUp uses this together
   * with `noChange` to decide whether the gesture was actually a
   * "second click" — which enters text-edit mode, matching PowerPoint
   * (one click selects, one more click enters edit).
   */
  readonly mayEnterEditOnClick: boolean;
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

/**
 * Live state while the user is drawing a brand-new connector by
 * dragging from one of a shape's port dots. The source endpoint is
 * always anchored (you can only start a connector by clicking a port);
 * the destination endpoint becomes anchored if the pointer ends near
 * another shape's port, otherwise it lands as a free EMU coordinate.
 *
 * Scoped purely to the canvas — committed to the agent on pointerup
 * via `pptx:add-connector`, then reset.
 */
interface ConnectorDraft {
  readonly source: {
    readonly shapeId: string;
    readonly cNvPrId: number;
    readonly side: "n" | "s" | "e" | "w" | "center";
    readonly x: number;
    readonly y: number;
  };
  /** Live cursor position in slide EMU coordinates. */
  readonly currentX: number;
  readonly currentY: number;
  readonly emuPerPx: number;
  /** Anchor we'd snap the destination endpoint onto if released now. */
  readonly snapped: ShapeAnchor | null;
  readonly nearby: ReadonlyArray<ShapeAnchor>;
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
  const [hoveredShapeId, setHoveredShapeId] = React.useState<string | null>(null);
  const [connectorDraft, setConnectorDraft] = React.useState<ConnectorDraft | null>(null);
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
  const shapesByCNvPrId = React.useMemo(() => {
    return slide ? buildShapesByCNvPrId(slide.shapes) : new Map();
  }, [slide]);
  const ctx: SvgRenderCtx = React.useMemo(
    () => ({ slideSize, mediaUrls: props.mediaUrls, theme: themeDefault, charts, shapesByCNvPrId }),
    [slideSize, props.mediaUrls, themeDefault, charts, shapesByCNvPrId]
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
  //
  // We also keep a reference to the live `Shape` so the ghost layer can
  // re-emit text shapes at the live preview size during a resize. A
  // simple `scale(sx sy)` transform is fine for rectangles/ellipses
  // (the geometry stretches and looks identical to the post-commit
  // result), but text rendered via `<foreignObject>` would visually
  // stretch the glyphs — instead we want the text to reflow at the
  // new width, matching PowerPoint/Google Slides behaviour.
  const dragGhosts = React.useMemo(() => {
    if (!drag || !slide)
      return [] as ReadonlyArray<{
        id: string;
        svg: string;
        origin: BoundingBox;
        shape: Shape;
      }>;
    const out: {
      id: string;
      svg: string;
      origin: BoundingBox;
      shape: Shape;
    }[] = [];
    for (const t of drag.targets) {
      const sh = findShape(slide.shapes, t.id);
      if (!sh) continue;
      out.push({ id: t.id, svg: shapeToSvg(sh, ctx), origin: t.origin, shape: sh });
    }
    return out;
  }, [drag, slide, ctx]);

  const onPointerDown = React.useCallback(
    (e: React.PointerEvent<HTMLDivElement>) => {
      if (!slide || !containerRef.current) return;
      if (e.button !== 0) return;
      const target = e.target as Element | null;
      const portEl = target?.closest("[data-port]") as HTMLElement | null;
      const handleEl = target?.closest("[data-handle]") as HTMLElement | null;
      const shapeEl = target?.closest("[data-shape-id]") as SVGGElement | null;
      const shapeId =
        portEl?.dataset.portShapeId ?? handleEl?.dataset.shapeId ?? shapeEl?.dataset.shapeId ?? null;
      const rect = containerRef.current.getBoundingClientRect();
      const emuPerPx = slideSize.cxEmu / rect.width;
      const shiftHeld = e.shiftKey || e.metaKey || e.ctrlKey;

      // Port click: start a brand-new connector. This takes precedence
      // over the regular select/move/resize gesture so a hovered shape
      // surface stays draggable everywhere except the four port dots.
      if (portEl && shapeId) {
        const side = portEl.dataset.port as "n" | "s" | "e" | "w" | "center" | undefined;
        const sourceShape = findShape(slide.shapes, shapeId);
        if (side && sourceShape) {
          const box = shapeBoundingBox(sourceShape);
          const cNvPrId = sourceShape.cNvPrId;
          if (box && cNvPrId > 0) {
            const ap = anchorPointFor(box, side);
            const cursorEmuX = (e.clientX - rect.left) * emuPerPx;
            const cursorEmuY = (e.clientY - rect.top) * emuPerPx;
            setHoveredShapeId(null);
            setConnectorDraft({
              source: { shapeId, cNvPrId, side, x: ap.x, y: ap.y },
              currentX: cursorEmuX,
              currentY: cursorEmuY,
              emuPerPx,
              snapped: null,
              nearby: [],
            });
            containerRef.current.setPointerCapture?.(e.pointerId);
            e.preventDefault();
            e.stopPropagation();
            return;
          }
        }
      }

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

      // Resize handles drag the primary shape on single-select, and
      // proportionally scale every selected shape relative to the union
      // bounding box on multi-select. Move handles drag every shape in
      // the (post-shift) selection together.
      const handle = handleEl?.dataset.handle as ResizeHandle | "move" | undefined;
      const isResize = !!handle && handle !== "move";
      const dragIds = isResize ? (nextSelection.length > 1 ? nextSelection : [shapeId]) : nextSelection;
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
      // Click-to-edit eligibility: this click landed on the only shape
      // that was already selected (no shift, no resize, no edit overlay
      // currently open), and the target is a text shape. PointerUp will
      // enter edit mode if the gesture didn't turn into a drag.
      const mayEnterEditOnClick =
        !isResize &&
        !shiftHeld &&
        editingId === null &&
        selectedIds.length === 1 &&
        selectedIds[0] === shapeId &&
        primaryShape?.kind === "text";
      const next: DragState = {
        mode,
        targets,
        startX: e.clientX,
        startY: e.clientY,
        emuPerPx,
        primaryIsLine: primaryShape ? isLineShape(primaryShape) : false,
        mayEnterEditOnClick,
      };
      setDrag(next);
      setPreview(computePreview(next, 0, 0, null));
      // Capture on the container — handles get unmounted/replaced when
      // the preview state updates, which would otherwise lose pointer
      // capture mid-drag and freeze the gesture.
      containerRef.current.setPointerCapture?.(e.pointerId);
      e.preventDefault();
    },
    [slide, slideSize, setSelectedIds, selectedIds, editingId]
  );

  const onPointerMove = React.useCallback(
    (e: React.PointerEvent<HTMLDivElement>) => {
      if (connectorDraft) {
        const rect = containerRef.current?.getBoundingClientRect();
        if (!rect || !slide) return;
        const cursorEmuX = (e.clientX - rect.left) * connectorDraft.emuPerPx;
        const cursorEmuY = (e.clientY - rect.top) * connectorDraft.emuPerPx;
        const others = collectOtherBoxes(slide.shapes, new Set([connectorDraft.source.shapeId]));
        const snap = snapToAnchor({ x: cursorEmuX, y: cursorEmuY }, others, ANCHOR_THRESHOLD_EMU);
        setConnectorDraft({
          ...connectorDraft,
          currentX: cursorEmuX,
          currentY: cursorEmuY,
          snapped: snap.anchor,
          nearby: snap.nearby,
        });
        return;
      }
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
      // Idle hover: surface port dots over the shape currently under
      // the pointer so the user discovers connector creation. Skip
      // connectors themselves (they don't carry attachment ports) and
      // groups (their children carry the ports).
      const target = e.target as Element | null;
      const shapeEl = target?.closest("[data-shape-id]") as SVGGElement | null;
      const id = shapeEl?.dataset.shapeId ?? null;
      if (id === hoveredShapeId) return;
      if (!id) {
        setHoveredShapeId(null);
        return;
      }
      const sh = slide ? findShape(slide.shapes, id) : null;
      if (!sh || sh.kind === "connector" || sh.kind === "group" || sh.cNvPrId <= 0) {
        setHoveredShapeId(null);
        return;
      }
      setHoveredShapeId(id);
    },
    [connectorDraft, drag, marquee, slide, slideSize, hoveredShapeId]
  );

  const onPointerUp = React.useCallback(
    async (e: React.PointerEvent<HTMLDivElement>) => {
      containerRef.current?.releasePointerCapture?.(e.pointerId);
      if (connectorDraft && slide) {
        const draft = connectorDraft;
        setConnectorDraft(null);
        const startEndDistEmu = Math.hypot(draft.currentX - draft.source.x, draft.currentY - draft.source.y);
        // Tiny drags (< ~6 px) are treated as accidental clicks and
        // discarded so the user doesn't end up with a pile of zero-
        // length connectors when they merely tap a port.
        if (startEndDistEmu < 6 * draft.emuPerPx) {
          return;
        }
        try {
          const targetCNvPrId =
            draft.snapped !== null ? findCNvPrIdByShapeId(slide.shapes, draft.snapped.shapeId) : null;
          const endPayload =
            draft.snapped !== null && targetCNvPrId !== null
              ? {
                  kind: "anchored" as const,
                  targetCNvPrId,
                  side: draft.snapped.side,
                }
              : { kind: "free" as const, xEmu: draft.currentX, yEmu: draft.currentY };
          await props.agent.applyCommand({
            type: "pptx:add-connector",
            source: "human",
            payload: {
              slideIndex: props.slideIndex,
              connectorType: "elbow",
              start: {
                kind: "anchored",
                targetCNvPrId: draft.source.cNvPrId,
                side: draft.source.side,
              },
              end: endPayload,
            },
          });
        } catch (err) {
          props.onError?.(err as Error);
        }
        return;
      }
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
      const mayEnterEditOnClick = drag.mayEnterEditOnClick;
      // Clear drag/preview FIRST so the React-managed ghost vanishes the
      // moment we hand off to the snapshot — otherwise there's a flicker
      // between command apply and snapshot re-render.
      setDrag(null);
      setPreview(null);
      if (noChange) {
        // PowerPoint-style "second click enters edit": if the gesture
        // started on the only-selected text shape and the user didn't
        // actually drag, open the editing overlay so the user doesn't
        // have to hunt for a double-click target.
        if (mayEnterEditOnClick && targets[0]) {
          setEditingId(targets[0].id);
        }
        return;
      }
      try {
        for (const t of targets) {
          const box = final.boxes.get(t.id);
          if (!box) continue;
          // Connector resize hits a different command path: dragging a
          // corner handle moves one endpoint, and the snap result tells
          // us whether to record an anchored or free endpoint. The
          // serialised set-position/set-size commands wouldn't capture
          // the new anchor relationship.
          const draggedShape = slide ? findShape(slide.shapes, t.id) : null;
          const resizeHandle =
            mode !== "move" && typeof mode === "object" && "resize" in mode
              ? (mode as { resize: ResizeHandle }).resize
              : null;
          if (draggedShape?.kind === "connector" && resizeHandle && isCornerHandle(resizeHandle) && slide) {
            const which = endpointForHandleSide(resizeHandle);
            const snappedAnchor = final.anchorSnap;
            const targetCNvPrId = snappedAnchor
              ? findCNvPrIdByShapeId(slide.shapes, snappedAnchor.shapeId)
              : null;
            const endpointPt = endpointForHandle(resizeHandle, box);
            const endpointPayload =
              snappedAnchor && targetCNvPrId !== null
                ? {
                    kind: "anchored" as const,
                    targetCNvPrId,
                    side: snappedAnchor.side,
                  }
                : { kind: "free" as const, xEmu: endpointPt.x, yEmu: endpointPt.y };
            await props.agent.applyCommand({
              type: "pptx:set-connector-endpoint",
              source: "human",
              payload: {
                slideIndex: props.slideIndex,
                shapeId: t.id,
                which,
                endpoint: endpointPayload,
              },
            });
            continue;
          }
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
    [connectorDraft, drag, marquee, props, selectedIds, setSelectedIds, slide, slideSize]
  );

  const onTextSelectionChange = props.onTextSelectionChange;

  const startEditing = React.useCallback(
    (shapeId: string) => {
      setEditingId(shapeId);
      setSelectedIds([shapeId]);
    },
    [setSelectedIds]
  );

  const finishEditing = React.useCallback(
    async (
      shape: TextShape,
      payloadParagraphs: ReadonlyArray<{
        readonly runs: ReadonlyArray<{
          readonly text: string;
          readonly isLineBreak?: boolean;
          readonly inheritFromRun?: number;
        }>;
      }>,
      plain: string
    ) => {
      setEditingId(null);
      onTextSelectionChange?.(null);
      const original = textShapePlain(shape);
      if (original === plain) return;
      try {
        await props.agent.applyCommand({
          type: "pptx:set-text",
          source: "human",
          payload: {
            slideIndex: props.slideIndex,
            shapeId: shape.id,
            paragraphs: payloadParagraphs,
          },
        });
      } catch (err) {
        props.onError?.(err as Error);
      }
    },
    [props, onTextSelectionChange]
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
        <DragGhostLayer slideSize={slideSize} ghosts={dragGhosts} preview={preview} ctx={ctx} />
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
      {props.commentedShapeIds && props.commentedShapeIds.length > 0 ? (
        <CommentMarkerOverlay slide={slide} slideSize={slideSize} shapeIds={props.commentedShapeIds} />
      ) : null}
      {props.commentFlashTarget ? (
        <CommentFlashOverlay
          key={`flash-${props.commentFlashTarget.nonce}`}
          slide={slide}
          slideSize={slideSize}
          target={props.commentFlashTarget}
        />
      ) : null}
      <SelectionOverlaySvg
        slide={slide}
        slideSize={slideSize}
        selectedIds={selectedIds}
        previewBoxes={preview?.boxes ?? null}
        dragMode={drag?.mode ?? null}
        editingId={editingId}
      />
      {marquee && containerRef.current ? (
        <MarqueeOverlay marquee={marquee} containerRect={containerRef.current.getBoundingClientRect()} />
      ) : null}
      {/* Port-hover layer — surfaced when the user is idling over a
          non-connector shape, so the four cardinal anchor dots become
          drag-from sources for new connectors. Suppressed during any
          active drag/marquee/draft to keep the canvas calm. */}
      {!drag && !marquee && !connectorDraft && hoveredShapeId
        ? renderPortHoverOverlay(slide, slideSize, hoveredShapeId)
        : null}
      {connectorDraft ? <ConnectorDraftOverlay slideSize={slideSize} draft={connectorDraft} /> : null}
      {editingId
        ? renderEditingOverlay(slide, editingId, slideSize, dpi, finishEditing, onTextSelectionChange)
        : null}
    </div>
  );
}

function renderPortHoverOverlay(slide: Slide, slideSize: SlideSize, shapeId: string): React.ReactNode {
  const sh = findShape(slide.shapes, shapeId);
  if (!sh) return null;
  const box = shapeBoundingBox(sh);
  if (!box) return null;
  return <PortHoverOverlay slideSize={slideSize} shapeId={shapeId} box={box} />;
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
function isCornerHandle(h: ResizeHandle): boolean {
  return h === "nw" || h === "ne" || h === "sw" || h === "se";
}

/**
 * Map a resize handle to which connector endpoint (start vs end) it
 * controls. The convention mirrors how the parser/serializer place
 * endpoints on a connector's bounding box: start = nw corner, end =
 * se corner. ne/sw map the same way (top-right ≈ start because
 * connectors with `flipH` set start on the right, but we don't track
 * flip in the canvas — both handles edit the same endpoint as if
 * flip were absent, which lets the reflow logic re-pick flips).
 */
function endpointForHandleSide(h: ResizeHandle): "start" | "end" {
  switch (h) {
    case "nw":
    case "ne":
      return "start";
    case "sw":
    case "se":
      return "end";
    default:
      return "end";
  }
}

function findCNvPrIdByShapeId(shapes: ReadonlyArray<Shape>, id: string): number | null {
  for (const s of shapes) {
    if (s.id === id) return s.cNvPrId > 0 ? s.cNvPrId : null;
    if (s.kind === "group") {
      const inner = findCNvPrIdByShapeId(s.children, id);
      if (inner !== null) return inner;
    }
  }
  return null;
}

function isLineShape(shape: Shape): boolean {
  if (shape.kind === "connector") return true;
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

/**
 * Snap radius (in EMU) for the destination endpoint of a draft
 * connector. The user is hand-targeting a port dot, so we use a more
 * forgiving threshold than the regular drag-snap — roughly twice the
 * visual port radius. The constant is `emuPerPx`-independent: 100 000
 * EMU ≈ 10.5 px @ 96 DPI, which feels right at every zoom level when
 * you remember `currentX/Y` are already in slide coordinates.
 */
const ANCHOR_THRESHOLD_EMU = 100_000;

/** Slide-coordinate location of a side anchor on a bounding box. */
function anchorPointFor(box: BoundingBox, side: AnchorSide): { x: number; y: number } {
  const all = anchorsFor("", box);
  for (const a of all) if (a.side === side) return { x: a.x, y: a.y };
  return { x: box.x + box.cx / 2, y: box.y + box.cy / 2 };
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
  const h = drag.mode.resize;

  // Multi-shape group resize: scale every selected shape relative to
  // the union bounding box. Each shape's box gets its position offset
  // and dimensions multiplied by the union's resize factor, matching
  // PowerPoint / Figma group resize semantics. Anchor-snap and per-shape
  // line behaviour are intentionally skipped here — they would require
  // per-shape policy decisions that don't translate cleanly to a group.
  if (drag.targets.length > 1) {
    const union = unionOf(drag.targets.map((t) => t.origin));
    let nxU = union.x;
    let nyU = union.y;
    let nwU = union.cx;
    let nhU = union.cy;
    if (h.includes("e")) nwU = Math.max(MIN, union.cx + dxEmu);
    if (h.includes("s")) nhU = Math.max(MIN, union.cy + dyEmu);
    if (h.includes("w")) {
      const newCx = Math.max(MIN, union.cx - dxEmu);
      nxU = union.x + (union.cx - newCx);
      nwU = newCx;
    }
    if (h.includes("n")) {
      const newCy = Math.max(MIN, union.cy - dyEmu);
      nyU = union.y + (union.cy - newCy);
      nhU = newCy;
    }
    const sx = union.cx === 0 ? 1 : nwU / union.cx;
    const sy = union.cy === 0 ? 1 : nhU / union.cy;
    for (const t of drag.targets) {
      const o = t.origin;
      boxes.set(t.id, {
        x: nxU + Math.round((o.x - union.x) * sx),
        y: nyU + Math.round((o.y - union.y) * sy),
        cx: Math.max(0, Math.round(o.cx * sx)),
        cy: Math.max(0, Math.round(o.cy * sy)),
      });
    }
    return {
      boxes,
      dx: nxU - union.x,
      dy: nyU - union.y,
      dw: nwU - union.cx,
      dh: nhU - union.cy,
      guides: [],
      anchorCandidates: [],
      anchorSnap: null,
    };
  }

  const t = drag.targets[0];
  const o = t.origin;
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

function endpointForHandle(h: ResizeHandle, box: BoundingBox): { x: number; y: number } {
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

function applyAnchorDelta(h: ResizeHandle, box: BoundingBox, dx: number, dy: number): BoundingBox {
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
  readonly ghosts: ReadonlyArray<{
    id: string;
    svg: string;
    origin: BoundingBox;
    shape: Shape;
  }>;
  readonly preview: DragPreview;
  readonly ctx: SvgRenderCtx;
}

/**
 * Renders every dragged shape at its live position. For most shape
 * kinds we reuse each shape's baked SVG output and apply a translate
 * (+ scale, when resizing) via a wrapping <g> — this avoids re-emitting
 * the static SVG string on every pointer move while still giving the
 * user pixel-perfect feedback even for multi-shape group drags.
 *
 * Text shapes are special-cased: they render via `<foreignObject>` +
 * HTML, and a CSS scale would visually stretch the glyphs instead of
 * reflowing the text. To match PowerPoint/Google Slides behaviour
 * during a resize, we re-emit the text shape's SVG at the live
 * preview position+size on every pointer move so word-wrap kicks in
 * at the new width. `shapeToSvg` is just string concat, so the cost
 * is negligible compared to the visual quality win.
 */
function DragGhostLayer({ slideSize, ghosts, preview, ctx }: DragGhostLayerProps): React.ReactElement {
  const inner = ghosts
    .map((g) => {
      const box = preview.boxes.get(g.id);
      if (!box) return "";
      if (g.shape.kind === "text") {
        const synth: TextShape = {
          ...g.shape,
          position: { xEmu: box.x, yEmu: box.y },
          size: { cxEmu: box.cx, cyEmu: box.cy },
        };
        return shapeToSvg(synth, ctx);
      }
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
  /**
   * The shape currently being text-edited, if any. The TextEditOverlay
   * draws its own dashed border in the same purple, so duplicating it
   * here (plus the 8 resize handles) creates the "chunky textarea"
   * look we explicitly want to avoid — PowerPoint suppresses the
   * shape chrome while the caret is in the text body.
   */
  readonly editingId: string | null;
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
  editingId,
}: SelectionOverlayProps): React.ReactElement | null {
  if (selectedIds.length === 0) return null;
  // Suppress chrome (outline + handles + move zone) for the shape being
  // text-edited. The TextEditOverlay paints a dashed border itself and
  // the contenteditable owns hit-testing, so any extra rect here would
  // either visually double the border or steal pointer events away
  // from the caret.
  const visibleIds = editingId ? selectedIds.filter((id) => id !== editingId) : selectedIds;
  if (visibleIds.length === 0) return null;
  const entries: { id: string; box: BoundingBox }[] = [];
  for (const id of visibleIds) {
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
  // Visible handle is 12 user units (≈12 CSS px @ 96 DPI / 100 % zoom).
  // The hit-zone is intentionally larger than the visible square so the
  // user doesn't have to land precisely on a tiny target — matches what
  // Figma / PowerPoint do (hit slop ~6 px on every side).
  const handleSize = 12;
  const handleHitSize = 24;

  const unionBox = entries.reduce<BoundingBox>((acc, e, i) => {
    if (i === 0) return e.box;
    const x = Math.min(acc.x, e.box.x);
    const y = Math.min(acc.y, e.box.y);
    const right = Math.max(acc.x + acc.cx, e.box.x + e.box.cx);
    const bottom = Math.max(acc.y + acc.cy, e.box.y + e.box.cy);
    return { x, y, cx: right - x, cy: bottom - y };
  }, entries[0].box);
  const ux = px(unionBox.x);
  const uy = px(unionBox.y);
  const ucx = px(unionBox.cx);
  const ucy = px(unionBox.cy);

  // Resize handles wrap the union box, so single-shape and multi-shape
  // selections share the same control affordances. The handle's
  // data-shape-id targets the primary so single-shape gestures keep
  // their existing per-shape command path; multi-shape gestures recruit
  // every selected shape via dragIds in onPointerDown.
  const handles: { readonly h: ResizeHandle; readonly hx: number; readonly hy: number }[] = [
    { h: "nw", hx: ux, hy: uy },
    { h: "n", hx: ux + ucx / 2, hy: uy },
    { h: "ne", hx: ux + ucx, hy: uy },
    { h: "e", hx: ux + ucx, hy: uy + ucy / 2 },
    { h: "se", hx: ux + ucx, hy: uy + ucy },
    { h: "s", hx: ux + ucx / 2, hy: uy + ucy },
    { h: "sw", hx: ux, hy: uy + ucy },
    { h: "w", hx: ux, hy: uy + ucy / 2 },
  ];

  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      viewBox={slideViewBox(slideSize)}
      preserveAspectRatio="xMidYMid meet"
      // The overlay <svg> must not absorb clicks itself — otherwise a
      // shift-click on a non-selected shape stacked underneath the selected
      // one's bbox would target the overlay root instead of the underlying
      // shape, breaking multi-selection. Children that need to receive
      // events (move zones, resize handles) opt back in explicitly below.
      pointerEvents="none"
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
          pointerEvents="auto"
          style={{ cursor: "move" }}
        />
      ))}
      {/* Resize handles. On multi-select they wrap the union box and
          scale every selected shape proportionally; on single-select
          they wrap the shape itself. During a resize gesture we keep
          them rendered so the user sees the corner they're pulling.
          We render two rects per handle: an invisible larger hit-zone
          on top so pointer capture is forgiving, and the visible
          purple-bordered square underneath. */}
      {handles.map((it) => (
        <g key={it.h}>
          <rect
            x={it.hx}
            y={it.hy}
            width={handleSize}
            height={handleSize}
            transform={`translate(${-handleSize / 2} ${-handleSize / 2})`}
            fill="#ffffff"
            stroke="#7c3aed"
            strokeWidth={1.5}
            vectorEffect="non-scaling-stroke"
            pointerEvents="none"
            opacity={isResizing ? 0.6 : 1}
          />
          <rect
            data-shape-id={escAttr(primary.id)}
            data-handle={it.h}
            x={it.hx}
            y={it.hy}
            width={handleHitSize}
            height={handleHitSize}
            transform={`translate(${-handleHitSize / 2} ${-handleHitSize / 2})`}
            fill="transparent"
            pointerEvents="auto"
            style={{ cursor: cursorForHandle(it.h) }}
          />
        </g>
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

interface PortHoverOverlayProps {
  readonly slideSize: SlideSize;
  readonly shapeId: string;
  readonly box: BoundingBox;
}

/**
 * Renders the four cardinal port dots over a hovered shape so the
 * user can drag-from one of them to start a brand-new connector. The
 * dots live in their own `pointerEvents: "all"` SVG layer and carry
 * `data-port`/`data-port-shape-id` attributes; the canvas pointerdown
 * handler picks them up before falling through to the regular
 * select/move gesture. The centre is intentionally NOT exposed as a
 * port — clicking the middle of a shape should select it, not start
 * a connector.
 */
function PortHoverOverlay({ slideSize, shapeId, box }: PortHoverOverlayProps): React.ReactElement {
  const sides: AnchorSide[] = ["n", "s", "e", "w"];
  const anchors = anchorsFor(shapeId, box).filter((a) => sides.includes(a.side));
  const r = 65_000; // ≈ 6.8 px @ 96 DPI; visible but unobtrusive
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      viewBox={slideViewBox(slideSize)}
      preserveAspectRatio="xMidYMid meet"
      style={{ position: "absolute", inset: 0, width: "100%", height: "100%", pointerEvents: "none" }}
    >
      {anchors.map((a) => (
        <g key={a.side}>
          <circle
            cx={px(a.x)}
            cy={px(a.y)}
            r={px(r * 1.6)}
            fill="rgba(14,165,233,0.15)"
            stroke="none"
            data-port={a.side}
            data-port-shape-id={shapeId}
            style={{ pointerEvents: "all", cursor: "crosshair" }}
          />
          <circle
            cx={px(a.x)}
            cy={px(a.y)}
            r={px(r)}
            fill="white"
            stroke="#0ea5e9"
            strokeWidth={1.5}
            vectorEffect="non-scaling-stroke"
            pointerEvents="none"
          />
        </g>
      ))}
    </svg>
  );
}

interface ConnectorDraftOverlayProps {
  readonly slideSize: SlideSize;
  readonly draft: ConnectorDraft;
}

/**
 * Live preview of a connector being drawn from a port. Renders a
 * dashed orthogonal polyline from the source anchor to the current
 * pointer position. When a candidate destination anchor is in range
 * we paint it as a solid sky-blue dot to confirm the snap target so
 * the user knows the destination will be anchored (not free).
 */
function ConnectorDraftOverlay({ slideSize, draft }: ConnectorDraftOverlayProps): React.ReactElement {
  const sx = draft.source.x;
  const sy = draft.source.y;
  const ex = draft.snapped ? draft.snapped.x : draft.currentX;
  const ey = draft.snapped ? draft.snapped.y : draft.currentY;
  const startSide = draft.source.side;
  const endSide = draft.snapped?.side ?? null;
  const points = routeElbowPoints({ x: sx, y: sy }, { x: ex, y: ey }, startSide, endSide);
  const pts = points.map((p) => `${px(p.x)},${px(p.y)}`).join(" ");
  const r = 80_000;
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      viewBox={slideViewBox(slideSize)}
      preserveAspectRatio="xMidYMid meet"
      pointerEvents="none"
      style={{ position: "absolute", inset: 0, width: "100%", height: "100%" }}
    >
      <polyline
        points={pts}
        fill="none"
        stroke="#0ea5e9"
        strokeWidth={2}
        strokeDasharray="6 4"
        vectorEffect="non-scaling-stroke"
      />
      <circle
        cx={px(sx)}
        cy={px(sy)}
        r={px(r)}
        fill="#0ea5e9"
        stroke="white"
        strokeWidth={1.5}
        vectorEffect="non-scaling-stroke"
      />
      {draft.nearby.map((a, i) => {
        const isSnap =
          draft.snapped !== null && a.shapeId === draft.snapped.shapeId && a.side === draft.snapped.side;
        return (
          <circle
            key={`${a.shapeId}-${a.side}-${i}`}
            cx={px(a.x)}
            cy={px(a.y)}
            r={px(r)}
            fill={isSnap ? "#0ea5e9" : "white"}
            stroke="#0ea5e9"
            strokeWidth={1.5}
            vectorEffect="non-scaling-stroke"
          />
        );
      })}
    </svg>
  );
}

/**
 * Local mirror of `routeElbow` from `renderer/svg/shapes.ts` so the
 * draft preview matches the committed connector's path exactly. We
 * avoid importing the SVG-string version because that one returns
 * pre-stringified markup; here we just need the raw points.
 */
function routeElbowPoints(
  sp: { x: number; y: number },
  ep: { x: number; y: number },
  startSide: AnchorSide | null,
  endSide: AnchorSide | null
): ReadonlyArray<{ x: number; y: number }> {
  const dx = ep.x - sp.x;
  const dy = ep.y - sp.y;
  const LEAD = 228_600;
  const leadX = Math.min(LEAD, Math.max(0, Math.abs(dx) / 2));
  const leadY = Math.min(LEAD, Math.max(0, Math.abs(dy) / 2));
  const sV = sideUnit(startSide);
  const eV = sideUnit(endSide);
  if (sV.x === 0 && sV.y === 0 && eV.x === 0 && eV.y === 0) {
    const horizontalFirst = Math.abs(dx) >= Math.abs(dy);
    const midX = horizontalFirst ? sp.x + dx / 2 : sp.x;
    const midY = horizontalFirst ? sp.y : sp.y + dy / 2;
    return horizontalFirst
      ? [sp, { x: midX, y: sp.y }, { x: midX, y: ep.y }, ep]
      : [sp, { x: sp.x, y: midY }, { x: ep.x, y: midY }, ep];
  }
  const p1 = { x: sp.x + sV.x * leadX, y: sp.y + sV.y * leadY };
  const p2 = { x: ep.x + eV.x * leadX, y: ep.y + eV.y * leadY };
  const sIsHoriz = sV.x !== 0;
  const eIsHoriz = eV.x !== 0;
  if (sIsHoriz && eIsHoriz) {
    const midX = (p1.x + p2.x) / 2;
    return [sp, p1, { x: midX, y: p1.y }, { x: midX, y: p2.y }, p2, ep];
  }
  if (!sIsHoriz && !eIsHoriz) {
    const midY = (p1.y + p2.y) / 2;
    return [sp, p1, { x: p1.x, y: midY }, { x: p2.x, y: midY }, p2, ep];
  }
  const corner = sIsHoriz ? { x: p2.x, y: p1.y } : { x: p1.x, y: p2.y };
  return [sp, p1, corner, p2, ep];
}

function sideUnit(side: AnchorSide | null): { x: number; y: number } {
  switch (side) {
    case "n":
      return { x: 0, y: -1 };
    case "s":
      return { x: 0, y: 1 };
    case "e":
      return { x: 1, y: 0 };
    case "w":
      return { x: -1, y: 0 };
    case "center":
    case null:
      return { x: 0, y: 0 };
  }
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

interface CommentMarkerOverlayProps {
  readonly slide: Slide;
  readonly slideSize: SlideSize;
  readonly shapeIds: ReadonlyArray<string>;
}

/**
 * SVG overlay that paints a soft yellow outline + corner badge over
 * every shape that carries an unresolved comment. Drawn behind the
 * selection chrome so a selected commented shape still reads as
 * "selected"; the yellow border bleeds out from underneath.
 */
function CommentMarkerOverlay({
  slide,
  slideSize,
  shapeIds,
}: CommentMarkerOverlayProps): React.ReactElement | null {
  const boxes: { id: string; box: BoundingBox }[] = [];
  for (const id of shapeIds) {
    const sh = findShape(slide.shapes, id);
    if (!sh) continue;
    const box = shapeBoundingBox(sh);
    if (!box) continue;
    boxes.push({ id, box });
  }
  if (boxes.length === 0) return null;
  const badgeR = 8;
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      viewBox={slideViewBox(slideSize)}
      preserveAspectRatio="xMidYMid meet"
      pointerEvents="none"
      style={{ position: "absolute", inset: 0, width: "100%", height: "100%" }}
    >
      {boxes.map(({ id, box }) => {
        const x = px(box.x);
        const y = px(box.y);
        const w = px(box.cx);
        const h = px(box.cy);
        return (
          <g key={`comment-marker-${id}`}>
            <rect
              x={x}
              y={y}
              width={w}
              height={h}
              fill="rgba(250, 204, 21, 0.12)"
              stroke="#f59e0b"
              strokeWidth={2}
              strokeOpacity={0.85}
              vectorEffect="non-scaling-stroke"
              pointerEvents="none"
            />
            <circle
              cx={x + w}
              cy={y}
              r={badgeR}
              fill="#f59e0b"
              stroke="white"
              strokeWidth={1}
              vectorEffect="non-scaling-stroke"
              pointerEvents="none"
            />
          </g>
        );
      })}
    </svg>
  );
}

interface CommentFlashOverlayProps {
  readonly slide: Slide;
  readonly slideSize: SlideSize;
  readonly target:
    | { readonly kind: "shape"; readonly shapeId: string; readonly nonce: number }
    | { readonly kind: "pin"; readonly xEmu: number; readonly yEmu: number; readonly nonce: number };
}

/**
 * One-shot yellow flash painted over a shape (or around a pin
 * coordinate) so the user can spot the comment they just clicked in
 * the sidebar. Uses inline SVG SMIL animations so the file is
 * self-contained and doesn't depend on any host-app stylesheet.
 */
function CommentFlashOverlay({
  slide,
  slideSize,
  target,
}: CommentFlashOverlayProps): React.ReactElement | null {
  let box: BoundingBox | null = null;
  if (target.kind === "shape") {
    const sh = findShape(slide.shapes, target.shapeId);
    if (sh) box = shapeBoundingBox(sh);
  } else {
    // Pin-anchored flash: draw a fixed ~1 inch square around the pin
    // so there's something visible even when no shape is associated.
    const halfEmu = 457200; // 0.5"
    box = {
      x: Math.max(0, target.xEmu - halfEmu),
      y: Math.max(0, target.yEmu - halfEmu),
      cx: halfEmu * 2,
      cy: halfEmu * 2,
    };
  }
  if (!box) return null;
  const x = px(box.x);
  const y = px(box.y);
  const w = px(box.cx);
  const h = px(box.cy);
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      viewBox={slideViewBox(slideSize)}
      preserveAspectRatio="xMidYMid meet"
      pointerEvents="none"
      style={{ position: "absolute", inset: 0, width: "100%", height: "100%" }}
    >
      <rect
        x={x}
        y={y}
        width={w}
        height={h}
        fill="#facc15"
        fillOpacity={0.55}
        stroke="#d97706"
        strokeWidth={3}
        vectorEffect="non-scaling-stroke"
        pointerEvents="none"
      >
        <animate attributeName="fill-opacity" from="0.55" to="0" dur="1.4s" fill="freeze" />
        <animate attributeName="stroke-opacity" from="1" to="0" dur="1.4s" fill="freeze" />
      </rect>
    </svg>
  );
}

function renderEditingOverlay(
  slide: Slide,
  shapeId: string,
  slideSize: SlideSize,
  dpi: number,
  onCommit: (
    shape: TextShape,
    paragraphs: ReadonlyArray<{
      readonly runs: ReadonlyArray<{
        readonly text: string;
        readonly isLineBreak?: boolean;
        readonly inheritFromRun?: number;
      }>;
    }>,
    plain: string
  ) => void,
  onTextSelectionChange: ((sel: PptxTextSelection | null) => void) | undefined
): React.ReactElement | null {
  const shape = findShape(slide.shapes, shapeId);
  if (!shape || shape.kind !== "text") return null;
  const box = shapeBoundingBox(shape);
  if (!box) return null;
  return (
    <TextEditOverlay
      key={shapeId}
      shape={shape as TextShape}
      box={box}
      slideSize={slideSize}
      dpi={dpi}
      onCommit={(paragraphs, plain) => onCommit(shape as TextShape, paragraphs, plain)}
      onSelectionChange={onTextSelectionChange}
    />
  );
}

interface TextEditOverlayProps {
  readonly shape: TextShape;
  readonly box: BoundingBox;
  readonly slideSize: SlideSize;
  readonly dpi: number;
  readonly onCommit: (
    paragraphs: ReadonlyArray<{
      readonly runs: ReadonlyArray<{
        readonly text: string;
        readonly isLineBreak?: boolean;
        readonly inheritFromRun?: number;
      }>;
    }>,
    plain: string
  ) => void;
  readonly onSelectionChange?: (sel: PptxTextSelection | null) => void;
}

/**
 * Selection-aware contenteditable overlay.
 *
 * Renders each paragraph as its own `<div data-paragraph={i}>` and
 * each run as a `<span data-run={j}>` so we can preserve inline
 * formatting visually AND map a DOM Selection back to a
 * `PptxTextSelection` (paragraph + start/end character offsets) for
 * the shared text-format toolbar.
 *
 * Text edits still commit through `pptx:set-text` on blur (lossy:
 * collapses formatting). Format toolbar interactions are expected to
 * `e.preventDefault()` mousedown so this overlay keeps focus and the
 * selection ref stays valid while a `pptx:format-text` patch is
 * dispatched.
 */
function TextEditOverlay({
  shape,
  box,
  slideSize,
  dpi,
  onCommit,
  onSelectionChange,
}: TextEditOverlayProps): React.ReactElement {
  const ref = React.useRef<HTMLDivElement>(null);
  const initialPlain = React.useMemo(() => textShapePlain(shape), [shape]);

  // The SVG renders at 1 SVG-unit = 1 CSS px at 96 DPI, then the
  // browser scales the SVG to fit its container. Our HTML overlay
  // lives in the same container but receives raw CSS px, so we have
  // to multiply font sizes / insets by the same scale factor or the
  // overlay text appears noticeably larger than the underlying
  // rendered shape (and the user perceives it as "the font changed
  // when I clicked to edit").
  const [scale, setScale] = React.useState(1);
  const containerRef = React.useRef<HTMLDivElement | null>(null);
  React.useEffect(() => {
    const node = ref.current;
    if (!node) return;
    const parent = node.parentElement as HTMLDivElement | null;
    containerRef.current = parent;
    if (!parent || typeof ResizeObserver === "undefined") return;
    const slideNativeWidth = slideSize.cxEmu / EMU_PER_PX_AT_96DPI;
    const update = (): void => {
      const w = parent.getBoundingClientRect().width;
      if (w > 0 && slideNativeWidth > 0) setScale(w / slideNativeWidth);
    };
    update();
    const ro = new ResizeObserver(update);
    ro.observe(parent);
    return () => ro.disconnect();
  }, [slideSize.cxEmu]);

  React.useEffect(() => {
    const node = ref.current;
    if (!node) return;
    node.focus();
    const sel = window.getSelection?.();
    if (sel) {
      const range = document.createRange();
      range.selectNodeContents(node);
      range.collapse(false);
      sel.removeAllRanges();
      sel.addRange(range);
    }
  }, [shape.id]);

  // Mirror the live DOM selection into the shared callback so the
  // text-format toolbar can compute MIXED-state and dispatch
  // pptx:format-text against a real range.
  //
  // IMPORTANT — never report `null` here just because the native
  // selection happens to be empty/outside the editable. Clicking a
  // toolbar button (or opening a `<select>`) momentarily moves
  // focus and may collapse the native Selection; if we cleared the
  // ref the format dispatch would immediately see `selectionRef =
  // null` and bail with a "Select some text first." toast.
  // `null` is reserved for the explicit close path (commit/blur
  // outside the keep-edit zone), which calls `onSelectionChange?.
  // (null)` directly in the blur handler below.
  React.useEffect(() => {
    if (!onSelectionChange) return;
    const handler = () => {
      const node = ref.current;
      if (!node) return;
      const sel = window.getSelection?.();
      if (!sel || sel.rangeCount === 0) return;
      const range = sel.getRangeAt(0);
      if (!node.contains(range.startContainer) || !node.contains(range.endContainer)) {
        return;
      }
      const a = domPointToCharOffset(node, range.startContainer, range.startOffset);
      const b = domPointToCharOffset(node, range.endContainer, range.endOffset);
      if (!a || !b || a.paragraph !== b.paragraph) {
        // Cross-paragraph selections aren't supported by the toolbar
        // (pptx:format-text is per-paragraph) — leave the previously
        // captured single-paragraph selection in place.
        return;
      }
      const start = Math.min(a.offset, b.offset);
      const end = Math.max(a.offset, b.offset);
      onSelectionChange({ shapeId: shape.id, paragraph: a.paragraph, start, end });
    };
    document.addEventListener("selectionchange", handler);
    handler();
    return () => document.removeEventListener("selectionchange", handler);
  }, [onSelectionChange, shape.id]);

  const leftPct = (box.x / slideSize.cxEmu) * 100;
  const topPct = (box.y / slideSize.cyEmu) * 100;
  const widthPct = (box.cx / slideSize.cxEmu) * 100;
  const heightPct = (box.cy / slideSize.cyEmu) * 100;
  const insetsEmu = readBodyInsetsFromShape(shape);
  const anchor = readBodyAnchorFromShape(shape);
  const baseFontPx = (estimateFontPtFromShape(shape) * dpi) / 72;
  // Pick the same default font the SVG renderer uses so the text
  // doesn't visibly switch to the system sans-serif when the caret
  // enters the shape. We prefer an explicit run-level font (the most
  // accurate signal), then fall back to PowerPoint's default body
  // typeface (Calibri) — only finally to "sans-serif" if a browser
  // doesn't have Calibri installed.
  const baseFontFamily = estimateFontFamilyFromShape(shape);
  const justifyContent = anchor === "ctr" ? "center" : anchor === "b" ? "flex-end" : "flex-start";
  const padTop = (insetsEmu.t / EMU_PER_PX_AT_96DPI) * scale;
  const padRight = (insetsEmu.r / EMU_PER_PX_AT_96DPI) * scale;
  const padBottom = (insetsEmu.b / EMU_PER_PX_AT_96DPI) * scale;
  const padLeft = (insetsEmu.l / EMU_PER_PX_AT_96DPI) * scale;

  return (
    <div
      ref={ref}
      contentEditable
      suppressContentEditableWarning
      data-testid="pptx-text-overlay"
      style={{
        position: "absolute",
        left: `${leftPct}%`,
        top: `${topPct}%`,
        width: `${widthPct}%`,
        height: `${heightPct}%`,
        boxSizing: "border-box",
        // Match the visual idiom of the dashed selection rectangle
        // shown in shape-select mode — never the heavy solid frame
        // that made the overlay read as a chunky textarea instead of
        // a still-selected shape.
        outline: "1.5px dashed #7c3aed",
        outlineOffset: 0,
        background: "transparent",
        caretColor: "#7c3aed",
        paddingTop: padTop,
        paddingRight: padRight,
        paddingBottom: padBottom,
        paddingLeft: padLeft,
        fontSize: baseFontPx * scale,
        fontFamily: baseFontFamily,
        lineHeight: 1.2,
        whiteSpace: "pre-wrap",
        wordWrap: "break-word",
        overflow: "hidden",
        display: "flex",
        flexDirection: "column",
        justifyContent,
      }}
      onBlur={(e) => {
        // Don't commit when focus moves to a sibling we explicitly
        // marked as "keep editing focus" (the format toolbar). The
        // toolbar suppresses mousedown so the relatedTarget stays
        // null in most browsers; we still bail out if the user
        // clicked a button that opted in via data-pptx-keep-edit.
        const next = e.relatedTarget as HTMLElement | null;
        if (next?.closest?.("[data-pptx-keep-edit]")) return;
        const node = ref.current;
        const paragraphs = node ? extractParagraphsFromOverlay(node) : [];
        const plain = node?.innerText ?? "";
        onCommit(paragraphs, plain);
        onSelectionChange?.(null);
      }}
      onKeyDown={(e) => {
        if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) {
          e.preventDefault();
          ref.current?.blur();
        } else if (e.key === "Escape") {
          e.preventDefault();
          if (ref.current) ref.current.innerText = initialPlain;
          ref.current?.blur();
        }
      }}
    >
      {shape.txBody.paragraphs.map((p, pi) => (
        <div key={pi} data-paragraph={pi} style={paragraphStyle(p)}>
          {paragraphToReact(p, pi, scale, dpi)}
        </div>
      ))}
    </div>
  );
}

/**
 * Walk the contenteditable DOM and produce a structured paragraph
 * patch list that mirrors the original `<div data-paragraph>` /
 * `<span data-run>` markup. Each text slice carries an
 * `inheritFromRun` index so the `pptx:set-text` handler can copy the
 * original run's full property bag (incl. opaque XML) and avoid
 * collapsing bold/italic/colour spans down to a single run.
 *
 * Runs the user typed *outside* any existing span (e.g. fresh text
 * appended after a styled span) report `inheritFromRun: undefined`
 * and the handler falls back to the paragraph's first run's
 * properties — matches PowerPoint's "new text inherits from the
 * insertion point's run" behaviour.
 */
function extractParagraphsFromOverlay(root: HTMLElement): ReadonlyArray<{
  readonly runs: ReadonlyArray<{
    readonly text: string;
    readonly isLineBreak?: boolean;
    readonly inheritFromRun?: number;
  }>;
}> {
  const out: Array<{
    runs: Array<{ text: string; isLineBreak?: boolean; inheritFromRun?: number }>;
  }> = [];
  const paragraphNodes = Array.from(root.querySelectorAll<HTMLElement>("[data-paragraph]"));
  // If the user wiped everything (browsers can leave a bare <br> or
  // empty editable), fall back to a single empty paragraph so we at
  // least clear the shape rather than throwing.
  if (paragraphNodes.length === 0) {
    return [{ runs: [{ text: root.innerText.split("\n")[0] ?? "" }] }];
  }
  for (const pNode of paragraphNodes) {
    const runs: Array<{ text: string; isLineBreak?: boolean; inheritFromRun?: number }> = [];
    const walker = document.createTreeWalker(pNode, NodeFilter.SHOW_ALL);
    let cur: Node | null = walker.nextNode();
    while (cur) {
      if (cur.nodeType === Node.TEXT_NODE) {
        const text = (cur.nodeValue ?? "").replace(/\u00a0/g, " ");
        if (text.length > 0) {
          const span = (cur.parentElement?.closest("[data-run]") as HTMLElement | null) ?? null;
          const idxAttr = span?.getAttribute("data-run");
          const inheritFromRun =
            idxAttr !== null && idxAttr !== undefined ? Number.parseInt(idxAttr, 10) : undefined;
          runs.push({
            text,
            ...(inheritFromRun !== undefined && Number.isFinite(inheritFromRun) ? { inheritFromRun } : {}),
          });
        }
      } else if (cur.nodeType === Node.ELEMENT_NODE) {
        const el = cur as HTMLElement;
        if (el.tagName === "BR" && !el.hasAttribute("data-paragraph-eol")) {
          // Soft line break inside a paragraph (Shift+Enter) — modeled
          // as an explicit isLineBreak run so the serializer can emit
          // <a:br/> and round-trip cleanly.
          const span = el.closest("[data-run]") as HTMLElement | null;
          const idxAttr = span?.getAttribute("data-run");
          const inheritFromRun =
            idxAttr !== null && idxAttr !== undefined ? Number.parseInt(idxAttr, 10) : undefined;
          runs.push({
            text: "",
            isLineBreak: true,
            ...(inheritFromRun !== undefined && Number.isFinite(inheritFromRun) ? { inheritFromRun } : {}),
          });
        }
      }
      cur = walker.nextNode();
    }
    if (runs.length === 0) runs.push({ text: "" });
    out.push({ runs });
  }
  return out;
}

function paragraphStyle(p: TextShape["txBody"]["paragraphs"][number]): React.CSSProperties {
  const align =
    p.properties.alignment === "center"
      ? "center"
      : p.properties.alignment === "right"
        ? "right"
        : p.properties.alignment === "justify"
          ? "justify"
          : "left";
  return { textAlign: align };
}

function paragraphToReact(
  p: TextShape["txBody"]["paragraphs"][number],
  paragraphIndex: number,
  scale: number,
  dpi: number
): React.ReactNode {
  const flatLen = p.runs.reduce((acc, r) => acc + (r.isLineBreak ? 0 : r.text.length), 0);
  if (flatLen === 0) {
    // Empty paragraphs need a <br> so contenteditable gives them a
    // visible row + a stable caret target.
    return <br data-paragraph-eol={paragraphIndex} />;
  }
  return p.runs.map((r, ri) => {
    if (r.isLineBreak) return <br key={ri} data-run={ri} />;
    return (
      <span key={ri} data-run={ri} style={runStyle(r.properties, scale, dpi)}>
        {r.text}
      </span>
    );
  });
}

function runStyle(
  props: TextShape["txBody"]["paragraphs"][number]["runs"][number]["properties"],
  scale: number,
  dpi: number
): React.CSSProperties {
  const out: React.CSSProperties = {};
  if (props.bold) out.fontWeight = 700;
  if (props.italic) out.fontStyle = "italic";
  const decorations: string[] = [];
  if (props.underline) decorations.push("underline");
  if (props.strike) decorations.push("line-through");
  if (decorations.length > 0) out.textDecoration = decorations.join(" ");
  if (props.color) out.color = `#${props.color}`;
  if (props.highlight) out.backgroundColor = `#${props.highlight}`;
  if (props.fontFamily) out.fontFamily = props.fontFamily;
  if (props.fontSizeHundredths !== undefined) {
    const pt = props.fontSizeHundredths / 100;
    out.fontSize = `${(pt * dpi * scale) / 72}px`;
  }
  return out;
}

function readBodyAnchorFromShape(shape: TextShape): "t" | "ctr" | "b" {
  const bodyPr = shape.txBody.bodyPrRaw;
  const v = bodyPr?.attrs?.anchor ?? bodyPr?.rawAttrs?.["@_anchor"];
  if (v === "ctr") return "ctr";
  if (v === "b") return "b";
  return "t";
}

function readBodyInsetsFromShape(shape: TextShape): { l: number; r: number; t: number; b: number } {
  const bodyPr = shape.txBody.bodyPrRaw;
  const get = (key: string, dflt: number): number => {
    const raw = bodyPr?.attrs?.[key] ?? bodyPr?.rawAttrs?.[`@_${key}`];
    if (typeof raw !== "string") return dflt;
    const n = Number(raw);
    return Number.isFinite(n) && n >= 0 ? n : dflt;
  };
  return {
    l: get("lIns", 91440),
    r: get("rIns", 91440),
    t: get("tIns", 45720),
    b: get("bIns", 45720),
  };
}

function estimateFontPtFromShape(shape: TextShape): number {
  const r = shape.txBody.paragraphs[0]?.runs.find((x) => !x.isLineBreak);
  if (r?.properties.fontSizeHundredths !== undefined) {
    return r.properties.fontSizeHundredths / 100;
  }
  return 18;
}

function estimateFontFamilyFromShape(shape: TextShape): string {
  for (const para of shape.txBody.paragraphs) {
    for (const r of para.runs) {
      if (r.isLineBreak) continue;
      if (r.properties.fontFamily) return `${r.properties.fontFamily}, sans-serif`;
    }
  }
  return "Calibri, sans-serif";
}

function domPointToCharOffset(
  root: HTMLElement,
  node: Node,
  offset: number
): { paragraph: number; offset: number } | null {
  // Find the paragraph element containing `node` (or root, when
  // selection is on the root itself).
  let paraEl: HTMLElement | null = null;
  let cursor: Node | null = node;
  while (cursor && cursor !== root) {
    if (cursor.nodeType === 1 && (cursor as HTMLElement).hasAttribute("data-paragraph")) {
      paraEl = cursor as HTMLElement;
      break;
    }
    cursor = cursor.parentNode;
  }
  if (!paraEl) {
    // Selection landed on the root or between paragraphs. Map the
    // root-level child index to the nearest paragraph index.
    if (node === root) {
      const child = root.children[Math.min(offset, root.children.length - 1)] as HTMLElement | undefined;
      if (child?.hasAttribute("data-paragraph")) {
        return { paragraph: Number(child.dataset.paragraph), offset: 0 };
      }
    }
    return null;
  }
  const paragraph = Number(paraEl.dataset.paragraph);

  // Walk all text nodes inside paraEl in document order and sum char
  // counts up to (node, offset).
  const walker = document.createTreeWalker(paraEl, NodeFilter.SHOW_TEXT, null);
  let total = 0;
  let n = walker.nextNode();
  while (n) {
    if (n === node) {
      total += offset;
      return { paragraph, offset: total };
    }
    total += (n.textContent ?? "").length;
    n = walker.nextNode();
  }
  // The caller passed an element node — figure out how many chars
  // precede it within the paragraph.
  if (node.nodeType === 1) {
    const el = node as HTMLElement;
    let pre = 0;
    const w2 = document.createTreeWalker(paraEl, NodeFilter.SHOW_TEXT, null);
    let m = w2.nextNode();
    while (m) {
      if (el.contains(m)) break;
      pre += (m.textContent ?? "").length;
      m = w2.nextNode();
    }
    return { paragraph, offset: pre };
  }
  return { paragraph, offset: total };
}

function textShapePlain(shape: TextShape): string {
  return shape.txBody.paragraphs
    .map((p) => p.runs.map((r) => (r.isLineBreak ? "\n" : r.text)).join(""))
    .join("\n");
}
