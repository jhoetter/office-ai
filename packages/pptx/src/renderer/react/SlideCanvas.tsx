import * as React from "react";
import type { PptxAgent } from "../../agent/agent.js";
import type {
  ConnectorShape,
  ConnectorType,
  Shape,
  Slide,
  SlideLayout,
  SlideSize,
  TextShape,
} from "../../model/types.js";
import { resolveEndpoint } from "../../model/connector-geometry.js";
import { DEFAULT_THEME } from "../layout/color.js";
import { boxesIntersect, shapeBoundingBox, type BoundingBox } from "../layout/shape.js";
import { resolveRotatedResize } from "../layout/resize.js";
import { computeStageLayout, slideStageViewBox } from "../layout/slide.js";
import { resolvePlaceholderTextDefaults, resolvedShapeBoundingBox } from "../layout/placeholder-defaults.js";
import { computeSnap, type SnapGuide } from "../layout/snap.js";
import { anchorsFor, snapToAnchor, type AnchorSide, type ShapeAnchor } from "../layout/anchors.js";
import { DEFAULT_DPI, EMU_PER_PX_AT_96DPI, clampZoom } from "../layout/units.js";
import type { SvgRenderCtx } from "../svg/shapes.js";
import { shapeToSvg } from "../svg/shapes.js";
import { buildShapesByCNvPrId, resolveSlideBackgroundColor } from "../svg/slide.js";
import {
  collectObstacles,
  routeConnector as routeConnectorShared,
  type RouterObstacle,
} from "../connector-router/index.js";
import { useAgentSnapshot } from "./use-agent-snapshot.js";

/**
 * Dynamic stage viewBox for the editor canvas.
 *
 * The outer stage div fills the entire visible slide section so that
 * pointer events anywhere in the user's viewport (clicks → deselect,
 * drags → marquee) are live, regardless of where the slide card sits
 * inside it. This is what makes the "scratch canvas" feel like a real
 * artboard instead of a small sticker pasted on the page-backdrop.
 *
 * The slide card is positioned with absolute pixels (computed from a
 * ResizeObserver on the stage div, see {@link useStageLayout}). We
 * then publish a viewBox whose user units pin (0..slideW, 0..slideH)
 * to the slide card's pixel rect. With the SVG's
 * `preserveAspectRatio="xMidYMid meet"` and a viewBox aspect ratio
 * that exactly matches the stage div's, no letterboxing happens and
 * every overlay's `(0, 0) = top-left of slide` invariant survives.
 *
 * `null` means "fall back to {@link slideStageViewBox}" — used by
 * tests / Storybook that mount overlays in isolation without a parent
 * stage div, so this Context never breaks them.
 */
const StageViewBoxContext = React.createContext<string | null>(null);

/**
 * Read the active stage viewBox, falling back to the static
 * `slideStageViewBox(slideSize)` when no provider is mounted (tests,
 * isolated previews). Lets every overlay paint with the correct
 * coordinate frame without prop-drilling a `stageViewBox` prop
 * through a dozen function components.
 */
function useStageViewBox(slideSize: SlideSize): string {
  const ctx = React.useContext(StageViewBoxContext);
  return ctx ?? slideStageViewBox(slideSize);
}

/**
 * Pixel rect of the slide card within the outer stage div, plus the
 * dynamic viewBox that pins SVG user-units to that rect. Derived from
 * a ResizeObserver on the stage div so we recompute exactly when
 * needed (mount, parent resize, sidebar toggle, zoom change, …).
 */
interface StageLayout {
  readonly stageW: number;
  readonly stageH: number;
  readonly slidePxLeft: number;
  readonly slidePxTop: number;
  readonly slidePxW: number;
  readonly slidePxH: number;
  readonly stageViewBox: string;
}

/**
 * Measure the stage div with a ResizeObserver and derive the slide
 * card's pixel rect + a matching SVG viewBox.
 *
 * The slide card preserves the slide aspect ratio and fits inside the
 * stage with `min(stageW / slideAspect, stageH) * zoom` so changing
 * the zoom multiplier scales the slide visually without distorting
 * it. The viewBox is sized so 1 user unit per slide pixel, with a
 * negative origin equal to the slide card's pixel offset — that way
 * coordinate (0, 0) in the SVG remains the top-left of the slide and
 * shapes positioned in the scratch margin (negative EMU) sit in the
 * surrounding gray area exactly where the user dropped them.
 */
function useStageLayout(
  slideSize: SlideSize,
  zoom: number
): {
  readonly containerRef: React.RefObject<HTMLDivElement | null>;
  readonly layout: StageLayout | null;
} {
  const containerRef = React.useRef<HTMLDivElement>(null);
  const [size, setSize] = React.useState<{ w: number; h: number } | null>(null);
  React.useLayoutEffect(() => {
    const el = containerRef.current;
    if (!el) return;
    const update = () => {
      const r = el.getBoundingClientRect();
      setSize({ w: r.width, h: r.height });
    };
    update();
    if (typeof ResizeObserver === "undefined") return;
    const ro = new ResizeObserver(update);
    ro.observe(el);
    return () => ro.disconnect();
  }, []);
  const layout = React.useMemo<StageLayout | null>(() => {
    if (!size) return null;
    const dyn = computeStageLayout(slideSize, size.w, size.h, zoom);
    if (!dyn) return null;
    return {
      stageW: size.w,
      stageH: size.h,
      slidePxLeft: dyn.slidePxLeft,
      slidePxTop: dyn.slidePxTop,
      slidePxW: dyn.slidePxW,
      slidePxH: dyn.slidePxH,
      stageViewBox: dyn.stageViewBox,
    };
  }, [size, slideSize, zoom]);
  return { containerRef, layout };
}

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
  /**
   * Fired when the user "activates" an empty layout placeholder — e.g.
   * double-clicks the dashed ghost box for a `pic` placeholder. The
   * parent decides what to do (typically: open a file picker for `pic`,
   * a chart wizard for `chart`, etc.) and dispatches the appropriate
   * commands. Returning without doing anything is fine; the canvas
   * will simply keep the placeholder selected.
   *
   * Activation deliberately does NOT fall through to text-edit mode
   * for placeholder types where typing isn't the natural action
   * (`pic`, `chart`, `tbl`, `dgm`, `media`). Title / body / subTitle
   * placeholders still open the text overlay on activate so the
   * "click to add text" hint remains directly editable.
   */
  readonly onPlaceholderActivate?: (info: {
    readonly shapeId: string;
    readonly placeholder: { readonly type?: string; readonly idx?: number };
  }) => void;
  /**
   * Fired when the user double-clicks a shape kind that is not text-
   * editable but does have an "open in editor" affordance — currently
   * `chart` and `ole-spreadsheet`. The host editor is expected to open
   * an inline edit-data modal (see `EmbeddedXlsxModal`) and dispatch
   * the appropriate `*:update-spreadsheet` / `*:set-chart-data`
   * commands when the user finishes.
   *
   * Returning without doing anything is a valid no-op; the canvas
   * leaves the shape selected and does not fall through to text-edit
   * mode.
   */
  readonly onShapeActivate?: (info: {
    readonly shapeId: string;
    readonly shape: Shape;
  }) => void;
  /**
   * When non-null, the canvas is in "draw a connector" tool mode: any
   * shape under the pointer surfaces ports (regardless of selection),
   * the cursor becomes a crosshair, and a press-drag gesture starts a
   * draft of the requested type — even from empty space. Set by the
   * editor when the user picks a connector type from the toolbar or
   * command palette; cleared when the draft commits, on Esc, or when
   * the user re-clicks the same toolbar item.
   */
  readonly connectorTool?: { readonly type: ConnectorType } | null;
  /**
   * Fired when the canvas exits tool mode on its own (Esc, draft
   * commit). The editor mirrors the state so the toolbar's pressed
   * indicator stays in sync.
   */
  readonly onConnectorToolExit?: () => void;
  /**
   * Optional remote-peer selection presence (Yjs awareness). For each
   * peer whose `slideId` matches the active slide, the canvas paints
   * a colored 2px outline around the shape(s) they have selected,
   * with a small name tag in the peer color. When a peer is on the
   * slide but has no shape selected (`shapeIds: []`), the canvas
   * paints a small "X is here" badge in the top-right corner instead.
   * Empty / undefined → nothing is drawn.
   */
  readonly remotePeers?: ReadonlyArray<RemoteSelectionPeer>;
}

export interface RemoteSelectionPeer {
  readonly clientId: number;
  readonly slideId: string;
  readonly shapeIds: ReadonlyArray<string>;
  readonly name: string;
  readonly color: string;
}

type ResizeHandle = "n" | "s" | "e" | "w" | "nw" | "ne" | "sw" | "se";
type DragMode = "move" | { resize: ResizeHandle };

interface DragState {
  readonly mode: DragMode;
  /**
   * Shapes participating in the drag with their pre-drag bounding
   * boxes and rotation. `originDeg` is captured at pointerdown so
   * the rotation-aware resize math can run without a fresh snapshot
   * lookup mid-gesture (and to keep the math stable if the user's
   * own commands somehow re-enter while the drag is live).
   */
  readonly targets: ReadonlyArray<{
    readonly id: string;
    readonly origin: BoundingBox;
    readonly originDeg: number;
  }>;
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
 * Live state while the user is drawing a brand-new connector. The
 * source can be anchored (drag started on a port) or free (drag
 * started on empty space while in tool mode). The destination tracks
 * the pointer and becomes anchored when it snaps onto another shape's
 * port, free otherwise.
 *
 * Scoped purely to the canvas — committed to the agent on pointerup
 * via `pptx:add-connector`, then reset.
 */
interface ConnectorDraft {
  readonly source:
    | {
        readonly kind: "anchored";
        readonly shapeId: string;
        readonly cNvPrId: number;
        readonly side: "n" | "s" | "e" | "w" | "center";
        readonly t: number;
        readonly x: number;
        readonly y: number;
      }
    | {
        readonly kind: "free";
        readonly x: number;
        readonly y: number;
      };
  /** Connector type to commit; kept in state so Esc/replace mid-draft works. */
  readonly connectorType: ConnectorType;
  /** Live cursor position in slide EMU coordinates. */
  readonly currentX: number;
  readonly currentY: number;
  readonly emuPerPx: number;
  /** Anchor we'd snap the destination endpoint onto if released now. */
  readonly snapped: ShapeAnchor | null;
  readonly nearby: ReadonlyArray<ShapeAnchor>;
}

/**
 * Live state while the user is dragging one endpoint of an
 * already-existing connector. The endpoint behaves like a brand-new
 * draft destination: snap, optionally re-anchor on release.
 */
interface EndpointEditDraft {
  readonly shapeId: string;
  readonly which: "start" | "end";
  /** Whether this endpoint started the gesture as anchored. */
  readonly wasAnchored: boolean;
  readonly currentX: number;
  readonly currentY: number;
  readonly emuPerPx: number;
  readonly snapped: ShapeAnchor | null;
  readonly nearby: ReadonlyArray<ShapeAnchor>;
  /**
   * The other (un-dragged) endpoint, resolved to slide coords once at
   * gesture start so the live overlay can re-route an elbow preview.
   */
  readonly otherPoint: { readonly x: number; readonly y: number };
  readonly otherSide: AnchorSide | null;
  /**
   * The `cNvPrId` of the shape the OTHER (un-dragged) endpoint is
   * anchored to, or null when that endpoint is free. Used by the live
   * preview overlay to exclude that shape from the obstacle list so
   * the route doesn't try to detour around its own anchor target.
   */
  readonly otherEndpointCNvPrId: number | null;
  readonly connectorType: ConnectorType;
}

/**
 * Live state while the user is dragging a perpendicular slider on an
 * elbow connector's interior segment.
 */
interface WaypointDraft {
  readonly shapeId: string;
  readonly segmentIndex: number;
  /** "horizontal" segments slide vertically, "vertical" segments slide horizontally. */
  readonly axis: "horizontal" | "vertical";
  readonly originValueEmu: number;
  readonly startClientX: number;
  readonly startClientY: number;
  readonly emuPerPx: number;
  /** Live offset value (originValue + cursor delta), held for preview. */
  readonly currentValueEmu: number;
}

interface RotateDraftTarget {
  readonly id: string;
  /** Pre-rotation (axis-aligned) bbox in EMU; needed to draw the ghost
   * at the original position with the new rotation applied. */
  readonly origin: BoundingBox;
  /** Rotation the shape carried before the gesture started. */
  readonly originDeg: number;
  /** Captured shape so the ghost layer can re-emit it cheaply. */
  readonly shape: Shape;
}

/**
 * Live state while the user drags the rotation grip. We capture the
 * pre-rotation rotation per shape on pointerdown and only commit on
 * pointerup, so the gesture round-trips through one `pptx:set-rotation`
 * per selected shape — the snapshot stays untouched until release,
 * keeping undo a single step regardless of how much the user wiggled
 * during the gesture. The pivot is the union-bbox centre (also where
 * the cursor angle is measured from). Each shape rotates around its
 * OWN centre by the cursor's angular delta — matching what the
 * `pptx:set-rotation` handler does, and what users expect from a
 * "rotate the selection by N degrees" gesture (vs. PowerPoint's
 * "rotate the whole group around the union centre", which would
 * require also moving each shape's position).
 */
interface RotateDraft {
  readonly targets: ReadonlyArray<RotateDraftTarget>;
  /** Pivot for cursor-angle math, in slide EMU. */
  readonly pivotX: number;
  readonly pivotY: number;
  /** Pointer angle (radians, atan2) at gesture start. */
  readonly startAngleRad: number;
  readonly currentAngleRad: number;
  readonly emuPerPx: number;
  /** Holding shift snaps to 15° increments — same convention as Figma / Keynote. */
  readonly shiftSnap: boolean;
}

/** Shared formula so pointermove preview, ghost layer, and pointerup
 * commit all derive the same delta from the draft state. */
function rotateDraftDeltaDeg(draft: RotateDraft): number {
  const raw = ((draft.currentAngleRad - draft.startAngleRad) * 180) / Math.PI;
  return draft.shiftSnap ? Math.round(raw / 15) * 15 : raw;
}

export function SlideCanvas(props: SlideCanvasProps): React.ReactElement | null {
  const snap = useAgentSnapshot(props.agent);
  const slide: Slide | undefined = snap.root.slides[props.slideIndex];
  const slideSize: SlideSize = snap.root.slideSize;
  const zoomEarly = clampZoom(props.zoom ?? 1);
  const { containerRef, layout: stageLayout } = useStageLayout(slideSize, zoomEarly);
  /**
   * Bounding rect of the inner white slide card, used to convert client
   * coordinates to slide-EMU. We measure the slide rect (not the outer
   * stage rect) so points inside the surrounding scratch area produce
   * the negative / over-slide EMU values that off-slide shapes live in.
   */
  const slideRectRef = React.useRef<HTMLDivElement>(null);
  const [drag, setDrag] = React.useState<DragState | null>(null);
  const [preview, setPreview] = React.useState<DragPreview | null>(null);
  const [marquee, setMarquee] = React.useState<MarqueeState | null>(null);
  const [editingId, setEditingId] = React.useState<string | null>(null);
  /**
   * Callable handle the active `TextEditOverlay` registers on mount,
   * so the canvas can commit + tear down the overlay synchronously
   * from `onPointerDown` BEFORE mutating `selectedIds`. Without this
   * the user briefly sees "B selected + A still in edit chrome"
   * because the editable's `onBlur` only fires after the click event
   * is fully dispatched (and thus after `setSelectedIds` has run).
   */
  const editCommitRef = React.useRef<(() => void) | null>(null);
  const [hoveredShapeId, setHoveredShapeId] = React.useState<string | null>(null);
  const [connectorDraft, setConnectorDraft] = React.useState<ConnectorDraft | null>(null);
  const [endpointDraft, setEndpointDraft] = React.useState<EndpointEditDraft | null>(null);
  const [waypointDraft, setWaypointDraft] = React.useState<WaypointDraft | null>(null);
  const [rotateDraft, setRotateDraft] = React.useState<RotateDraft | null>(null);
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
    // The rotate ghost layer paints rotated copies of every target;
    // hide the underlying shapes so we don't double-render the static
    // (pre-rotation) version underneath the live one.
    if (rotateDraft) for (const t of rotateDraft.targets) set.add(t.id);
    return set;
  }, [editingId, drag, rotateDraft]);

  // Per-shape rotation override published while the rotation grip is
  // being dragged. Consumed by the selection overlay so the dashed
  // outline + handles + grip rotate frame-by-frame in step with the
  // ghost layer; null at rest so the chrome falls back to each
  // shape's saved `rotation` (which is also what the static SVG
  // renderer uses, so chrome and shape stay locked together
  // post-commit).
  const liveRotations = React.useMemo(() => {
    if (!rotateDraft) return null;
    const delta = rotateDraftDeltaDeg(rotateDraft);
    const map = new Map<string, number>();
    for (const t of rotateDraft.targets) map.set(t.id, t.originDeg + delta);
    return map;
  }, [rotateDraft]);

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
      if (!slide || !containerRef.current || !slideRectRef.current) return;
      if (e.button !== 0) return;
      const target = e.target as Element | null;
      const portEl = target?.closest("[data-port]") as HTMLElement | null;
      const handleEl = target?.closest("[data-handle]") as HTMLElement | null;
      const shapeEl = target?.closest("[data-shape-id]") as SVGGElement | null;
      const shapeId =
        portEl?.dataset.portShapeId ?? handleEl?.dataset.shapeId ?? shapeEl?.dataset.shapeId ?? null;
      // If a text-edit overlay is open and the click is going to
      // change selection (different shape or empty workspace), commit
      // the edit synchronously NOW so the next render is consistent
      // (no half-frame "B selected + A still editing" flash). Clicks
      // inside the overlay itself are skipped — those should keep
      // editing.
      const insideOverlay = !!target?.closest("[data-testid='pptx-text-overlay']");
      if (editingId && !insideOverlay && editingId !== shapeId) {
        editCommitRef.current?.();
      }
      // Pointer math is anchored to the slide rect (not the surrounding
      // stage) so client → EMU conversion stays consistent regardless of
      // how much scratch padding the layout chooses, and clicks in the
      // scratch margin produce negative / over-slide EMU coordinates
      // that off-slide shapes naturally live in.
      const rect = slideRectRef.current.getBoundingClientRect();
      const emuPerPx = slideSize.cxEmu / rect.width;
      const shiftHeld = e.shiftKey || e.metaKey || e.ctrlKey;

      // Connector chrome handles (endpoint or waypoint) take precedence
      // over the regular shape gesture path. We pick them up before
      // anything else so a connector's selection chrome can sit ON TOP
      // of the shape rect without losing the click.
      const connectorEndpointEl = target?.closest("[data-connector-endpoint]") as HTMLElement | null;
      if (connectorEndpointEl) {
        const cShapeId = connectorEndpointEl.dataset.connectorShapeId ?? null;
        const which = connectorEndpointEl.dataset.connectorEndpoint as "start" | "end" | undefined;
        if (cShapeId && which) {
          const sh = findShape(slide.shapes, cShapeId);
          if (sh && sh.kind === "connector") {
            const ep = which === "start" ? sh.start : sh.end;
            const otherEp = which === "start" ? sh.end : sh.start;
            const cur = resolveEndpoint(ep, shapesByCNvPrId) ?? fallbackEndpoint(sh, which);
            const otherPt =
              resolveEndpoint(otherEp, shapesByCNvPrId) ??
              fallbackEndpoint(sh, which === "start" ? "end" : "start");
            const otherSide = otherEp.kind === "anchored" ? otherEp.side : null;
            const otherEndpointCNvPrId = otherEp.kind === "anchored" ? otherEp.targetCNvPrId : null;
            setEndpointDraft({
              shapeId: cShapeId,
              which,
              wasAnchored: ep.kind === "anchored",
              currentX: cur.x,
              currentY: cur.y,
              emuPerPx,
              snapped: null,
              nearby: [],
              otherPoint: otherPt,
              otherSide,
              otherEndpointCNvPrId,
              connectorType: sh.connectorType,
            });
            setSelectedIds([cShapeId]);
            containerRef.current.setPointerCapture?.(e.pointerId);
            e.preventDefault();
            e.stopPropagation();
            return;
          }
        }
      }
      const waypointEl = target?.closest("[data-connector-waypoint]") as HTMLElement | null;
      if (waypointEl) {
        const cShapeId = waypointEl.dataset.connectorShapeId ?? null;
        const segIdxStr = waypointEl.dataset.connectorWaypoint;
        const axis = waypointEl.dataset.connectorAxis as "horizontal" | "vertical" | undefined;
        const originStr = waypointEl.dataset.connectorOrigin;
        const segmentIndex = segIdxStr ? Number(segIdxStr) : NaN;
        const origin = originStr ? Number(originStr) : NaN;
        if (cShapeId && axis && Number.isFinite(segmentIndex) && Number.isFinite(origin)) {
          setWaypointDraft({
            shapeId: cShapeId,
            segmentIndex,
            axis,
            originValueEmu: origin,
            startClientX: e.clientX,
            startClientY: e.clientY,
            emuPerPx,
            currentValueEmu: origin,
          });
          setSelectedIds([cShapeId]);
          containerRef.current.setPointerCapture?.(e.pointerId);
          e.preventDefault();
          e.stopPropagation();
          return;
        }
      }

      // Port click: start a brand-new connector. This takes precedence
      // over the regular select/move/resize gesture so a hovered shape
      // surface stays draggable everywhere except the four port dots.
      if (portEl && shapeId) {
        const side = portEl.dataset.port as "n" | "s" | "e" | "w" | "center" | undefined;
        const located = findShapeWithOffset(slide.shapes, shapeId);
        if (side && located) {
          const sourceShape = located.shape;
          const localBox = shapeBoundingBox(sourceShape);
          const cNvPrId = sourceShape.cNvPrId;
          if (localBox && cNvPrId > 0) {
            const box: BoundingBox = {
              x: localBox.x + located.offsetX,
              y: localBox.y + located.offsetY,
              cx: localBox.cx,
              cy: localBox.cy,
            };
            const ap = anchorPointFor(box, side);
            const cursorEmuX = (e.clientX - rect.left) * emuPerPx;
            const cursorEmuY = (e.clientY - rect.top) * emuPerPx;
            setHoveredShapeId(null);
            setConnectorDraft({
              source: {
                kind: "anchored",
                shapeId,
                cNvPrId,
                side,
                t: 0.5,
                x: ap.x,
                y: ap.y,
              },
              connectorType: props.connectorTool?.type ?? "elbow",
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

      // Tool mode: a click anywhere on the canvas (port, shape body, or
      // empty space) starts a draft. Port clicks were handled above and
      // produce an anchored source; everything else produces a free
      // source so the user can draw "in mid-air" the way Google Slides
      // and PowerPoint allow.
      if (props.connectorTool) {
        const cursorEmuX = (e.clientX - rect.left) * emuPerPx;
        const cursorEmuY = (e.clientY - rect.top) * emuPerPx;
        setHoveredShapeId(null);
        setConnectorDraft({
          source: { kind: "free", x: cursorEmuX, y: cursorEmuY },
          connectorType: props.connectorTool.type,
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

      // Rotation grip — single dedicated handle above the union bbox.
      // We branch BEFORE the generic resize/move dispatch so the
      // gesture never gets confused with a corner drag, and we take
      // the existing selection verbatim (the grip only renders when
      // there's already a selection).
      if (handleEl?.dataset.handle === "rotate") {
        const ids = selectedIds.length > 0 ? selectedIds : [shapeId];
        const targets: RotateDraftTarget[] = [];
        let unionBox: BoundingBox | null = null;
        for (const id of ids) {
          const sh = findShape(slide.shapes, id);
          if (!sh) continue;
          // The same kinds the `pptx:set-rotation` handler refuses.
          // Rotating their wrapper would desync the model from what
          // PowerPoint reads back, so we exclude them from the gesture
          // entirely (no ghost, no commit).
          if (sh.kind === "connector" || sh.kind === "group" || sh.kind === "opaque") continue;
          const box = shapeBoundingBox(sh);
          if (!box) continue;
          const originDeg =
            "rotation" in sh && typeof sh.rotation === "number" ? sh.rotation : 0;
          targets.push({ id, origin: box, originDeg, shape: sh });
          if (!unionBox) {
            unionBox = box;
          } else {
            const x = Math.min(unionBox.x, box.x);
            const y = Math.min(unionBox.y, box.y);
            const right = Math.max(unionBox.x + unionBox.cx, box.x + box.cx);
            const bottom = Math.max(unionBox.y + unionBox.cy, box.y + box.cy);
            unionBox = { x, y, cx: right - x, cy: bottom - y };
          }
        }
        if (!unionBox || targets.length === 0) {
          e.preventDefault();
          return;
        }
        const pivotX = unionBox.x + unionBox.cx / 2;
        const pivotY = unionBox.y + unionBox.cy / 2;
        const cursorEmuX = (e.clientX - rect.left) * emuPerPx;
        const cursorEmuY = (e.clientY - rect.top) * emuPerPx;
        const startAngleRad = Math.atan2(cursorEmuY - pivotY, cursorEmuX - pivotX);
        setRotateDraft({
          targets,
          pivotX,
          pivotY,
          startAngleRad,
          currentAngleRad: startAngleRad,
          emuPerPx,
          shiftSnap: e.shiftKey,
        });
        containerRef.current.setPointerCapture?.(e.pointerId);
        e.preventDefault();
        e.stopPropagation();
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
      const targets: { id: string; origin: BoundingBox; originDeg: number }[] = [];
      for (const id of dragIds) {
        const sh = findShape(slide.shapes, id);
        if (!sh) continue;
        const box = shapeBoundingBox(sh);
        if (!box) continue;
        // Connectors no longer expose 8 generic resize handles —
        // their selection chrome paints two endpoint dots and zero or
        // more waypoint sliders, all routed through dedicated commands.
        // Body drags translate the whole connector by the cursor delta;
        // if either endpoint was anchored we detach it on commit (set
        // it to a free endpoint at the translated coords). This matches
        // PowerPoint/Slides where dragging a connected line's body
        // moves it as a unit and breaks the connections — the user can
        // then re-snap each endpoint by dragging the dots.
        if (sh.kind === "connector" && isResize) continue;
        const originDeg =
          "rotation" in sh && typeof sh.rotation === "number" ? sh.rotation : 0;
        targets.push({ id, origin: box, originDeg });
      }
      if (targets.length === 0) {
        // The selection update already ran; nothing else to do (e.g.
        // user clicked a connector body but both endpoints are
        // anchored, so no body drag is allowed).
        e.preventDefault();
        return;
      }

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
    [slide, slideSize, setSelectedIds, selectedIds, editingId, props.connectorTool]
  );

  const onPointerMove = React.useCallback(
    (e: React.PointerEvent<HTMLDivElement>) => {
      if (connectorDraft) {
        const rect = slideRectRef.current?.getBoundingClientRect();
        if (!rect || !slide) return;
        const cursorEmuX = (e.clientX - rect.left) * connectorDraft.emuPerPx;
        const cursorEmuY = (e.clientY - rect.top) * connectorDraft.emuPerPx;
        const excludeIds = new Set<string>(
          connectorDraft.source.kind === "anchored" ? [connectorDraft.source.shapeId] : []
        );
        const others = collectAllConnectableBoxes(slide.shapes, excludeIds);
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
      if (endpointDraft) {
        const rect = slideRectRef.current?.getBoundingClientRect();
        if (!rect || !slide) return;
        const cursorEmuX = (e.clientX - rect.left) * endpointDraft.emuPerPx;
        const cursorEmuY = (e.clientY - rect.top) * endpointDraft.emuPerPx;
        const others = collectAllConnectableBoxes(slide.shapes, new Set([endpointDraft.shapeId]));
        const snap = snapToAnchor({ x: cursorEmuX, y: cursorEmuY }, others, ANCHOR_THRESHOLD_EMU);
        setEndpointDraft({
          ...endpointDraft,
          currentX: cursorEmuX,
          currentY: cursorEmuY,
          snapped: snap.anchor,
          nearby: snap.nearby,
        });
        return;
      }
      if (waypointDraft) {
        const deltaPx =
          waypointDraft.axis === "horizontal"
            ? e.clientY - waypointDraft.startClientY
            : e.clientX - waypointDraft.startClientX;
        const deltaEmu = Math.round(deltaPx * waypointDraft.emuPerPx);
        setWaypointDraft({
          ...waypointDraft,
          currentValueEmu: waypointDraft.originValueEmu + deltaEmu,
        });
        return;
      }
      if (rotateDraft) {
        const rect = slideRectRef.current?.getBoundingClientRect();
        if (!rect) return;
        const cursorEmuX = (e.clientX - rect.left) * rotateDraft.emuPerPx;
        const cursorEmuY = (e.clientY - rect.top) * rotateDraft.emuPerPx;
        const angle = Math.atan2(
          cursorEmuY - rotateDraft.pivotY,
          cursorEmuX - rotateDraft.pivotX
        );
        setRotateDraft({
          ...rotateDraft,
          currentAngleRad: angle,
          shiftSnap: e.shiftKey,
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
    [
      connectorDraft,
      endpointDraft,
      waypointDraft,
      rotateDraft,
      drag,
      marquee,
      slide,
      slideSize,
      hoveredShapeId,
    ]
  );

  const onPointerUp = React.useCallback(
    async (e: React.PointerEvent<HTMLDivElement>) => {
      containerRef.current?.releasePointerCapture?.(e.pointerId);
      if (connectorDraft && slide) {
        const draft = connectorDraft;
        setConnectorDraft(null);
        // The "kept tool mode active" decision lives in the parent —
        // we always notify on commit/cancel and let the editor decide
        // whether to reset the tool state.
        const sourcePt =
          draft.source.kind === "anchored"
            ? { x: draft.source.x, y: draft.source.y }
            : { x: draft.source.x, y: draft.source.y };
        const startEndDistEmu = Math.hypot(draft.currentX - sourcePt.x, draft.currentY - sourcePt.y);
        // Tiny drags (< ~6 px) are treated as accidental clicks and
        // discarded so the user doesn't end up with a pile of zero-
        // length connectors when they merely tap a port. We still
        // exit tool mode so the next click selects normally.
        if (startEndDistEmu < 6 * draft.emuPerPx) {
          props.onConnectorToolExit?.();
          return;
        }
        try {
          const targetCNvPrId =
            draft.snapped !== null ? findCNvPrIdByShapeId(slide.shapes, draft.snapped.shapeId) : null;
          const startPayload =
            draft.source.kind === "anchored"
              ? {
                  kind: "anchored" as const,
                  targetCNvPrId: draft.source.cNvPrId,
                  side: draft.source.side,
                  t: draft.source.t,
                }
              : { kind: "free" as const, xEmu: draft.source.x, yEmu: draft.source.y };
          const endPayload =
            draft.snapped !== null && targetCNvPrId !== null
              ? {
                  kind: "anchored" as const,
                  targetCNvPrId,
                  side: draft.snapped.side,
                  t: draft.snapped.t,
                }
              : { kind: "free" as const, xEmu: draft.currentX, yEmu: draft.currentY };
          await props.agent.applyCommand({
            type: "pptx:add-connector",
            source: "human",
            payload: {
              slideIndex: props.slideIndex,
              connectorType: draft.connectorType,
              start: startPayload,
              end: endPayload,
            },
          });
        } catch (err) {
          props.onError?.(err as Error);
        }
        props.onConnectorToolExit?.();
        return;
      }
      if (endpointDraft && slide) {
        const draft = endpointDraft;
        setEndpointDraft(null);
        try {
          const targetCNvPrId =
            draft.snapped !== null ? findCNvPrIdByShapeId(slide.shapes, draft.snapped.shapeId) : null;
          const endpointPayload =
            draft.snapped !== null && targetCNvPrId !== null
              ? {
                  kind: "anchored" as const,
                  targetCNvPrId,
                  side: draft.snapped.side,
                  t: draft.snapped.t,
                }
              : { kind: "free" as const, xEmu: draft.currentX, yEmu: draft.currentY };
          await props.agent.applyCommand({
            type: "pptx:set-connector-endpoint",
            source: "human",
            payload: {
              slideIndex: props.slideIndex,
              shapeId: draft.shapeId,
              which: draft.which,
              endpoint: endpointPayload,
            },
          });
        } catch (err) {
          props.onError?.(err as Error);
        }
        return;
      }
      if (rotateDraft) {
        const draft = rotateDraft;
        setRotateDraft(null);
        const deltaDeg = rotateDraftDeltaDeg(draft);
        // Tiny wiggles (< 0.05°) are a no-op — round-tripping near-zero
        // deltas through `set-rotation` would still bump the snapshot
        // revision and pollute undo with empty steps.
        if (Math.abs(deltaDeg) < 0.05) return;
        try {
          for (const t of draft.targets) {
            try {
              await props.agent.applyCommand({
                type: "pptx:set-rotation",
                source: "human",
                payload: {
                  slideIndex: props.slideIndex,
                  shapeId: t.id,
                  degrees: t.originDeg + deltaDeg,
                },
              });
            } catch (innerErr) {
              // Per-shape "not-applicable" (the handler refuses
              // connectors/groups/opaque) is a best-effort skip in a
              // mixed selection — unexpected errors still propagate.
              const code = (innerErr as { code?: string } | null)?.code;
              if (code !== "not-applicable") throw innerErr;
            }
          }
        } catch (err) {
          props.onError?.(err as Error);
        }
        return;
      }
      if (waypointDraft && slide) {
        const draft = waypointDraft;
        setWaypointDraft(null);
        try {
          await props.agent.applyCommand({
            type: "pptx:set-connector-waypoint",
            source: "human",
            payload: {
              slideIndex: props.slideIndex,
              shapeId: draft.shapeId,
              segmentIndex: draft.segmentIndex,
              valueEmu: draft.currentValueEmu,
            },
          });
        } catch (err) {
          props.onError?.(err as Error);
        }
        return;
      }
      if (marquee && slide) {
        // Convert marquee start/end to EMU against the slide rect (not
        // the stage rect) so a drag that begins or ends in the scratch
        // margin produces negative / over-slide coordinates that still
        // intersect off-slide shapes correctly.
        const rect = slideRectRef.current?.getBoundingClientRect();
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
        // have to hunt for a double-click target. For empty non-text
        // placeholders (pic / chart / …) we forward to the activation
        // callback instead — the placeholder needs media or a wizard,
        // not a caret.
        if (mayEnterEditOnClick && targets[0]) {
          const id = targets[0].id;
          const sh = slide ? findShape(slide.shapes, id) : null;
          const placeholderActivation = sh?.kind === "text" ? activatableEmptyPlaceholder(sh) : null;
          if (placeholderActivation) {
            props.onPlaceholderActivate?.({ shapeId: id, placeholder: placeholderActivation });
          } else {
            setEditingId(id);
          }
        }
        return;
      }
      try {
        for (const t of targets) {
          const box = final.boxes.get(t.id);
          if (!box) continue;
          const draggedShape = slide ? findShape(slide.shapes, t.id) : null;
          // Connectors no longer participate in the generic resize
          // path — endpoint and waypoint drags route through their
          // dedicated drafts. A connector still appearing here is a
          // pure body translate, which we commit per endpoint with
          // `set-connector-endpoint` so the model stays consistent and
          // an undo cleanly reverts both ends in one step. Anchored
          // endpoints get DETACHED on body drag: we resolve their live
          // coords (via the same shape map the renderer uses), apply
          // the cursor delta, and write them back as `free`. This
          // matches PowerPoint/Slides where dragging the body of a
          // connected line moves it as a unit and breaks the snap —
          // the user can then re-snap each endpoint by dragging the
          // dots in the selection chrome.
          if (draggedShape?.kind === "connector") {
            if (mode !== "move") continue;
            const dx = box.x - t.origin.x;
            const dy = box.y - t.origin.y;
            if (dx === 0 && dy === 0) continue;
            const startEp = draggedShape.start;
            const endEp = draggedShape.end;
            const startPt =
              startEp.kind === "free"
                ? { x: startEp.xEmu, y: startEp.yEmu }
                : (resolveEndpoint(startEp, shapesByCNvPrId) ?? fallbackEndpoint(draggedShape, "start"));
            const endPt =
              endEp.kind === "free"
                ? { x: endEp.xEmu, y: endEp.yEmu }
                : (resolveEndpoint(endEp, shapesByCNvPrId) ?? fallbackEndpoint(draggedShape, "end"));
            await props.agent.applyCommand({
              type: "pptx:set-connector-endpoint",
              source: "human",
              payload: {
                slideIndex: props.slideIndex,
                shapeId: t.id,
                which: "start",
                endpoint: { kind: "free", xEmu: startPt.x + dx, yEmu: startPt.y + dy },
              },
            });
            await props.agent.applyCommand({
              type: "pptx:set-connector-endpoint",
              source: "human",
              payload: {
                slideIndex: props.slideIndex,
                shapeId: t.id,
                which: "end",
                endpoint: { kind: "free", xEmu: endPt.x + dx, yEmu: endPt.y + dy },
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
    [
      connectorDraft,
      endpointDraft,
      waypointDraft,
      rotateDraft,
      drag,
      marquee,
      props,
      selectedIds,
      setSelectedIds,
      slide,
      slideSize,
    ]
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

  // Esc cancels any in-flight connector gesture (draft draw, endpoint
  // re-attach, waypoint slide) AND exits connector tool mode. PowerPoint
  // / Slides users reach for Esc instinctively whenever they realise
  // they picked the wrong tool, and we want every escape route to land
  // back in the default selection mode without any committed change.
  const onConnectorToolExit = props.onConnectorToolExit;
  React.useEffect(() => {
    function onKey(e: KeyboardEvent) {
      if (e.key !== "Escape") return;
      if (connectorDraft) {
        setConnectorDraft(null);
        onConnectorToolExit?.();
        return;
      }
      if (endpointDraft) {
        setEndpointDraft(null);
        return;
      }
      if (waypointDraft) {
        setWaypointDraft(null);
        return;
      }
      if (rotateDraft) {
        setRotateDraft(null);
        return;
      }
      if (props.connectorTool) {
        onConnectorToolExit?.();
      }
    }
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [connectorDraft, endpointDraft, waypointDraft, rotateDraft, props.connectorTool, onConnectorToolExit]);

  if (!slide) {
    return null;
  }
  const zoom = zoomEarly;
  const dpi = props.dpi ?? DEFAULT_DPI;
  const isToolMode = !!props.connectorTool;
  const slideWUser = px(slideSize.cxEmu);
  const slideHUser = px(slideSize.cyEmu);
  // Pixel rect of the slide card inside the stage div. While the
  // ResizeObserver hasn't reported a size yet (first render of a
  // newly-mounted canvas), fall back to absolute "fill the parent"
  // values so the layout-effect's measurement still produces a
  // sensible viewBox on the very next pass without flashing a wrong
  // aspect.
  const slidePxLeft = stageLayout?.slidePxLeft ?? 0;
  const slidePxTop = stageLayout?.slidePxTop ?? 0;
  const slidePxW = stageLayout?.slidePxW ?? 0;
  const slidePxH = stageLayout?.slidePxH ?? 0;
  const stageViewBox = stageLayout?.stageViewBox ?? slideStageViewBox(slideSize);

  return (
    <StageViewBoxContext.Provider value={stageViewBox}>
      <div
        ref={containerRef}
        data-testid="pptx-slide-canvas"
        data-zoom={zoom.toFixed(2)}
        data-dpi={dpi}
        className="officeai-pptx-canvas"
        style={{
          position: "relative",
          width: "100%",
          height: "100%",
          // The slide rests on the shared "grey desk" backdrop used by
          // every editor surface (var(--page-backdrop), defined in
          // apps/web/app/globals.css). The fallback hex keeps the look
          // sensible when the canvas is rendered outside the app shell
          // (e.g. Storybook, isolated tests, embedded preview).
          background: "var(--page-backdrop, #ecebe8)",
          userSelect: "none",
          cursor: rotateDraft
            ? "grabbing"
            : drag
              ? cursorForDrag(drag.mode)
              : connectorDraft && connectorDraft.snapped
                ? "copy"
                : endpointDraft && endpointDraft.snapped
                  ? "copy"
                  : isToolMode || connectorDraft || endpointDraft
                    ? "crosshair"
                    : "default",
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
          if (!sh) return;
          // "Open in editor" shapes (chart / OLE spreadsheet): hand the
          // activation to the parent so it can pop the embedded-xlsx
          // modal. PowerPoint's analogue is double-clicking a chart /
          // embedded spreadsheet to launch Excel inline.
          if (sh.kind === "chart" || sh.kind === "ole-spreadsheet") {
            props.onShapeActivate?.({ shapeId: id, shape: sh });
            return;
          }
          if (sh.kind !== "text") return;
          // Empty non-text placeholder (pic / chart / tbl / dgm / media):
          // hand the activation to the parent instead of opening the
          // text-edit overlay. Typing into a picture placeholder would
          // produce real `<a:t>` runs and silently turn it into a text
          // box on save, which is exactly the wrong default.
          const placeholderActivation = activatableEmptyPlaceholder(sh as TextShape);
          if (placeholderActivation) {
            props.onPlaceholderActivate?.({ shapeId: id, placeholder: placeholderActivation });
            return;
          }
          startEditing(id);
        }}
      >
        {/* Visual slide card sized to the slide rectangle; pointer events
          fall through to the stage div so a click on the white area
          still reaches `onPointerDown`. The SVG below paints the real
          slide background fill on top of this card. Positioned in
          absolute pixels (vs % of an aspect-locked stage) because the
          stage now fills the user's entire viewport — see
          `useStageLayout`. */}
        <div
          ref={slideRectRef}
          data-testid="pptx-slide-card"
          style={{
            position: "absolute",
            left: `${slidePxLeft}px`,
            top: `${slidePxTop}px`,
            width: `${slidePxW}px`,
            height: `${slidePxH}px`,
            background: "white",
            boxShadow: "0 1px 3px rgba(0,0,0,0.08)",
            pointerEvents: "none",
          }}
        />
        <svg
          xmlns="http://www.w3.org/2000/svg"
          viewBox={stageViewBox}
          preserveAspectRatio="xMidYMid meet"
          style={{ position: "absolute", inset: 0, width: "100%", height: "100%", display: "block" }}
          dangerouslySetInnerHTML={{
            __html: `<rect x="0" y="0" width="${slideWUser}" height="${slideHUser}" fill="${slideBackgroundFillAttr(slide, themeDefault)}"/>${svgInner}${animationBadgesSvg(slide, slideSize, hiddenIds)}`,
          }}
        />
        {drag && preview ? (
          <DragGhostLayer slideSize={slideSize} ghosts={dragGhosts} preview={preview} ctx={ctx} />
        ) : null}
        {rotateDraft ? (
          <RotateGhostLayer slideSize={slideSize} draft={rotateDraft} ctx={ctx} />
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
        {props.remotePeers && props.remotePeers.length > 0 ? (
          <RemoteSelectionOverlay slide={slide} slideSize={slideSize} peers={props.remotePeers} />
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
          endpointDraft={endpointDraft}
          waypointDraft={waypointDraft}
          isRotating={rotateDraft !== null}
          liveRotations={liveRotations}
        />
        {marquee && containerRef.current ? (
          <MarqueeOverlay marquee={marquee} containerRect={containerRef.current.getBoundingClientRect()} />
        ) : null}
        {/* Port-hover layer — surfaced when the user is idling over a
          non-connector shape, so the four cardinal anchor dots become
          drag-from sources for new connectors. We always show ports
          while the connector tool is armed (the user came here
          specifically to draw something), and the live target halo
          stays visible during draft / endpoint-edit gestures. */}
        {/* Slide-wide "you can connect to any of these" hint: shown
          whenever the user has armed the connector tool (about to
          start a brand-new connector) or is mid-drag of an existing
          connector endpoint. Skipped during slide-shape body drags
          and marquee selection because those gestures aren't about
          establishing new attachments. */}
        {!drag && !marquee && (isToolMode || connectorDraft || endpointDraft) ? (
          <ConnectableShapesOverlay
            slide={slide}
            slideSize={slideSize}
            skipId={endpointDraft?.shapeId ?? null}
            emphasisedId={
              connectorDraft?.snapped?.shapeId ?? endpointDraft?.snapped?.shapeId ?? hoveredShapeId
            }
          />
        ) : null}
        {!drag && !marquee && !connectorDraft && !endpointDraft && hoveredShapeId
          ? renderPortHoverOverlay(slide, slideSize, hoveredShapeId, isToolMode)
          : null}
        {connectorDraft && connectorDraft.snapped ? (
          <TargetHaloOverlay slide={slide} slideSize={slideSize} shapeId={connectorDraft.snapped.shapeId} />
        ) : null}
        {endpointDraft && endpointDraft.snapped ? (
          <TargetHaloOverlay slide={slide} slideSize={slideSize} shapeId={endpointDraft.snapped.shapeId} />
        ) : null}
        {connectorDraft ? (
          <ConnectorDraftOverlay slide={slide} slideSize={slideSize} draft={connectorDraft} />
        ) : null}
        {endpointDraft ? (
          <EndpointDraftOverlay slide={slide} slideSize={slideSize} draft={endpointDraft} />
        ) : null}
        {editingId
          ? renderEditingOverlay(
              slide,
              editingId,
              slideSize,
              dpi,
              slide.layoutPartPath ? snap.root.layouts.get(slide.layoutPartPath) : undefined,
              finishEditing,
              onTextSelectionChange,
              editCommitRef
            )
          : null}
        {isToolMode ? <ConnectorToolBanner type={props.connectorTool!.type} /> : null}
      </div>
    </StageViewBoxContext.Provider>
  );
}

function renderPortHoverOverlay(
  slide: Slide,
  slideSize: SlideSize,
  shapeId: string,
  emphasised: boolean
): React.ReactNode {
  const located = findShapeWithOffset(slide.shapes, shapeId);
  if (!located) return null;
  const local = shapeBoundingBox(located.shape);
  if (!local) return null;
  const box: BoundingBox = {
    x: local.x + located.offsetX,
    y: local.y + located.offsetY,
    cx: local.cx,
    cy: local.cy,
  };
  return <PortHoverOverlay slideSize={slideSize} shapeId={shapeId} box={box} emphasised={emphasised} />;
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
 * Recursive group-aware variant used by the connector snap path. Each
 * `<p:grpSp>` translates its children by `position`, so a top-level
 * snap walker that returned local boxes for grouped shapes would land
 * the connector somewhere off-screen. We compute the cumulative
 * translate down each group spine and emit absolute slide-coordinate
 * boxes for every connectable child (text shapes, pictures, charts,
 * tables, opaque shapes). Connectors themselves and group containers
 * are intentionally NOT exposed as snap targets — connectors carry no
 * useful "shape body" to attach to, and snapping to a group's bbox
 * would feel arbitrary when the user clearly meant one of the visible
 * children.
 */
function collectAllConnectableBoxes(
  shapes: ReadonlyArray<Shape>,
  excludeIds: ReadonlySet<string>,
  offsetX: number = 0,
  offsetY: number = 0
): { id: string; box: BoundingBox }[] {
  const out: { id: string; box: BoundingBox }[] = [];
  for (const sh of shapes) {
    if (sh.kind === "group") {
      const childOffsetX = offsetX + (sh.position?.xEmu ?? 0);
      const childOffsetY = offsetY + (sh.position?.yEmu ?? 0);
      out.push(...collectAllConnectableBoxes(sh.children, excludeIds, childOffsetX, childOffsetY));
      continue;
    }
    if (excludeIds.has(sh.id)) continue;
    if (sh.kind === "connector") continue;
    if (sh.cNvPrId <= 0) continue;
    const local = shapeBoundingBox(sh);
    if (!local) continue;
    out.push({
      id: sh.id,
      box: { x: local.x + offsetX, y: local.y + offsetY, cx: local.cx, cy: local.cy },
    });
  }
  return out;
}

/**
 * Locate a shape by id along with the cumulative translate of every
 * group ancestor. Returns `null` when the shape isn't on the slide.
 * The translate matters for port surfacing on grouped children — without
 * it the four port dots would render at the child's local origin
 * instead of where it actually appears on the slide.
 */
function findShapeWithOffset(
  shapes: ReadonlyArray<Shape>,
  id: string,
  offsetX: number = 0,
  offsetY: number = 0
): { shape: Shape; offsetX: number; offsetY: number } | null {
  for (const s of shapes) {
    if (s.id === id) return { shape: s, offsetX, offsetY };
    if (s.kind === "group") {
      const inner = findShapeWithOffset(
        s.children,
        id,
        offsetX + (s.position?.xEmu ?? 0),
        offsetY + (s.position?.yEmu ?? 0)
      );
      if (inner) return inner;
    }
  }
  return null;
}

/**
 * Detect whether a shape is a `prstGeom` line (or arrow-style line).
 * We peek at its `spPrTail` for the `<a:prstGeom prst="line">` marker
 * since the model layer doesn't expose preset as a typed field —
 * everything decorative lives in the opaque tail to keep byte-roundtrip
 * cheap.
 */
function _isCornerHandle(h: ResizeHandle): boolean {
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
function _endpointForHandleSide(h: ResizeHandle): "start" | "end" {
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

/**
 * Last-resort coordinate when an anchored endpoint's target shape has
 * been deleted out from under the connector. We fall back to the
 * connector's own bounding-box corner so the chrome still renders
 * something the user can grab to reconnect, instead of teleporting
 * the dot to (0, 0). `start` = top-left, `end` = bottom-right when
 * `flipH/flipV` are absent — which mirrors how the parser/serializer
 * place the endpoints on a connector's `<a:xfrm>`.
 */
function fallbackEndpoint(connector: ConnectorShape, which: "start" | "end"): { x: number; y: number } {
  const x = connector.position?.xEmu ?? 0;
  const y = connector.position?.yEmu ?? 0;
  const cx = connector.size?.cxEmu ?? 0;
  const cy = connector.size?.cyEmu ?? 0;
  return which === "start" ? { x, y } : { x: x + cx, y: y + cy };
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

/**
 * Set of placeholder types whose "click to fill" affordance is non-
 * textual: a picture placeholder needs an image, a chart placeholder a
 * dataset, a table placeholder a row/column count, and so on. Typing
 * into any of these would silently produce `<a:t>` runs on save and
 * turn the placeholder into a plain text box — exactly the surprise
 * we want to avoid. The renderer paints these with an icon glyph
 * (see `renderPlaceholderHint`); the canvas routes their activation
 * (double-click or PowerPoint-style second-click) to
 * `onPlaceholderActivate` instead of opening the text overlay.
 */
const NON_TEXT_PLACEHOLDER_TYPES: ReadonlySet<string> = new Set(["pic", "chart", "tbl", "dgm", "media"]);

/**
 * Returns the `placeholder` descriptor when `shape` is an empty layout
 * placeholder whose activation should NOT open the text edit overlay
 * (currently: pic, chart, tbl, dgm, media). Returns `null` for any
 * other shape, including `title`/`body`/`subTitle` placeholders (which
 * are perfectly happy receiving caret focus) and for placeholders that
 * already carry user-entered text (we don't want a typed-in title to
 * suddenly fire the file picker on the next click).
 */
function activatableEmptyPlaceholder(shape: TextShape): { type?: string; idx?: number } | null {
  const ph = shape.placeholder;
  if (!ph) return null;
  if (!ph.type || !NON_TEXT_PLACEHOLDER_TYPES.has(ph.type)) return null;
  const hasText = shape.txBody.paragraphs.some((p) =>
    p.runs.some((r) => !r.isLineBreak && r.text.length > 0)
  );
  if (hasText) return null;
  return { type: ph.type, idx: ph.idx };
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

/**
 * Locate a shape by its OOXML `cNvPrId` (encoded as a string). Used
 * by the remote-selection overlay because peers exchange stable
 * OOXML identifiers — the local `Shape.id` is a randomly-minted
 * UUID that is *not* shared across browsers.
 */
function findShapeByCNvPrId(shapes: ReadonlyArray<Shape>, cNvPrIdStr: string): Shape | null {
  const target = Number.parseInt(cNvPrIdStr, 10);
  if (!Number.isFinite(target)) return null;
  for (const s of shapes) {
    if (s.cNvPrId === target) return s;
    if (s.kind === "group") {
      const inner = findShapeByCNvPrId(s.children, cNvPrIdStr);
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

/**
 * Slide-coordinate location of a side anchor on a bounding box.
 * Honours the optional `t` parameter (0..1) so the cardinal midpoint,
 * the two quarter-points, and any future intermediate value resolve to
 * the same coordinate the geometry/route helpers would compute.
 */
function anchorPointFor(box: BoundingBox, side: AnchorSide, t?: number): { x: number; y: number } {
  const u = clampT01(t);
  switch (side) {
    case "n":
      return { x: Math.round(box.x + box.cx * u), y: box.y };
    case "s":
      return { x: Math.round(box.x + box.cx * u), y: box.y + box.cy };
    case "w":
      return { x: box.x, y: Math.round(box.y + box.cy * u) };
    case "e":
      return { x: box.x + box.cx, y: Math.round(box.y + box.cy * u) };
    case "center":
      return { x: Math.round(box.x + box.cx / 2), y: Math.round(box.y + box.cy / 2) };
  }
}

function clampT01(t: number | undefined): number {
  if (t === undefined || !Number.isFinite(t)) return 0.5;
  if (t <= 0) return 0;
  if (t >= 1) return 1;
  return t;
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

  // Rotated single-shape resize routes through `resolveRotatedResize`
  // so the dragged handle tracks the cursor in screen space while the
  // OPPOSITE corner of the rotated body stays anchored — PowerPoint
  // semantics. We bypass anchor-snap here: it expects axis-aligned
  // endpoint coords and connectors (the only line shapes that snap)
  // aren't user-rotatable in our model anyway.
  if (t.originDeg !== 0) {
    const resized = resolveRotatedResize({
      o,
      rotDeg: t.originDeg,
      h,
      dxEmu,
      dyEmu,
      minSize,
    });
    boxes.set(t.id, resized);
    return {
      boxes,
      dx: resized.x - o.x,
      dy: resized.y - o.y,
      dw: resized.cx - o.cx,
      dh: resized.cy - o.cy,
      guides: [],
      anchorCandidates: [],
      anchorSnap: null,
    };
  }

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
interface RotateGhostLayerProps {
  readonly slideSize: SlideSize;
  readonly draft: RotateDraft;
  readonly ctx: SvgRenderCtx;
}

/**
 * Live preview while the user drags the rotation grip. We re-emit each
 * target's SVG with the new rotation baked into `shape.rotation`, so
 * the existing `<g transform="rotate(deg cx cy)">` wrapper that
 * `shapeToSvg` already produces does the visual work — no parallel
 * code path that could drift from the post-commit render. The
 * underlying static layer hides these shapes (`hiddenIds`) so the
 * user only sees the rotated copy.
 *
 * `shapeToSvg` is plain string concat and runs at most a few times per
 * pointermove (one per selected rotatable shape), so cost is the same
 * order as the resize ghost path.
 */
function RotateGhostLayer({ slideSize, draft, ctx }: RotateGhostLayerProps): React.ReactElement {
  const stageViewBox = useStageViewBox(slideSize);
  const deltaDeg = rotateDraftDeltaDeg(draft);
  const inner = draft.targets
    .map((t) => {
      const synth = { ...t.shape, rotation: t.originDeg + deltaDeg } as Shape;
      return shapeToSvg(synth, ctx);
    })
    .join("");
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      viewBox={stageViewBox}
      preserveAspectRatio="xMidYMid meet"
      pointerEvents="none"
      style={{
        position: "absolute",
        inset: 0,
        width: "100%",
        height: "100%",
        opacity: 0.95,
      }}
      dangerouslySetInnerHTML={{ __html: inner }}
    />
  );
}

function DragGhostLayer({ slideSize, ghosts, preview, ctx }: DragGhostLayerProps): React.ReactElement {
  const stageViewBox = useStageViewBox(slideSize);
  const inner = ghosts
    .map((g) => {
      const box = preview.boxes.get(g.id);
      if (!box) return "";
      const rot =
        "rotation" in g.shape && typeof g.shape.rotation === "number"
          ? g.shape.rotation
          : 0;
      if (g.shape.kind === "text" || rot !== 0) {
        // Text and rotated shapes both need a fresh `shapeToSvg` so
        // word-wrap (text) and the rotation pivot (rotated) recalc
        // around the new centre. Wrapping the baked SVG in a
        // translate+scale would scale the rotated body along screen
        // axes and break the screen-space anchor invariant the
        // rotation-aware resize math relies on.
        const synth = {
          ...g.shape,
          position: { xEmu: box.x, yEmu: box.y },
          size: { cxEmu: box.cx, cyEmu: box.cy },
        } as Shape;
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
      viewBox={stageViewBox}
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
  /** Active endpoint-edit gesture so the chrome can hide that endpoint dot. */
  readonly endpointDraft: EndpointEditDraft | null;
  /** Active waypoint-slide gesture so the elbow preview uses live offsets. */
  readonly waypointDraft: WaypointDraft | null;
  /** True while the rotation grip is being dragged — dims resize handles
   * and keeps the grip styled "engaged" until pointerup commits. */
  readonly isRotating: boolean;
  /**
   * Per-shape rotation override (degrees) used while the rotation grip
   * is being dragged. When provided, takes precedence over the
   * snapshot's `shape.rotation` so the chrome tracks the live ghost
   * frame-by-frame instead of snapping back to the pre-gesture pose.
   * `null` (the common case) means "use the saved rotation on each
   * shape", which is what we want once the gesture has committed.
   */
  readonly liveRotations: ReadonlyMap<string, number> | null;
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
  endpointDraft,
  waypointDraft,
  isRotating,
  liveRotations,
}: SelectionOverlayProps): React.ReactElement | null {
  const stageViewBox = useStageViewBox(slideSize);
  if (selectedIds.length === 0) return null;
  // Suppress chrome (outline + handles + move zone) for the shape being
  // text-edited. The TextEditOverlay paints a dashed border itself and
  // the contenteditable owns hit-testing, so any extra rect here would
  // either visually double the border or steal pointer events away
  // from the caret.
  const visibleIds = editingId ? selectedIds.filter((id) => id !== editingId) : selectedIds;
  if (visibleIds.length === 0) return null;

  // Single-connector selection gets its own dedicated chrome layer:
  // two endpoint dots and (for elbow) one perpendicular slider per
  // interior segment. We render this BEFORE the generic bbox/8-handle
  // path because connectors don't want either of those — a connector's
  // bounding box is mostly empty space, the eight handles distort the
  // route, and PowerPoint/Slides agree that endpoint + segment
  // controls are the only sensible affordances.
  if (visibleIds.length === 1) {
    const onlyId = visibleIds[0];
    const sh = findShape(slide.shapes, onlyId);
    if (sh && sh.kind === "connector") {
      return (
        <ConnectorSelectionChrome
          slide={slide}
          slideSize={slideSize}
          connector={sh}
          endpointDraft={endpointDraft}
          waypointDraft={waypointDraft}
        />
      );
    }
  }

  const entries: { id: string; box: BoundingBox; rotatable: boolean; rotation: number }[] = [];
  for (const id of visibleIds) {
    const sh = findShape(slide.shapes, id);
    if (!sh) continue;
    const base = shapeBoundingBox(sh);
    if (!base) continue;
    const previewBox = previewBoxes?.get(id) ?? null;
    // Same kinds the `pptx:set-rotation` handler refuses; the rotation
    // grip skips rendering for selections that contain no rotatable
    // shapes so we don't tease an affordance the gesture would no-op.
    const rotatable = sh.kind !== "connector" && sh.kind !== "group" && sh.kind !== "opaque";
    // The rotate-grip drag publishes a `liveRotations` map so the
    // chrome tracks the ghost. When that's absent we fall back to the
    // committed rotation from the snapshot — which the renderer uses
    // for the static SVG layer too, so chrome and shape stay in sync
    // post-commit.
    const liveRot = liveRotations?.get(id);
    const savedRot =
      "rotation" in sh && typeof sh.rotation === "number" ? sh.rotation : 0;
    const rotation = liveRot ?? savedRot;
    entries.push({ id, box: previewBox ?? base, rotatable, rotation });
  }
  if (entries.length === 0) return null;
  const isMulti = entries.length > 1;
  const isResizing = dragMode !== null && dragMode !== "move";
  const hasRotatable = entries.some((e) => e.rotatable);
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
      viewBox={stageViewBox}
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
          select but kept consistent for clarity). Each shape's outline
          rotates with its own `rotation` so the dashed box stays
          glued to the rotated body — pivot is the shape's
          (axis-aligned, pre-rotation) bbox centre, the same pivot
          `wrapWithRotation` in `svg/shapes.ts` uses for the static
          SVG layer so chrome and shape stay locked frame-by-frame. */}
      {isMulti
        ? entries.map((e) => {
            const cxPx = px(e.box.x + e.box.cx / 2);
            const cyPx = px(e.box.y + e.box.cy / 2);
            const rotAttr =
              e.rotation !== 0 ? `rotate(${e.rotation} ${cxPx} ${cyPx})` : undefined;
            return (
              <g key={`outline-${e.id}`} transform={rotAttr}>
                <rect
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
              </g>
            );
          })
        : null}
      {/* Move hit-zones — one per selected shape so a click on any of
          them moves the entire group. We deliberately use per-shape
          rects (not the union) so the user can click between two
          selected shapes without grabbing the gap. Each rect rotates
          with its shape so the click target follows the rotated body
          rather than its axis-aligned bbox (otherwise the user would
          have to chase the original pose to start a move). */}
      {entries.map((e) => {
        const cxPx = px(e.box.x + e.box.cx / 2);
        const cyPx = px(e.box.y + e.box.cy / 2);
        const rotAttr =
          e.rotation !== 0 ? `rotate(${e.rotation} ${cxPx} ${cyPx})` : undefined;
        return (
          <g key={`move-${e.id}`} transform={rotAttr}>
            <rect
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
          </g>
        );
      })}
      {/*
       * Union outline + 8 resize handles + rotate grip.
       *
       * For single-shape selections the "union" IS the shape, so we
       * wrap the whole bundle in the shape's rotation transform and
       * the chrome stays glued to the rotated body — handles appear
       * at the rotated corners, the grip floats above the rotated
       * top edge.
       *
       * For multi-shape we keep the union/handles/grip axis-aligned
       * around the combined bbox: per-shape rotations differ so a
       * rotated union would mislead about which axis the resize
       * handles act along, and PowerPoint matches this behaviour.
       */}
      {(() => {
        const wrapRot =
          !isMulti && primary.rotation !== 0
            ? `rotate(${primary.rotation} ${ux + ucx / 2} ${uy + ucy / 2})`
            : undefined;
        const handleCursor = (h: ResizeHandle): string =>
          !isMulti && primary.rotation !== 0
            ? cursorForRotatedHandle(h, primary.rotation)
            : cursorForHandle(h);
        return (
          <g transform={wrapRot}>
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
            {/* Resize handles. On multi-select they wrap the union
                box and scale every selected shape proportionally; on
                single-select they wrap the shape itself (and inherit
                the wrapper's rotation). During a resize gesture we
                keep them rendered so the user sees the corner they're
                pulling. We render two rects per handle: an invisible
                larger hit-zone on top so pointer capture is forgiving,
                and the visible purple-bordered square underneath. */}
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
                  opacity={isResizing || isRotating ? 0.6 : 1}
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
                  style={{ cursor: handleCursor(it.h) }}
                />
              </g>
            ))}
            {/* Rotation grip — single circle floating above the union
                top edge with a short tether line. Skipped when the
                selection contains nothing rotatable (e.g. only a
                connector) so we don't dangle an affordance the
                gesture would no-op. */}
            {hasRotatable
              ? (() => {
                  const gripGap = 24;
                  const gripR = 7;
                  const gripHitR = 14;
                  const gx = ux + ucx / 2;
                  const gy = uy - gripGap;
                  return (
                    <g key="rotate-grip">
                      <line
                        x1={gx}
                        y1={uy}
                        x2={gx}
                        y2={gy + gripR}
                        stroke="#7c3aed"
                        strokeWidth={1.5}
                        vectorEffect="non-scaling-stroke"
                        pointerEvents="none"
                        opacity={isResizing ? 0.6 : 1}
                      />
                      <circle
                        cx={gx}
                        cy={gy}
                        r={gripR}
                        fill={isRotating ? "#7c3aed" : "#ffffff"}
                        stroke="#7c3aed"
                        strokeWidth={1.5}
                        vectorEffect="non-scaling-stroke"
                        pointerEvents="none"
                        opacity={isResizing ? 0.6 : 1}
                      />
                      <circle
                        data-shape-id={escAttr(primary.id)}
                        data-handle="rotate"
                        cx={gx}
                        cy={gy}
                        r={gripHitR}
                        fill="transparent"
                        pointerEvents="auto"
                        style={{ cursor: "grab" }}
                      />
                    </g>
                  );
                })()
              : null}
          </g>
        );
      })()}
    </svg>
  );
}

interface ConnectorSelectionChromeProps {
  readonly slide: Slide;
  readonly slideSize: SlideSize;
  readonly connector: ConnectorShape;
  readonly endpointDraft: EndpointEditDraft | null;
  readonly waypointDraft: WaypointDraft | null;
}

/**
 * Connector-specific selection chrome. We deliberately avoid drawing
 * the dashed bbox + 8 resize handles that other shapes get — those
 * controls are nonsensical for a 1-D primitive (the bbox is mostly
 * empty space, and "resizing" by a corner handle would arbitrarily
 * distort the route). Instead we paint:
 *
 *   - a soft outline of the route itself, so the user sees what's
 *     actually selected;
 *   - one large dot per endpoint, which the canvas pointerdown handler
 *     promotes to an `EndpointEditDraft`;
 *   - one perpendicular slider per interior elbow segment (only when
 *     the connector is the elbow type), which becomes a
 *     `WaypointDraft` on grab.
 *
 * Endpoints render in two visual variants: anchored = filled purple
 * circle (it's "stuck" to a target), free = white-filled circle with
 * purple ring (it's floating in space). When the user is mid-drag on
 * one endpoint we hide that dot — the live preview overlay paints
 * the new position with a snap halo, so doubling up would be noisy.
 */
function ConnectorSelectionChrome({
  slide,
  slideSize,
  connector,
  endpointDraft,
  waypointDraft,
}: ConnectorSelectionChromeProps): React.ReactElement | null {
  const stageViewBox = useStageViewBox(slideSize);
  const shapesByCNvPrId = React.useMemo(() => buildShapesByCNvPrId(slide.shapes), [slide.shapes]);
  const startPt = resolveEndpoint(connector.start, shapesByCNvPrId) ?? fallbackEndpoint(connector, "start");
  const endPt = resolveEndpoint(connector.end, shapesByCNvPrId) ?? fallbackEndpoint(connector, "end");
  const startSide = connector.start.kind === "anchored" ? connector.start.side : null;
  const endSide = connector.end.kind === "anchored" ? connector.end.side : null;
  // Live waypoint preview: substitute the dragged offset into the
  // visible polyline so the user sees the segment slide before the
  // command commits. Geometry mirrors `routeElbow` in the SVG renderer.
  const liveWaypoints = (() => {
    if (!waypointDraft || waypointDraft.shapeId !== connector.id) return connector.waypoints;
    const arr = [...(connector.waypoints ?? [])];
    while (arr.length <= waypointDraft.segmentIndex) arr.push(0);
    arr[waypointDraft.segmentIndex] = waypointDraft.currentValueEmu;
    return arr;
  })();
  const obstacles = React.useMemo(
    () => collectObstaclesExcluding(slide.shapes, connector),
    [slide.shapes, connector]
  );
  const points = computeRoutePoints(
    connector.connectorType,
    startPt,
    endPt,
    startSide,
    endSide,
    liveWaypoints,
    obstacles
  );
  const editingStart = endpointDraft?.shapeId === connector.id && endpointDraft.which === "start";
  const editingEnd = endpointDraft?.shapeId === connector.id && endpointDraft.which === "end";
  const r = 110_000; // ≈ 11.5 px @ 96 DPI; large enough to grab without zoom-in
  const hitR = 220_000;
  const purple = "#7c3aed";

  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      viewBox={stageViewBox}
      preserveAspectRatio="xMidYMid meet"
      pointerEvents="none"
      style={{ position: "absolute", inset: 0, width: "100%", height: "100%" }}
    >
      {/* Route highlight — sits behind the dots so they always read
          on top. Slightly thicker and more transparent than the actual
          stroke so the user perceives "selected" without obscuring the
          rendered connector. */}
      {connector.connectorType === "curved" ? (
        <path
          d={curvedPath(points)}
          fill="none"
          stroke={purple}
          strokeOpacity={0.35}
          strokeWidth={6}
          vectorEffect="non-scaling-stroke"
        />
      ) : (
        <polyline
          points={points.map((p) => `${px(p.x)},${px(p.y)}`).join(" ")}
          fill="none"
          stroke={purple}
          strokeOpacity={0.35}
          strokeWidth={6}
          vectorEffect="non-scaling-stroke"
        />
      )}

      {/* Endpoint dots */}
      {!editingStart ? (
        <ConnectorEndpointHandle
          shapeId={connector.id}
          which="start"
          x={startPt.x}
          y={startPt.y}
          r={r}
          hitR={hitR}
          anchored={connector.start.kind === "anchored"}
          color={purple}
        />
      ) : null}
      {!editingEnd ? (
        <ConnectorEndpointHandle
          shapeId={connector.id}
          which="end"
          x={endPt.x}
          y={endPt.y}
          r={r}
          hitR={hitR}
          anchored={connector.end.kind === "anchored"}
          color={purple}
        />
      ) : null}

      {/* Waypoint sliders for elbow connectors. We expose one slider
          per interior segment (not the lead-out stubs at either end);
          each one slides perpendicular to the segment's axis. Skip
          when the connector type isn't elbow — straight has no bend,
          and curved doesn't expose explicit waypoints in our model. */}
      {connector.connectorType === "elbow" ? renderWaypointSliders(connector.id, points, purple) : null}
    </svg>
  );
}

interface ConnectorEndpointHandleProps {
  readonly shapeId: string;
  readonly which: "start" | "end";
  readonly x: number;
  readonly y: number;
  readonly r: number;
  readonly hitR: number;
  readonly anchored: boolean;
  readonly color: string;
}

function ConnectorEndpointHandle({
  shapeId,
  which,
  x,
  y,
  r,
  hitR,
  anchored,
  color,
}: ConnectorEndpointHandleProps): React.ReactElement {
  return (
    <g>
      <circle
        cx={px(x)}
        cy={px(y)}
        r={px(r)}
        fill={anchored ? color : "white"}
        stroke={color}
        strokeWidth={2}
        vectorEffect="non-scaling-stroke"
        pointerEvents="none"
      />
      <circle
        data-connector-shape-id={escAttr(shapeId)}
        data-connector-endpoint={which}
        cx={px(x)}
        cy={px(y)}
        r={px(hitR)}
        fill="transparent"
        style={{ pointerEvents: "all", cursor: "grab" }}
      />
    </g>
  );
}

/**
 * Emit one perpendicular slider per interior segment of an elbow
 * route. The route polyline always starts and ends at the endpoint
 * dots; the segments in between are the ones the user can slide.
 * Each slider draws as a small "pill" centered on the segment
 * midpoint, oriented across the segment's axis.
 */
function renderWaypointSliders(
  shapeId: string,
  points: ReadonlyArray<{ x: number; y: number }>,
  color: string
): React.ReactNode {
  if (points.length < 4) return null;
  const out: React.ReactNode[] = [];
  // The first segment (points[0]→points[1]) is the lead-out stub from
  // the start endpoint; the last segment is the lead-in stub to the
  // end endpoint. Skip both — only interior segments are user-
  // draggable. Index counting matches the `waypoints` array indexing
  // used by `set-connector-waypoint`.
  for (let i = 1; i < points.length - 2; i++) {
    const a = points[i];
    const b = points[i + 1];
    const isHoriz = Math.abs(b.y - a.y) < 1;
    const isVert = Math.abs(b.x - a.x) < 1;
    if (!isHoriz && !isVert) continue;
    const midX = (a.x + b.x) / 2;
    const midY = (a.y + b.y) / 2;
    const segLen = Math.hypot(b.x - a.x, b.y - a.y);
    if (segLen < 200_000) continue; // too short to grab without overlapping endpoint dots
    const axis: "horizontal" | "vertical" = isHoriz ? "horizontal" : "vertical";
    const w = isHoriz ? 280_000 : 90_000;
    const h = isHoriz ? 90_000 : 280_000;
    const segIndex = i - 1; // 0-based among interior segments
    // Origin value = the value we'd send back if the user nudged it
    // by 0; the canvas reads this off the data attribute on
    // pointerdown so the live drag has a baseline. For a horizontal
    // segment that's its y-coordinate, for a vertical segment its
    // x-coordinate.
    const origin = isHoriz ? a.y : a.x;
    out.push(
      <g key={`wp-${i}`}>
        <rect
          x={px(midX - w / 2)}
          y={px(midY - h / 2)}
          width={px(w)}
          height={px(h)}
          rx={px(45_000)}
          ry={px(45_000)}
          fill="white"
          stroke={color}
          strokeWidth={1.5}
          vectorEffect="non-scaling-stroke"
          pointerEvents="none"
          opacity={0.95}
        />
        <rect
          data-connector-shape-id={escAttr(shapeId)}
          data-connector-waypoint={String(segIndex)}
          data-connector-axis={axis}
          data-connector-origin={String(Math.round(origin))}
          x={px(midX - w)}
          y={px(midY - h)}
          width={px(w * 2)}
          height={px(h * 2)}
          fill="transparent"
          style={{
            pointerEvents: "all",
            cursor: isHoriz ? "ns-resize" : "ew-resize",
          }}
        />
      </g>
    );
  }
  return out;
}

/**
 * Slide-wide obstacle list with the connector's own anchored target
 * shapes excluded so the router doesn't try to route around the very
 * shapes the connector is anchored to. Free endpoints contribute no
 * exclusion.
 */
function collectObstaclesExcluding(
  shapes: ReadonlyArray<Shape>,
  connector: ConnectorShape
): RouterObstacle[] {
  const exclude = new Set<number>();
  if (connector.start.kind === "anchored") exclude.add(connector.start.targetCNvPrId);
  if (connector.end.kind === "anchored") exclude.add(connector.end.targetCNvPrId);
  return collectObstacles(shapes, exclude);
}

/**
 * Build the visible route polyline for a connector for use by the
 * selection chrome and the waypoint preview. Delegates to the shared
 * `routeConnector` engine in `connector-router/` so what the user
 * clicks on lines up byte-for-byte with what the SVG renderer paints.
 *
 * `obstacles` is optional — chrome callers that don't have a slide
 * handy (e.g. the draft preview before the connector is committed)
 * can omit it and still get a sane heuristic route.
 */
function computeRoutePoints(
  type: ConnectorType,
  sp: { x: number; y: number },
  ep: { x: number; y: number },
  startSide: AnchorSide | null,
  endSide: AnchorSide | null,
  waypoints: ReadonlyArray<number> | undefined,
  obstacles?: ReadonlyArray<RouterObstacle>
): ReadonlyArray<{ x: number; y: number }> {
  return routeConnectorShared(type, sp, ep, startSide, endSide, {
    waypoints,
    obstacles,
  }).points;
}

function curvedPath(points: ReadonlyArray<{ x: number; y: number }>): string {
  if (points.length === 0) return "";
  if (points.length === 1) return `M ${px(points[0].x)} ${px(points[0].y)}`;
  if (points.length === 2) {
    return `M ${px(points[0].x)} ${px(points[0].y)} L ${px(points[1].x)} ${px(points[1].y)}`;
  }
  if (points.length >= 4) {
    const [a, c1, c2, b] = points;
    return `M ${px(a.x)} ${px(a.y)} C ${px(c1.x)} ${px(c1.y)} ${px(c2.x)} ${px(c2.y)} ${px(b.x)} ${px(b.y)}`;
  }
  const [a, c, b] = [points[0], points[1], points[points.length - 1]];
  return `M ${px(a.x)} ${px(a.y)} Q ${px(c.x)} ${px(c.y)} ${px(b.x)} ${px(b.y)}`;
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
  const stageViewBox = useStageViewBox(slideSize);
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      viewBox={stageViewBox}
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
  const stageViewBox = useStageViewBox(slideSize);
  const r = 80_000; // ≈ 8.4 px @ 96 DPI; readable but not invasive
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      viewBox={stageViewBox}
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
  /**
   * When true (e.g. user has armed the connector tool, or the live
   * draft snapped onto this shape), the dots render larger and with a
   * tinted halo so the user clearly sees they're a snap target.
   */
  readonly emphasised?: boolean;
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
function PortHoverOverlay({
  slideSize,
  shapeId,
  box,
  emphasised,
}: PortHoverOverlayProps): React.ReactElement {
  const stageViewBox = useStageViewBox(slideSize);
  const sides: AnchorSide[] = ["n", "s", "e", "w"];
  // We use the cardinal midpoints (t=0.5) for hover ports — the
  // quarter-points exist in the model but cluttering the four-cardinal
  // affordance with eight extra dots makes the canvas look like a
  // pegboard. Snapping during a live draft already considers all 13
  // anchors, so the user can land a connector at a quarter-point even
  // though only the cardinals draw a hover dot.
  const allAnchors = anchorsFor(shapeId, box);
  const anchors = sides
    .map((side) => allAnchors.find((a) => a.side === side && Math.abs(a.t - 0.5) < 1e-3))
    .filter((a): a is ShapeAnchor => !!a);
  const r = emphasised ? 85_000 : 65_000;
  const haloColor = emphasised ? "rgba(14,165,233,0.35)" : "rgba(14,165,233,0.15)";
  const ringWidth = emphasised ? 2 : 1.5;
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      viewBox={stageViewBox}
      preserveAspectRatio="xMidYMid meet"
      style={{ position: "absolute", inset: 0, width: "100%", height: "100%", pointerEvents: "none" }}
    >
      {anchors.map((a) => (
        <g key={a.side}>
          <circle
            cx={px(a.x)}
            cy={px(a.y)}
            r={px(r * 1.6)}
            fill={haloColor}
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
            strokeWidth={ringWidth}
            vectorEffect="non-scaling-stroke"
            pointerEvents="none"
          />
        </g>
      ))}
    </svg>
  );
}

interface TargetHaloOverlayProps {
  readonly slide: Slide;
  readonly slideSize: SlideSize;
  readonly shapeId: string;
}

/**
 * Soft sky-blue halo painted around whichever shape the live snap is
 * about to attach to. Distinct from the port hover overlay — the user
 * needs to know not just which port will catch the endpoint, but which
 * SHAPE they're connecting to (especially in dense slides). The halo
 * also reveals all four cardinal ports on the target so the user can
 * see how the connector might re-route if they slide the cursor.
 */
function TargetHaloOverlay({ slide, slideSize, shapeId }: TargetHaloOverlayProps): React.ReactElement | null {
  const stageViewBox = useStageViewBox(slideSize);
  const located = findShapeWithOffset(slide.shapes, shapeId);
  if (!located) return null;
  const local = shapeBoundingBox(located.shape);
  if (!local) return null;
  const box: BoundingBox = {
    x: local.x + located.offsetX,
    y: local.y + located.offsetY,
    cx: local.cx,
    cy: local.cy,
  };
  // Pad slightly so the halo doesn't visually merge with the shape
  // outline at zoom levels where strokes touch.
  const pad = 40_000;
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      viewBox={stageViewBox}
      preserveAspectRatio="xMidYMid meet"
      pointerEvents="none"
      style={{ position: "absolute", inset: 0, width: "100%", height: "100%" }}
    >
      <rect
        x={px(box.x - pad)}
        y={px(box.y - pad)}
        width={px(box.cx + pad * 2)}
        height={px(box.cy + pad * 2)}
        rx={px(40_000)}
        ry={px(40_000)}
        fill="none"
        stroke="#0ea5e9"
        strokeOpacity={0.6}
        strokeWidth={2}
        strokeDasharray="6 3"
        vectorEffect="non-scaling-stroke"
      />
    </svg>
  );
}

interface ConnectableShapesOverlayProps {
  readonly slide: Slide;
  readonly slideSize: SlideSize;
  /**
   * Optional ID to skip — typically the connector currently being
   * edited so it doesn't paint a halo around itself. Connectors are
   * already excluded from the snap walker, but we accept any ID for
   * symmetry with the dragger / draft state.
   */
  readonly skipId?: string | null;
  /**
   * Optional ID to suppress because another, more emphatic halo is
   * already painted on top of it (the live target halo, or a port
   * hover overlay). Suppressing prevents visual double-painting.
   */
  readonly emphasisedId?: string | null;
}

/**
 * Faint dashed outline drawn around every shape on the slide that a
 * connector endpoint can snap to. Surfaced whenever the connector
 * tool is armed or the user is mid-drag of an existing endpoint, so
 * the answer to "where can I land this?" is visible at a glance
 * across the whole slide rather than only on the shape currently
 * under the cursor. Deliberately lightweight (1px, low opacity, no
 * fill) so it never competes with the live target halo or with the
 * shape contents themselves.
 */
function ConnectableShapesOverlay({
  slide,
  slideSize,
  skipId,
  emphasisedId,
}: ConnectableShapesOverlayProps): React.ReactElement | null {
  const stageViewBox = useStageViewBox(slideSize);
  const exclude = new Set<string>();
  if (skipId) exclude.add(skipId);
  const boxes = collectAllConnectableBoxes(slide.shapes, exclude);
  if (boxes.length === 0) return null;
  // Pad just outside the shape's bounding box so the hint outline
  // never touches the shape stroke. Smaller pad than the active
  // target halo so it visually subordinates to it when both are on.
  const pad = 24_000;
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      viewBox={stageViewBox}
      preserveAspectRatio="xMidYMid meet"
      pointerEvents="none"
      data-testid="pptx-connector-affordance-overlay"
      style={{ position: "absolute", inset: 0, width: "100%", height: "100%" }}
    >
      {boxes.map(({ id, box }) => {
        // The shape that's about to receive the snap is already
        // outlined by the brighter `TargetHaloOverlay`; suppress the
        // hint here so we don't double-paint it with two strokes.
        if (id === emphasisedId) return null;
        // Four cardinal port dots — n / e / s / w midpoints. These
        // give the user an "I can land here" affordance without them
        // having to hover over the shape body first. Quarter-points
        // stay reachable via the snap engine; we deliberately don't
        // surface them here to avoid pegboard clutter.
        const cx = box.x + box.cx / 2;
        const cy = box.y + box.cy / 2;
        const portR = 30_000; // ≈ 3 px @ 96 DPI — small but visible
        return (
          <g key={id}>
            <rect
              x={px(box.x - pad)}
              y={px(box.y - pad)}
              width={px(box.cx + pad * 2)}
              height={px(box.cy + pad * 2)}
              rx={px(28_000)}
              ry={px(28_000)}
              fill="none"
              stroke="#0ea5e9"
              strokeOpacity={0.35}
              strokeWidth={1}
              strokeDasharray="4 4"
              vectorEffect="non-scaling-stroke"
            />
            <circle cx={px(cx)} cy={px(box.y)} r={px(portR)} fill="#0ea5e9" fillOpacity={0.55} />
            <circle cx={px(box.x + box.cx)} cy={px(cy)} r={px(portR)} fill="#0ea5e9" fillOpacity={0.55} />
            <circle cx={px(cx)} cy={px(box.y + box.cy)} r={px(portR)} fill="#0ea5e9" fillOpacity={0.55} />
            <circle cx={px(box.x)} cy={px(cy)} r={px(portR)} fill="#0ea5e9" fillOpacity={0.55} />
          </g>
        );
      })}
    </svg>
  );
}

interface ConnectorDraftOverlayProps {
  readonly slide: Slide;
  readonly slideSize: SlideSize;
  readonly draft: ConnectorDraft;
}

/**
 * Live preview of a connector being drawn from a port. Renders a
 * dashed orthogonal polyline from the source anchor to the current
 * pointer position. When a candidate destination anchor is in range
 * we paint it as a solid sky-blue dot to confirm the snap target so
 * the user knows the destination will be anchored (not free).
 *
 * Obstacles for routing exclude the source anchor's shape and any
 * snap target's shape so the preview never tries to "go around" the
 * very shapes the route is intended to land on.
 */
function ConnectorDraftOverlay({ slide, slideSize, draft }: ConnectorDraftOverlayProps): React.ReactElement {
  const stageViewBox = useStageViewBox(slideSize);
  const sx = draft.source.x;
  const sy = draft.source.y;
  const ex = draft.snapped ? draft.snapped.x : draft.currentX;
  const ey = draft.snapped ? draft.snapped.y : draft.currentY;
  const startSide = draft.source.kind === "anchored" ? draft.source.side : null;
  const endSide = draft.snapped?.side ?? null;
  const obstacles = React.useMemo(() => {
    const exclude = new Set<number>();
    if (draft.source.kind === "anchored") {
      const sourceShape = findShape(slide.shapes, draft.source.shapeId);
      if (sourceShape && sourceShape.cNvPrId > 0) exclude.add(sourceShape.cNvPrId);
    }
    if (draft.snapped) {
      const snapShape = findShape(slide.shapes, draft.snapped.shapeId);
      if (snapShape && snapShape.cNvPrId > 0) exclude.add(snapShape.cNvPrId);
    }
    return collectObstacles(slide.shapes, exclude);
  }, [slide.shapes, draft.source, draft.snapped]);
  const draftPath = routeForType(
    draft.connectorType,
    { x: sx, y: sy },
    { x: ex, y: ey },
    startSide,
    endSide,
    obstacles
  );
  // Magnetic snap radius interpolates by approach distance: when the
  // pointer is right on a port the dot doubles in size; far away it
  // stays at the baseline. Pure visual cue, no model effect.
  const baseR = 80_000;
  const r = baseR;
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      viewBox={stageViewBox}
      preserveAspectRatio="xMidYMid meet"
      pointerEvents="none"
      style={{ position: "absolute", inset: 0, width: "100%", height: "100%" }}
    >
      {draftPath}
      {draft.source.kind === "anchored" ? (
        <circle
          cx={px(sx)}
          cy={px(sy)}
          r={px(r)}
          fill="#0ea5e9"
          stroke="white"
          strokeWidth={1.5}
          vectorEffect="non-scaling-stroke"
        />
      ) : (
        <circle
          cx={px(sx)}
          cy={px(sy)}
          r={px(r * 0.6)}
          fill="white"
          stroke="#0ea5e9"
          strokeWidth={1.5}
          strokeDasharray="2 2"
          vectorEffect="non-scaling-stroke"
        />
      )}
      {draft.nearby.map((a, i) => {
        const isSnap =
          draft.snapped !== null &&
          a.shapeId === draft.snapped.shapeId &&
          a.side === draft.snapped.side &&
          Math.abs(a.t - draft.snapped.t) < 1e-3;
        const dist = Math.hypot(a.x - draft.currentX, a.y - draft.currentY);
        // Magnetic snap: dot radius interpolates with how close the
        // cursor is. At the port itself the dot is full size; at the
        // edge of the threshold it shrinks to half. Pure visual cue;
        // the snap decision still happens in `snapToAnchor`.
        const t = Math.max(0, Math.min(1, 1 - dist / ANCHOR_THRESHOLD_EMU));
        const magnetR = baseR * (0.5 + t * 0.5);
        return (
          <circle
            key={`${a.shapeId}-${a.side}-${a.t}-${i}`}
            cx={px(a.x)}
            cy={px(a.y)}
            r={px(isSnap ? Math.max(magnetR, baseR) : magnetR)}
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

interface EndpointDraftOverlayProps {
  readonly slide: Slide;
  readonly slideSize: SlideSize;
  readonly draft: EndpointEditDraft;
}

/**
 * Live preview while the user is dragging an existing connector's
 * endpoint. Mirrors `ConnectorDraftOverlay` but routes from the
 * un-dragged endpoint to the cursor / snap target so the overlay
 * matches the post-commit path. We also draw a small "detached" hint
 * (open ring) when the endpoint has left the snap range, so the user
 * understands they're about to convert an anchored endpoint to a
 * free one.
 */
function EndpointDraftOverlay({ slide, slideSize, draft }: EndpointDraftOverlayProps): React.ReactElement {
  const stageViewBox = useStageViewBox(slideSize);
  const sx = draft.otherPoint.x;
  const sy = draft.otherPoint.y;
  const ex = draft.snapped ? draft.snapped.x : draft.currentX;
  const ey = draft.snapped ? draft.snapped.y : draft.currentY;
  const startSide = draft.otherSide;
  const endSide = draft.snapped?.side ?? null;
  const obstacles = React.useMemo(() => {
    const exclude = new Set<number>();
    // The connector being edited usually has its other endpoint
    // anchored to a shape; that shape isn't an obstacle for the
    // routing of this very connector. We also exclude any snap-target
    // shape so the preview doesn't try to detour around the shape it
    // is about to land on.
    const otherCNvPrId = draft.otherEndpointCNvPrId;
    if (otherCNvPrId !== null && otherCNvPrId > 0) exclude.add(otherCNvPrId);
    if (draft.snapped) {
      const snapShape = findShape(slide.shapes, draft.snapped.shapeId);
      if (snapShape && snapShape.cNvPrId > 0) exclude.add(snapShape.cNvPrId);
    }
    return collectObstacles(slide.shapes, exclude);
  }, [slide.shapes, draft.otherEndpointCNvPrId, draft.snapped]);
  const path = routeForType(
    draft.connectorType,
    { x: sx, y: sy },
    { x: ex, y: ey },
    startSide,
    endSide,
    obstacles
  );
  const detached = draft.wasAnchored && draft.snapped === null;
  const r = 80_000;
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      viewBox={stageViewBox}
      preserveAspectRatio="xMidYMid meet"
      pointerEvents="none"
      style={{ position: "absolute", inset: 0, width: "100%", height: "100%" }}
    >
      {path}
      {detached ? (
        <circle
          cx={px(ex)}
          cy={px(ey)}
          r={px(r * 0.7)}
          fill="white"
          stroke="#f97316"
          strokeWidth={2}
          strokeDasharray="3 2"
          vectorEffect="non-scaling-stroke"
        />
      ) : null}
      {draft.nearby.map((a, i) => {
        const isSnap =
          draft.snapped !== null &&
          a.shapeId === draft.snapped.shapeId &&
          a.side === draft.snapped.side &&
          Math.abs(a.t - draft.snapped.t) < 1e-3;
        return (
          <circle
            key={`${a.shapeId}-${a.side}-${a.t}-${i}`}
            cx={px(a.x)}
            cy={px(a.y)}
            r={px(isSnap ? r : r * 0.7)}
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
 * Picks the right SVG element for a draft preview given the connector
 * type. Routes through the shared `routeConnector` engine so the
 * dashed preview is byte-for-byte identical to the route that gets
 * committed — including obstacle detours when the caller supplies an
 * `obstacles` list. Stroke style is the same dashed sky-blue across
 * all three types so the user has a consistent visual signal that
 * they're in "drawing" mode.
 */
function routeForType(
  type: ConnectorType,
  sp: { x: number; y: number },
  ep: { x: number; y: number },
  startSide: AnchorSide | null,
  endSide: AnchorSide | null,
  obstacles?: ReadonlyArray<RouterObstacle>
): React.ReactElement {
  const stroke = "#0ea5e9";
  const dash = "6 4";
  const route = routeConnectorShared(type, sp, ep, startSide, endSide, { obstacles });
  if (type === "straight" || (route.kind === "polyline" && route.points.length === 2)) {
    const [a, b] = route.points;
    return (
      <line
        x1={px(a.x)}
        y1={px(a.y)}
        x2={px(b.x)}
        y2={px(b.y)}
        stroke={stroke}
        strokeWidth={2}
        strokeDasharray={dash}
        vectorEffect="non-scaling-stroke"
      />
    );
  }
  if (route.kind === "cubic") {
    const [a, c1, c2, b] = route.points;
    return (
      <path
        d={`M ${px(a.x)} ${px(a.y)} C ${px(c1.x)} ${px(c1.y)} ${px(c2.x)} ${px(c2.y)} ${px(b.x)} ${px(b.y)}`}
        fill="none"
        stroke={stroke}
        strokeWidth={2}
        strokeDasharray={dash}
        vectorEffect="non-scaling-stroke"
      />
    );
  }
  const pts = route.points.map((p) => `${px(p.x)},${px(p.y)}`).join(" ");
  return (
    <polyline
      points={pts}
      fill="none"
      stroke={stroke}
      strokeWidth={2}
      strokeDasharray={dash}
      vectorEffect="non-scaling-stroke"
    />
  );
}

interface ConnectorToolBannerProps {
  readonly type: ConnectorType;
}

/**
 * Tiny "[Tool] Click & drag to draw — Esc to cancel" hint pinned to
 * the top of the canvas while the connector tool is armed. Modeled on
 * Figma's tool-mode strip: low-contrast, never blocks shape clicks,
 * and disappears instantly on commit / Esc / tool toggle.
 */
function ConnectorToolBanner({ type }: ConnectorToolBannerProps): React.ReactElement {
  const label =
    type === "straight" ? "Straight connector" : type === "curved" ? "Curved connector" : "Elbow connector";
  return (
    <div
      data-testid="pptx-connector-tool-banner"
      style={{
        position: "absolute",
        top: 8,
        left: "50%",
        transform: "translateX(-50%)",
        background: "rgba(15,23,42,0.85)",
        color: "white",
        padding: "4px 12px",
        borderRadius: 6,
        fontSize: 12,
        fontFamily: "system-ui, sans-serif",
        pointerEvents: "none",
        boxShadow: "0 2px 8px rgba(0,0,0,0.18)",
        whiteSpace: "nowrap",
      }}
    >
      {label} — click &amp; drag from a shape, port, or empty space. Esc to cancel.
    </div>
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
 * Browser cursors are static glyphs — there's no way to rotate
 * `ns-resize` 30°. So for rotated shapes we pick the standard 8-way
 * cursor whose direction is closest to the rotated handle's screen
 * direction, matching the affordance Figma/PowerPoint use. The
 * effective direction is `handleDir + rotation` (CW positive in
 * both screen coords and our rotation field), snapped to the nearest
 * 45° bucket. Edge handles map to ns/ew, corners to nesw/nwse.
 */
function cursorForRotatedHandle(h: ResizeHandle, rotationDeg: number): string {
  // Logical screen-space angle of each handle direction relative to
  // the shape's centre, measured clockwise from screen-up (north = 0).
  const baseDeg: Record<ResizeHandle, number> = {
    n: 0,
    ne: 45,
    e: 90,
    se: 135,
    s: 180,
    sw: 225,
    w: 270,
    nw: 315,
  };
  const eff = (((baseDeg[h] + rotationDeg) % 360) + 360) % 360;
  // Snap to the nearest of n/ne/e/se/s/sw/w/nw (8-way, 45° steps).
  const bucket = Math.round(eff / 45) % 8;
  // ns / ew / nesw / nwse repeat every 180° — opposite directions
  // share a cursor since there's no asymmetric "north only" glyph.
  switch (bucket) {
    case 0:
    case 4:
      return "ns-resize";
    case 1:
    case 5:
      return "nesw-resize";
    case 2:
    case 6:
      return "ew-resize";
    case 3:
    case 7:
      return "nwse-resize";
    default:
      return "default";
  }
}

/**
 * F4 v2: render a small numbered badge near every shape that carries
 * a typed animation, plus a path overlay for `motionPath` animations.
 *
 * The badge fill colour is keyed by category so the editor reads at a
 * glance which shape gets which kind of effect — green for entrance,
 * amber for emphasis, rose for exit, sky for motion paths. The
 * palette mirrors the right-rail picker in `AnimationsPanel.tsx`.
 *
 * Multiple animations on the same shape stack horizontally so each
 * step is visible and addressable; a shape that holds entrance + spin
 * + exit shows three coloured circles labelled 1/2/3.
 *
 * Motion path overlays are drawn as a dashed stroke from the shape's
 * starting position along the path string (slide-relative coordinates,
 * 1 unit = full slide width / height).
 *
 * Everything has `pointer-events="none"` so it never steals clicks
 * from the underlying shape.
 */
const ANIM_BADGE_PALETTE: Record<string, string> = {
  entrance: "#10b981", // emerald-500
  emphasis: "#f59e0b", // amber-500
  exit: "#f43f5e", // rose-500
  motionPath: "#0ea5e9", // sky-500
};

function animationBadgesSvg(
  slide: Slide,
  slideSize: SlideSize,
  hiddenIds: ReadonlySet<string>
): string {
  if (slide.animations.length === 0) return "";
  const byCNvPrId = new Map<number, Shape>();
  collectShapesByCNvPrId(slide.shapes, byCNvPrId);
  const slideW = slideSize.cxEmu;
  const slideH = slideSize.cyEmu;
  const r = px(90000);
  const stride = r * 2.2;

  const overlays: string[] = [];
  const badgesByShape = new Map<string, number>();

  for (const a of slide.animations) {
    const shape = byCNvPrId.get(a.targetCNvPrId);
    if (!shape) continue;
    if (hiddenIds.has(shape.id)) continue;
    const box = shapeBoundingBox(shape);
    if (!box) continue;

    if (a.category === "motionPath" && a.motionPath) {
      const startCx = px(box.x + box.cx / 2);
      const startCy = px(box.y + box.cy / 2);
      const d = motionPathToSvgD(a.motionPath, startCx, startCy, px(slideW), px(slideH));
      if (d) {
        const stroke = ANIM_BADGE_PALETTE.motionPath;
        overlays.push(
          `<g class="anim-motion" pointer-events="none">`,
          `<path d="${d}" fill="none" stroke="${stroke}" stroke-width="${px(20000)}" stroke-dasharray="${px(60000)} ${px(40000)}" opacity="0.85"/>`,
          `</g>`
        );
      }
    }

    const slot = badgesByShape.get(shape.id) ?? 0;
    badgesByShape.set(shape.id, slot + 1);
    const cx = px(box.x) + r + slot * stride;
    const cy = px(box.y) + r;
    const order = a.order + 1;
    const fill = ANIM_BADGE_PALETTE[a.category] ?? "#facc15";
    overlays.push(
      `<g class="anim-badge anim-badge-${a.category}" pointer-events="none">`,
      `<circle cx="${cx}" cy="${cy}" r="${r}" fill="${fill}" stroke="#1f2937" stroke-width="${px(12000)}"/>`,
      `<text x="${cx}" y="${cy + px(36000)}" text-anchor="middle" font-size="${px(100000)}" font-family="sans-serif" font-weight="700" fill="#ffffff">${order}</text>`,
      `</g>`
    );
  }
  return overlays.join("");
}

/**
 * Convert an OOXML motion-path string (compact `M x y L x y C x1 y1
 * x2 y2 x y E` syntax with slide-relative coordinates) to an SVG path
 * `d` attribute anchored at the shape's centre. Returns `null` when
 * the string is empty or unparseable.
 *
 * OOXML coordinates are deltas from the starting position, so we add
 * each (dx, dy) to (anchorX, anchorY). The `E` token marks the path
 * end and is dropped. Anything we can't tokenise causes an early
 * return so we never paint a partial overlay.
 */
function motionPathToSvgD(
  path: string,
  anchorX: number,
  anchorY: number,
  slideWPx: number,
  slideHPx: number
): string | null {
  const tokens = path.trim().split(/\s+/);
  const out: string[] = [];
  let i = 0;
  while (i < tokens.length) {
    const cmd = tokens[i++];
    if (cmd === "E" || cmd === "Z" || cmd === "z") {
      // OOXML "End" marker; nothing to emit for SVG (we leave the path open).
      continue;
    }
    const consume = (n: number): number[] | null => {
      const xs: number[] = [];
      for (let k = 0; k < n; k++) {
        const t = tokens[i++];
        if (t === undefined) return null;
        const num = Number(t);
        if (!Number.isFinite(num)) return null;
        xs.push(num);
      }
      return xs;
    };
    if (cmd === "M" || cmd === "L") {
      const xs = consume(2);
      if (!xs) return null;
      const x = anchorX + xs[0]! * slideWPx;
      const y = anchorY + xs[1]! * slideHPx;
      out.push(`${cmd}${x.toFixed(2)},${y.toFixed(2)}`);
    } else if (cmd === "C") {
      const xs = consume(6);
      if (!xs) return null;
      const x1 = anchorX + xs[0]! * slideWPx;
      const y1 = anchorY + xs[1]! * slideHPx;
      const x2 = anchorX + xs[2]! * slideWPx;
      const y2 = anchorY + xs[3]! * slideHPx;
      const x = anchorX + xs[4]! * slideWPx;
      const y = anchorY + xs[5]! * slideHPx;
      out.push(
        `C${x1.toFixed(2)},${y1.toFixed(2)} ${x2.toFixed(2)},${y2.toFixed(2)} ${x.toFixed(2)},${y.toFixed(2)}`
      );
    } else {
      // Unknown token — bail rather than risk drawing a partial path.
      return null;
    }
  }
  if (out.length === 0) return null;
  // OOXML paths often start with a relative `M 0 0`, which collapses
  // to "the shape's start point". When the first command is a line or
  // curve (no leading M), prepend an `M anchor` so SVG knows where
  // the path begins.
  if (!/^M/.test(out[0]!)) {
    return `M${anchorX.toFixed(2)},${anchorY.toFixed(2)} ${out.join(" ")}`;
  }
  return out.join(" ");
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
  const stageViewBox = useStageViewBox(slideSize);
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
      viewBox={stageViewBox}
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

interface RemoteSelectionOverlayProps {
  readonly slide: Slide;
  readonly slideSize: SlideSize;
  readonly peers: ReadonlyArray<RemoteSelectionPeer>;
}

/**
 * SVG overlay that paints a per-peer colored outline around every
 * shape a remote peer has selected on the active slide, plus an
 * "X is here" badge for peers viewing this slide without a shape
 * selection. Filters by `slideId` so peers on a different slide are
 * silent here (the slide rail dots cover cross-slide visibility).
 *
 * Drawn behind the local selection chrome (`SelectionOverlaySvg`)
 * so the local user's selection visually wins on overlap.
 */
function RemoteSelectionOverlay({
  slide,
  slideSize,
  peers,
}: RemoteSelectionOverlayProps): React.ReactElement | null {
  const stageViewBox = useStageViewBox(slideSize);
  // Match on the slide's stable OOXML `partPath` (e.g.
  // `ppt/slides/slide3.xml`). Two browsers parsing the same .pptx
  // mint independent local NodeIds for `slide.id`, so the previous
  // `slide.id` comparison silently never matched across peers.
  const onSlide = peers.filter((p) => p.slideId === slide.partPath);
  if (onSlide.length === 0) return null;
  const outlines: React.ReactElement[] = [];
  const idleBadges: { peer: RemoteSelectionPeer; index: number }[] = [];
  for (let i = 0; i < onSlide.length; i++) {
    const peer = onSlide[i]!;
    if (peer.shapeIds.length === 0) {
      idleBadges.push({ peer, index: idleBadges.length });
      continue;
    }
    for (const shapeId of peer.shapeIds) {
      // `shapeId` is the OOXML `cNvPrId` (a small integer encoded as
      // a string) — also stable across peers, unlike the local
      // NodeId.
      const sh = findShapeByCNvPrId(slide.shapes, shapeId);
      if (!sh) continue;
      const box = shapeBoundingBox(sh);
      if (!box) continue;
      const x = px(box.x);
      const y = px(box.y);
      const w = px(box.cx);
      const h = px(box.cy);
      outlines.push(
        <g
          key={`remote-shape-${peer.clientId}-${shapeId}`}
          data-testid="pptx-remote-shape-outline"
          data-peer-color={peer.color}
        >
          <rect
            x={x}
            y={y}
            width={w}
            height={h}
            fill="none"
            stroke={peer.color}
            strokeWidth={2}
            strokeDasharray="6 4"
            strokeOpacity={0.95}
            vectorEffect="non-scaling-stroke"
            pointerEvents="none"
          />
          <foreignObject
            x={x}
            y={Math.max(0, y - px(180000))}
            width={Math.max(60, w)}
            height={px(180000)}
            pointerEvents="none"
          >
            <div
              style={{
                display: "inline-block",
                padding: "1px 6px",
                fontSize: 10,
                fontFamily: "system-ui, sans-serif",
                fontWeight: 500,
                color: "#fff",
                backgroundColor: peer.color,
                borderRadius: "2px 2px 2px 0",
                whiteSpace: "nowrap",
              }}
            >
              {peer.name}
            </div>
          </foreignObject>
        </g>
      );
    }
  }
  return (
    <svg
      xmlns="http://www.w3.org/2000/svg"
      viewBox={stageViewBox}
      preserveAspectRatio="xMidYMid meet"
      pointerEvents="none"
      style={{ position: "absolute", inset: 0, width: "100%", height: "100%" }}
      data-testid="pptx-remote-selection-layer"
    >
      {outlines}
      {idleBadges.length > 0 ? (
        <foreignObject
          x={px(slideSize.cxEmu) - px(2400000)}
          y={px(120000)}
          width={px(2400000)}
          height={px(180000) * idleBadges.length}
          pointerEvents="none"
        >
          <div
            style={{
              display: "flex",
              flexDirection: "column",
              alignItems: "flex-end",
              gap: 4,
              fontFamily: "system-ui, sans-serif",
            }}
          >
            {idleBadges.map(({ peer }) => (
              <div
                key={`remote-idle-${peer.clientId}`}
                style={{
                  padding: "1px 6px",
                  fontSize: 10,
                  fontWeight: 500,
                  color: "#fff",
                  backgroundColor: peer.color,
                  borderRadius: 4,
                  whiteSpace: "nowrap",
                }}
              >
                {peer.name} is here
              </div>
            ))}
          </div>
        </foreignObject>
      ) : null}
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
  const stageViewBox = useStageViewBox(slideSize);
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
      viewBox={stageViewBox}
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
  layout: SlideLayout | undefined,
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
  onTextSelectionChange: ((sel: PptxTextSelection | null) => void) | undefined,
  commitRef: React.MutableRefObject<(() => void) | null>
): React.ReactElement | null {
  const shape = findShape(slide.shapes, shapeId);
  if (!shape || shape.kind !== "text") return null;
  // Fall back to the layout placeholder's geometry when the shape
  // itself doesn't carry an `<a:xfrm>` (PowerPoint placeholders that
  // inherit from the layout). Without this, clicking a freshly-
  // inserted layout-only placeholder silently refused to open the
  // overlay because `shapeBoundingBox` returned null.
  const box = resolvedShapeBoundingBox(shape, layout);
  if (!box) return null;
  return (
    <TextEditOverlay
      key={shapeId}
      shape={shape as TextShape}
      box={box}
      slideSize={slideSize}
      dpi={dpi}
      layout={layout}
      onCommit={(paragraphs, plain) => onCommit(shape as TextShape, paragraphs, plain)}
      onSelectionChange={onTextSelectionChange}
      commitRef={commitRef}
    />
  );
}

interface TextEditOverlayProps {
  readonly shape: TextShape;
  readonly box: BoundingBox;
  readonly slideSize: SlideSize;
  readonly dpi: number;
  /** Slide layout for placeholder geometry / typography inheritance. */
  readonly layout: SlideLayout | undefined;
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
  /**
   * Mutable handle the overlay populates with a `commit()` function
   * the canvas can drive synchronously from `onPointerDown`. Lets a
   * click outside the editable update `editingId` BEFORE the parent
   * selection mutates, eliminating the cross-frame "two states open
   * at once" flash.
   */
  readonly commitRef?: React.MutableRefObject<(() => void) | null>;
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
  layout,
  onCommit,
  onSelectionChange,
  commitRef,
}: TextEditOverlayProps): React.ReactElement {
  // Resolve the placeholder defaults once per shape change so the
  // overlay's typography matches the SVG hint (and the master/layout
  // inheritance chain). Without this, an empty title placeholder
  // opens at Calibri 18pt while the SVG hint paints it at 36pt — the
  // user sees a one-frame "font shrink" on click.
  const defaults = React.useMemo(() => resolvePlaceholderTextDefaults(shape, layout), [shape, layout]);
  const ref = React.useRef<HTMLDivElement>(null);
  const initialPlain = React.useMemo(() => textShapePlain(shape), [shape]);

  // The SVG renders at 1 SVG-unit = 1 CSS px at 96 DPI, then the
  // browser scales the SVG to fit its container. Our HTML overlay
  // lives in the same container but receives raw CSS px, so we have
  // to multiply font sizes / insets by the same scale factor or the
  // overlay text appears noticeably larger than the underlying
  // rendered shape (and the user perceives it as "the font changed
  // when I clicked to edit").
  //
  // We also need the slide-card's pixel offset within the parent
  // stage — the stage now fills the full viewport, so positioning the
  // overlay with simple percentages of the parent would put it in the
  // scratch margin instead of on the slide. `useLayoutEffect` (vs
  // `useEffect`) commits the measurement BEFORE the first paint so
  // the overlay never flashes at the wrong scale or offset.
  const [metrics, setMetrics] = React.useState<{
    readonly scale: number;
    readonly slideLeft: number;
    readonly slideTop: number;
  }>({ scale: 1, slideLeft: 0, slideTop: 0 });
  const containerRef = React.useRef<HTMLDivElement | null>(null);
  React.useLayoutEffect(() => {
    const node = ref.current;
    if (!node) return;
    const parent = node.parentElement as HTMLDivElement | null;
    containerRef.current = parent;
    if (!parent) return;
    const slideCard = parent.querySelector<HTMLDivElement>("[data-testid='pptx-slide-card']");
    const target = slideCard ?? parent;
    const slideNativeWidth = slideSize.cxEmu / EMU_PER_PX_AT_96DPI;
    const measure = () => {
      const cardRect = target.getBoundingClientRect();
      const parentRect = parent.getBoundingClientRect();
      const w = cardRect.width;
      return {
        scale: w > 0 && slideNativeWidth > 0 ? w / slideNativeWidth : 1,
        slideLeft: cardRect.left - parentRect.left,
        slideTop: cardRect.top - parentRect.top,
      };
    };
    setMetrics(measure());
    if (typeof ResizeObserver === "undefined") return;
    const ro = new ResizeObserver(() => setMetrics(measure()));
    ro.observe(target);
    if (slideCard && slideCard !== parent) ro.observe(parent);
    return () => ro.disconnect();
  }, [slideSize.cxEmu]);
  const scale = metrics.scale;

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

  // Position the overlay in absolute pixels relative to the stage
  // div: stage now fills the entire viewport, so percentage-of-parent
  // would drift into the scratch margin. The slide card's offset +
  // EMU→px scale uniquely places the overlay on top of the rendered
  // shape, no matter how big the surrounding canvas grows.
  const overlayLeft = metrics.slideLeft + (box.x / EMU_PER_PX_AT_96DPI) * scale;
  const overlayTop = metrics.slideTop + (box.y / EMU_PER_PX_AT_96DPI) * scale;
  const overlayWidth = (box.cx / EMU_PER_PX_AT_96DPI) * scale;
  const overlayHeight = (box.cy / EMU_PER_PX_AT_96DPI) * scale;
  const insetsEmu = readBodyInsetsFromShape(shape);
  const baseFontPx = (defaults.fontSizePt * dpi) / 72;
  // Pick the same default font the SVG renderer uses so the text
  // doesn't visibly switch to the system sans-serif when the caret
  // enters the shape. The placeholder-defaults resolver walks
  // shape → layout → master/theme so an empty title placeholder
  // opens at the right family/size instead of the Calibri-18pt
  // fallback that used to make the prompt visibly resize on click.
  const baseFontFamily = defaults.fontFamily;
  const justifyContent =
    defaults.anchor === "middle" ? "center" : defaults.anchor === "bottom" ? "flex-end" : "flex-start";
  const padTop = (insetsEmu.t / EMU_PER_PX_AT_96DPI) * scale;
  const padRight = (insetsEmu.r / EMU_PER_PX_AT_96DPI) * scale;
  const padBottom = (insetsEmu.b / EMU_PER_PX_AT_96DPI) * scale;
  const padLeft = (insetsEmu.l / EMU_PER_PX_AT_96DPI) * scale;

  // PowerPoint shows a ghost prompt ("Click to add title", etc.) for
  // empty placeholder shapes that vanishes on the first keystroke.
  // The SVG renderer paints this hint when the overlay isn't open;
  // while editing, we reproduce it as a contentEditable=false sibling
  // node so the user sees the same prompt instead of a blank box.
  const isEffectivelyEmpty = React.useMemo(() => initialPlain.trim().length === 0, [initialPlain]);
  const [hasInput, setHasInput] = React.useState(false);
  const placeholderType = shape.placeholder?.type;
  const showPrompt = isEffectivelyEmpty && !hasInput && !!placeholderType;
  const promptText = placeholderType ? placeholderPromptLabel(placeholderType) : "";

  // `commitEdit` is the single canonical exit path for the overlay.
  // It's also exposed via `commitRef` so the canvas can drive a
  // synchronous commit from `onPointerDown` BEFORE selection mutates,
  // eliminating the cross-frame "B selected + A still editing" flash.
  // Guarded by a ref so a pointerdown-driven commit followed by the
  // browser's natural blur doesn't fire `onCommit` twice.
  const committedRef = React.useRef(false);
  const commitEdit = React.useCallback(() => {
    if (committedRef.current) return;
    committedRef.current = true;
    const node = ref.current;
    const paragraphs = node ? extractParagraphsFromOverlay(node) : [];
    const plain = node?.innerText ?? "";
    onCommit(paragraphs, plain);
    onSelectionChange?.(null);
  }, [onCommit, onSelectionChange]);

  React.useEffect(() => {
    if (!commitRef) return;
    commitRef.current = commitEdit;
    return () => {
      if (commitRef.current === commitEdit) commitRef.current = null;
    };
  }, [commitRef, commitEdit]);

  return (
    <div
      ref={ref}
      contentEditable
      suppressContentEditableWarning
      data-testid="pptx-text-overlay"
      style={{
        position: "absolute",
        left: `${overlayLeft}px`,
        top: `${overlayTop}px`,
        width: `${overlayWidth}px`,
        height: `${overlayHeight}px`,
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
      onInput={() => {
        if (!hasInput) setHasInput(true);
      }}
      onBlur={(e) => {
        // Don't commit when focus moves to a sibling we explicitly
        // marked as "keep editing focus" (the format toolbar). The
        // toolbar suppresses mousedown so the relatedTarget stays
        // null in most browsers; we still bail out if the user
        // clicked a button that opted in via data-pptx-keep-edit.
        const next = e.relatedTarget as HTMLElement | null;
        if (next?.closest?.("[data-pptx-keep-edit]")) return;
        commitEdit();
      }}
      onKeyDown={(e) => {
        if (e.key === "Enter" && (e.metaKey || e.ctrlKey)) {
          e.preventDefault();
          ref.current?.blur();
        } else if (e.key === "Escape") {
          // Esc reverts text and exits edit mode but keeps the shape
          // selected — the parent's `setSelectedIds` is untouched, so
          // when blur fires the shape stays in the selection set.
          e.preventDefault();
          if (ref.current) ref.current.innerText = initialPlain;
          ref.current?.blur();
        }
      }}
    >
      {showPrompt ? (
        <div
          contentEditable={false}
          aria-hidden
          data-testid="pptx-text-overlay-prompt"
          style={{
            position: "absolute",
            inset: 0,
            paddingTop: padTop,
            paddingRight: padRight,
            paddingBottom: padBottom,
            paddingLeft: padLeft,
            display: "flex",
            flexDirection: "column",
            justifyContent,
            color: "#9ca3af",
            pointerEvents: "none",
            userSelect: "none",
          }}
        >
          {promptText}
        </div>
      ) : null}
      {shape.txBody.paragraphs.map((p, pi) => (
        <div key={pi} data-paragraph={pi} style={paragraphStyle(p)}>
          {paragraphToReact(p, pi, scale, dpi)}
        </div>
      ))}
    </div>
  );
}

/**
 * Mirror of the SVG renderer's `placeholderHintLabel` so the HTML
 * overlay shows the same ghost prompt text PowerPoint does. Kept in
 * sync with `packages/pptx/src/renderer/svg/shapes.ts`.
 */
function placeholderPromptLabel(type: string): string {
  switch (type) {
    case "title":
    case "ctrTitle":
      return "Click to add title";
    case "subTitle":
      return "Click to add subtitle";
    case "body":
      return "Click to add text";
    case "ftr":
      return "Footer";
    case "hdr":
      return "Header";
    case "dt":
      return "Date";
    case "sldNum":
      return "Slide number";
    default:
      return "Click to add text";
  }
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
