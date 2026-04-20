import type {
  ChartPart,
  ChartShape,
  ConnectorShape,
  GroupShape,
  OpaqueShape,
  OpaqueXml,
  Picture,
  Shape,
  SlideSize,
  TableShape,
  TextParagraph,
  TextRun,
  TextShape,
} from "../../model/types.js";
import { resolveEndpoint } from "../../model/connector-geometry.js";
import { DEFAULT_THEME, type ThemeColorScheme } from "../layout/color.js";
import { shapeBoundingBox } from "../layout/shape.js";
import { SVG_UNIT_PER_EMU } from "../layout/slide.js";
import { escXml } from "./escape.js";
import {
  curvedPathD as curvedPathDShared,
  routeConnector as routeConnectorShared,
  type RouterObstacle,
} from "../connector-router/index.js";

/**
 * Convert an EMU value to the renderer's user-unit space (1 unit ≈ 1px @ 96 DPI).
 *
 * The SVG `viewBox` is emitted in pixel-equivalent units (see `slideViewBox`)
 * so every coordinate, size, font-size, and stroke-width inside the SVG
 * document must go through this helper. Earlier versions kept everything in
 * EMU and relied on a `transform="scale(…)"` wrapper, but Chrome quietly
 * degrades text rendering when `font-size` is in the 5–6-digit EMU range
 * even under such a wrapper. Emitting pre-scaled values avoids that gotcha
 * and produces SVG that round-trips cleanly through other tools.
 */
function u(emu: number): number {
  // Two decimals is plenty: 0.01 user units ≈ 0.01 px ≈ 0.0001 inch.
  return Math.round(emu * SVG_UNIT_PER_EMU * 100) / 100;
}

export interface SvgRenderCtx {
  readonly slideSize: SlideSize;
  readonly theme?: ThemeColorScheme;
  /** Map from media partPath → URL (object URL or data URL). */
  readonly mediaUrls?: ReadonlyMap<string, string>;
  /** F3: typed chart parts keyed by part path, used by chart renderer. */
  readonly charts?: ReadonlyMap<string, ChartPart>;
  /**
   * Map from `cNvPrId` → resolved shape (groups walked). Populated by the
   * caller and consumed by the connector renderer to look up anchored
   * endpoints. Optional so existing callers that only render a single
   * shape (e.g. tests) keep working — anchored connectors then fall
   * back to their stored bounding box corners.
   */
  readonly shapesByCNvPrId?: ReadonlyMap<number, Shape>;
  /**
   * Inflated obstacle boxes (other shapes on the slide) the connector
   * router should avoid. Populated by `slideToSvgString` from the
   * slide's full shape tree. Optional: when omitted the router falls
   * back to its heuristic-only path, which still produces sensible
   * routes — only the obstacle-avoiding A* fallback is unavailable.
   *
   * Per-connector exemption (the connector's own anchored endpoints'
   * shapes) is handled at the call site; obstacles passed here should
   * already EXCLUDE every shape the connector touches.
   */
  readonly connectorObstacles?: ReadonlyArray<RouterObstacle>;
  /**
   * When true (the default), text shapes that are empty *and* carry a
   * `placeholder` field render a dashed outline + ghost prompt label
   * (e.g. "Click to add title", or an image icon for `pic`-typed
   * placeholders) so a freshly-inserted layout slide doesn't read as a
   * blank canvas. The hint UI lives entirely in the renderer — it is
   * never serialised back into the saved `.pptx`. Pass `false` to
   * suppress (e.g. when rendering an export-style preview).
   */
  readonly renderPlaceholderHints?: boolean;
}

export function shapeToSvg(shape: Shape, ctx: SvgRenderCtx): string {
  switch (shape.kind) {
    case "text":
      return textShapeToSvg(shape, ctx);
    case "pic":
      return pictureToSvg(shape, ctx);
    case "group":
      return groupShapeToSvg(shape, ctx);
    case "table":
      return tableToSvg(shape, ctx);
    case "chart":
      return chartToSvg(shape, ctx);
    case "connector":
      return connectorToSvg(shape, ctx);
    case "opaque":
      return opaqueShapeToSvg(shape);
  }
}

// ─── Connector renderer ───────────────────────────────────────────────────

/**
 * Render a `ConnectorShape` as native SVG. Resolves anchored endpoints
 * via the slide's `shapesByCNvPrId` map (passed through `ctx`), or falls
 * back to the connector's stored bounding-box corners for unresolved
 * endpoints. We render straight lines as `<line>`, elbow connectors as a
 * 3-segment `<polyline>` with a midpoint pivot, and curved connectors as
 * a quadratic Bezier through the bounding box centre. Arrowheads are
 * declared once per slide by `slideToSvgString` via `<defs>`; here we
 * only reference them through `marker-start` / `marker-end`.
 */
function connectorToSvg(shape: ConnectorShape, ctx: SvgRenderCtx): string {
  const map = ctx.shapesByCNvPrId ?? new Map<number, Shape>();
  const startPt = resolveEndpoint(shape.start, map);
  const endPt = resolveEndpoint(shape.end, map);
  const fallbackStart = {
    x: shape.position?.xEmu ?? 0,
    y: shape.position?.yEmu ?? 0,
  };
  const fallbackEnd = {
    x: (shape.position?.xEmu ?? 0) + (shape.size?.cxEmu ?? 0),
    y: (shape.position?.yEmu ?? 0) + (shape.size?.cyEmu ?? 0),
  };
  const sp = startPt ?? fallbackStart;
  const ep = endPt ?? fallbackEnd;
  const stroke = shape.stroke?.color ?? "374151";
  const widthEmu = shape.stroke?.widthEmu && shape.stroke.widthEmu > 0 ? shape.stroke.widthEmu : 9525; // ≈ 0.75pt
  const dashAttr = dashArrayAttr(shape.stroke?.dash, widthEmu);
  const strokeAttrs = ` stroke="#${stroke}" stroke-width="${u(widthEmu)}" fill="none" stroke-linecap="round" stroke-linejoin="round"${dashAttr}`;
  const headAttr = endShapeMarkerAttr("end", shape.headEnd);
  const tailAttr = endShapeMarkerAttr("start", shape.tailEnd);
  // A separate transparent stroke painted UNDER the visible one so the
  // hit area is always at least ~14 px wide regardless of the line's
  // actual stroke width. Without this a 1px line is essentially un-
  // clickable — the user has to land exactly on the pixel-thin path,
  // which doesn't match how PowerPoint/Slides feel. We deliberately
  // omit arrowheads + dashes here since this layer is invisible and
  // exists only for hit detection.
  const HIT_STROKE_EMU = 130_000; // ≈ 13.6 px @ 96 DPI
  const hitWidthEmu = Math.max(HIT_STROKE_EMU, widthEmu * 4);
  // `pointer-events="stroke"` is required because the default
  // `visiblePainted` excludes strokes painted with opacity 0; without
  // this attribute the wide hit-band wouldn't actually catch clicks.
  const hitAttrs = ` stroke="#000" stroke-opacity="0" stroke-width="${u(hitWidthEmu)}" fill="none" stroke-linecap="round" stroke-linejoin="round" pointer-events="stroke"`;
  // Delegate routing to the shared engine. Both the SVG renderer and
  // the React chrome call into the same `routeConnector` so the
  // preview the user sees while drawing matches what gets committed.
  // We forward obstacles minus this connector's own anchored target
  // shapes — otherwise the router would try to detour around the very
  // shapes it's anchored to.
  const startSide = shape.start.kind === "anchored" ? shape.start.side : null;
  const endSide = shape.end.kind === "anchored" ? shape.end.side : null;
  const obstacles = filterConnectorObstacles(ctx.connectorObstacles, shape, map);
  const route = routeConnectorShared(shape.connectorType, sp, ep, startSide, endSide, {
    waypoints: shape.waypoints,
    obstacles,
  });
  let visibleSvg: string;
  let hitSvg: string;
  if (route.kind === "cubic") {
    const d = curvedPathDShared(route.points, u);
    hitSvg = `<path d="${d}"${hitAttrs}/>`;
    visibleSvg = `<path d="${d}"${strokeAttrs}${headAttr}${tailAttr}/>`;
  } else if (route.points.length === 2) {
    const [p0, p1] = route.points;
    const lineCoords = `x1="${u(p0.x)}" y1="${u(p0.y)}" x2="${u(p1.x)}" y2="${u(p1.y)}"`;
    hitSvg = `<line ${lineCoords}${hitAttrs}/>`;
    visibleSvg = `<line ${lineCoords}${strokeAttrs}${headAttr}${tailAttr}/>`;
  } else {
    // Visible path uses rounded corners — sharp 90° joins on a thin
    // stroke read as harsh ticks at typical zoom levels. The hit area
    // can stay as a polyline (it's invisible, and the polyline catches
    // anything inside its wider stroke envelope anyway).
    const visD = roundedPolylinePath(route.points, ELBOW_CORNER_RADIUS_EMU);
    const pts = route.points.map((p) => `${u(p.x)},${u(p.y)}`).join(" ");
    hitSvg = `<polyline points="${pts}"${hitAttrs}/>`;
    visibleSvg = `<path d="${visD}"${strokeAttrs}${headAttr}${tailAttr}/>`;
  }
  return `${groupOpen("connector", shape.id)}${hitSvg}${visibleSvg}${groupClose()}`;
}

/**
 * Drop the connector's own anchored target shapes from the obstacle
 * list so the router doesn't try to route around the very things
 * it's anchored to. Free endpoints contribute no exemption.
 *
 * The slide-wide obstacle list lives on `ctx.connectorObstacles` and
 * was assembled before we knew which connector we're rendering — we
 * couldn't bake per-connector exemptions into it then. We do it here
 * by mapping each anchored endpoint's `cNvPrId` back to its shape
 * (via the same `shapesByCNvPrId` map used for endpoint resolution)
 * and dropping any obstacle whose `id` matches a target shape's `id`.
 */
function filterConnectorObstacles(
  obstacles: ReadonlyArray<RouterObstacle> | undefined,
  shape: ConnectorShape,
  shapesByCNvPrId: ReadonlyMap<number, Shape>
): ReadonlyArray<RouterObstacle> {
  if (!obstacles || obstacles.length === 0) return [];
  const exemptShapeIds = new Set<string>();
  for (const ep of [shape.start, shape.end]) {
    if (ep.kind !== "anchored") continue;
    const target = shapesByCNvPrId.get(ep.targetCNvPrId);
    if (target) exemptShapeIds.add(target.id);
  }
  if (exemptShapeIds.size === 0) return obstacles;
  return obstacles.filter((o) => !exemptShapeIds.has(o.id));
}

/**
 * Corner-rounding radius for the elbow connector's visible path. Small
 * enough that diagonals through tight bridge segments still read as a
 * Manhattan route, large enough that 90° corners don't look like sharp
 * ticks on a thin stroke. Auto-clamped to half the shorter incident
 * segment by `roundedPolylinePath` so it never overshoots.
 */
const ELBOW_CORNER_RADIUS_EMU = 60_000; // ≈ 6.3 px @ 96 DPI

/**
 * Convert a polyline (≥ 2 points) into an SVG path string with each
 * interior corner replaced by a quadratic round of `r` EMU. The path
 * still passes through the start and end points exactly, so anchored
 * endpoints / arrowheads remain pixel-aligned with the model.
 */
function roundedPolylinePath(
  pts: ReadonlyArray<{ readonly x: number; readonly y: number }>,
  r: number
): string {
  if (pts.length === 0) return "";
  if (pts.length === 1) return `M ${u(pts[0].x)} ${u(pts[0].y)}`;
  if (pts.length === 2) {
    return `M ${u(pts[0].x)} ${u(pts[0].y)} L ${u(pts[1].x)} ${u(pts[1].y)}`;
  }
  const out: string[] = [`M ${u(pts[0].x)} ${u(pts[0].y)}`];
  for (let i = 1; i < pts.length - 1; i++) {
    const prev = pts[i - 1];
    const curr = pts[i];
    const next = pts[i + 1];
    const dIn = Math.hypot(curr.x - prev.x, curr.y - prev.y);
    const dOut = Math.hypot(next.x - curr.x, next.y - curr.y);
    const radius = Math.min(r, dIn / 2, dOut / 2);
    if (!(radius > 0) || dIn === 0 || dOut === 0) {
      out.push(`L ${u(curr.x)} ${u(curr.y)}`);
      continue;
    }
    const inT = radius / dIn;
    const inX = curr.x - (curr.x - prev.x) * inT;
    const inY = curr.y - (curr.y - prev.y) * inT;
    const outT = radius / dOut;
    const outX = curr.x + (next.x - curr.x) * outT;
    const outY = curr.y + (next.y - curr.y) * outT;
    out.push(`L ${u(inX)} ${u(inY)}`);
    out.push(`Q ${u(curr.x)} ${u(curr.y)} ${u(outX)} ${u(outY)}`);
  }
  const last = pts[pts.length - 1];
  out.push(`L ${u(last.x)} ${u(last.y)}`);
  return out.join(" ");
}

/**
 * Map a `ConnectorEndShape` to the marker URL the renderer should
 * reference. We always emit ALL four markers via `slide.ts` (`cxn-
 * arrow`, `cxn-triangle`, `cxn-oval`, `cxn-none`); this helper just
 * picks the right one. Returns an empty string for `none` /
 * `undefined` so the SVG attribute is omitted entirely (omitting is
 * cheaper than referencing the no-op marker for the common case).
 */
function endShapeMarkerAttr(
  position: "start" | "end",
  endShape: "none" | "arrow" | "triangle" | "oval" | undefined
): string {
  if (!endShape || endShape === "none") return "";
  const attr = position === "end" ? "marker-end" : "marker-start";
  switch (endShape) {
    case "arrow":
      return ` ${attr}="url(#cxn-arrow)"`;
    case "triangle":
      return ` ${attr}="url(#cxn-triangle)"`;
    case "oval":
      return ` ${attr}="url(#cxn-oval)"`;
    default: {
      const _exhaustive: never = endShape;
      void _exhaustive;
      return "";
    }
  }
}

function dashArrayAttr(
  dash: "solid" | "dashed" | "dotted" | "longDash" | "dashDot" | undefined,
  widthEmu: number
): string {
  if (!dash || dash === "solid") return "";
  // Patterns scale with stroke width so the dashes feel consistent
  // when users bump the weight. The values mimic PowerPoint's
  // `prstDash` presets at the default ~1pt width.
  const w = u(widthEmu);
  switch (dash) {
    case "dashed":
      return ` stroke-dasharray="${(w * 4).toFixed(2)} ${(w * 3).toFixed(2)}"`;
    case "longDash":
      return ` stroke-dasharray="${(w * 8).toFixed(2)} ${(w * 3).toFixed(2)}"`;
    case "dashDot":
      return ` stroke-dasharray="${(w * 4).toFixed(2)} ${(w * 3).toFixed(2)} ${(w * 1).toFixed(2)} ${(w * 3).toFixed(2)}"`;
    case "dotted":
      return ` stroke-dasharray="${(w * 1).toFixed(2)} ${(w * 2).toFixed(2)}"`;
    default: {
      const _exhaustive: never = dash;
      void _exhaustive;
      return "";
    }
  }
}

function textShapeToSvg(shape: TextShape, ctx: SvgRenderCtx): string {
  const box = shapeBoundingBox(shape);
  if (!box) return groupOpen("text", shape.id) + groupClose();
  const theme = ctx.theme ?? DEFAULT_THEME;

  const fill = readFillFromOpaque(shape.spPrTail, theme);
  const stroke = readStrokeFromOpaque(shape.spPrTail, theme);
  const prst = readPrstGeom(shape.spPrTail);

  if (prst === "line" && stroke) {
    return [
      groupOpen("text", shape.id, { transform: `translate(${u(box.x)} ${u(box.y)})` }),
      `<line x1="0" y1="0" x2="${u(box.cx)}" y2="${u(box.cy)}" stroke="#${stroke.color}" stroke-width="${u(stroke.widthEmu)}" stroke-linecap="round"/>`,
      groupClose(),
    ].join("");
  }

  const rectFill = fill ? `#${fill}` : "transparent";
  const strokeAttrs = stroke ? ` stroke="#${stroke.color}" stroke-width="${u(stroke.widthEmu)}"` : "";

  const hasText = shape.txBody.paragraphs.some((p) =>
    p.runs.some((r) => !r.isLineBreak && r.text.length > 0)
  );

  // Draw the actual geometry (ellipse, triangle, …) so `pptx:add-shape`
  // produces a visually distinct shape rather than an undifferentiated
  // rectangle. Falls back to a rect for `prst="rect"` and for any preset
  // we don't yet model — the bounding box is preserved either way.
  const geometry = renderGeometry(prst, box.cx, box.cy, rectFill, strokeAttrs);

  const out = [groupOpen("text", shape.id, { transform: `translate(${u(box.x)} ${u(box.y)})` }), geometry];
  if (hasText) {
    out.push(renderWrappedTextHtml(shape, box.cx, box.cy, fill, theme));
  } else if (shape.placeholder && (ctx.renderPlaceholderHints ?? true)) {
    // Empty layout placeholder. Without this hint the shape renders as
    // a transparent rect with no fill or stroke — i.e. invisible — so
    // a freshly-inserted "Title and Content" slide looks blank to the
    // user. We paint a dashed outline + ghost label (and an image icon
    // for `pic` placeholders) so the slot reads as "click here to fill
    // me", matching PowerPoint's authoring affordance. The hint UI is
    // emitted with `pointer-events="none"` so the underlying shape's
    // hit-testing (selection, drag) keeps working unchanged. Nothing
    // here gets serialised back into the saved .pptx — placeholder
    // prompts live entirely in the renderer.
    out.push(renderPlaceholderHint(shape.placeholder, box.cx, box.cy));
  }
  out.push(groupClose());
  return out.join("");
}

/**
 * Renderer-only ghost UI for an empty placeholder. Produces a dashed
 * outline that fills the placeholder's bounding box plus a label (and
 * an image-icon glyph for `pic` placeholders) styled to roughly match
 * how the user's first run of text would appear once they start typing
 * — title placeholders render large, body placeholders render smaller
 * and left-aligned, ctrTitle stays centered, etc. Without this match
 * the hints all read as "generic centered prompt" which makes the
 * eventual typed text feel like it jumps around when it appears.
 *
 * The wrapper `<g>` carries `pointer-events="none"` so clicks fall
 * through to the underlying `data-shape-id` group — the user can still
 * select / drag / resize the placeholder, and double-click still
 * enters edit mode.
 */
function renderPlaceholderHint(
  placeholder: { type?: string; idx?: number },
  cxEmu: number,
  cyEmu: number
): string {
  const w = u(cxEmu);
  const h = u(cyEmu);
  const type = placeholder.type ?? "body";
  const isPic = type === "pic";
  const labelColor = "#9ca3af"; // tailwind gray-400
  const outlineColor = "#cbd5e1"; // tailwind slate-300
  const dash = `${u(60_000)},${u(40_000)}`;
  const style = placeholderHintStyle(type);
  // Convert PowerPoint point sizes to user units. Min cap so micro
  // placeholders (footer, slide#) still surface a readable label.
  const ptToUserUnits = u(12_700); // 1pt = 12_700 EMU
  const baseFontSize = Math.max(10, style.fontPt * ptToUserUnits);
  // Auto-shrink for tiny placeholder boxes so the hint never overflows
  // — preserves the "what you'd see typed" feel without breaking the
  // dashed bounding box.
  const fitFontSize = Math.min(baseFontSize, h * 0.55, (w / Math.max(1, style.label?.length ?? 1)) * 1.6);
  const label = style.label ?? placeholderHintLabel(type);
  // Inset the label so it doesn't kiss the dashed outline (real
  // placeholders carry `lIns`/`tIns` defaults of ~0.1" / 0.05").
  const padX = u(91_440); // 0.1"
  const padY = u(45_720); // 0.05"
  const parts: string[] = [];
  parts.push(
    `<g class="placeholder-hint" pointer-events="none">`,
    `<rect x="0" y="0" width="${w}" height="${h}" fill="none" stroke="${outlineColor}" stroke-width="1.5" stroke-dasharray="${dash}" vector-effect="non-scaling-stroke"/>`
  );
  if (isPic) {
    // A small mountain-and-sun glyph (PowerPoint's "Insert Picture"
    // affordance uses the same metaphor) sized relative to the box, so
    // it stays legible for both half-slide hero images and tiny inline
    // thumbnails. Stacks above the label.
    const iconSize = Math.min(w, h) * 0.28;
    const iconX = (w - iconSize) / 2;
    const iconY = (h - iconSize) / 2 - fitFontSize * 1.4;
    parts.push(renderPictureIcon(iconX, iconY, iconSize, labelColor));
    parts.push(
      `<text x="${w / 2}" y="${h / 2 + iconSize / 2 + fitFontSize * 0.6}" text-anchor="middle" dominant-baseline="middle" font-family="${style.fontFamily}" font-weight="${style.fontWeight}" font-size="${fitFontSize}" fill="${labelColor}">${escXml(label)}</text>`
    );
  } else {
    const x = horizontalAnchorX(style.align, w, padX);
    const y = verticalAnchorY(style.anchor, h, padY, fitFontSize);
    parts.push(
      `<text x="${x}" y="${y}" text-anchor="${svgTextAnchor(style.align)}" dominant-baseline="${svgDominantBaseline(style.anchor)}" font-family="${style.fontFamily}" font-weight="${style.fontWeight}" font-size="${fitFontSize}" fill="${labelColor}">${escXml(label)}</text>`
    );
  }
  parts.push(`</g>`);
  return parts.join("");
}

/**
 * Per-placeholder-type rendering defaults that approximate the
 * built-in PowerPoint placeholder styles. Sizes / alignments are the
 * ones the user would see if they started typing into an empty
 * placeholder of the given type, so the ghost prompt previews the
 * eventual layout instead of always centring everything.
 */
interface PlaceholderHintStyle {
  readonly fontPt: number;
  readonly fontFamily: string;
  readonly fontWeight: number;
  readonly align: "left" | "center" | "right";
  readonly anchor: "top" | "middle" | "bottom";
  readonly label?: string;
}

function placeholderHintStyle(type: string): PlaceholderHintStyle {
  // PowerPoint defaults — Calibri across the board so the hint
  // matches what a freshly-typed run would render as. Colour stays
  // gray-400 (set at the call site) regardless of type.
  const FAMILY = "Calibri, 'Segoe UI', sans-serif";
  switch (type) {
    case "ctrTitle":
      // Title-Slide layout: large, fully centered both axes.
      return { fontPt: 40, fontFamily: FAMILY, fontWeight: 400, align: "center", anchor: "middle" };
    case "title":
      // Title bar at the top of content layouts: large, left-aligned,
      // bottom-anchored so the baseline sits where typed text would.
      return { fontPt: 36, fontFamily: FAMILY, fontWeight: 400, align: "left", anchor: "middle" };
    case "subTitle":
      // Subtitle under ctrTitle: medium, centered, top-anchored.
      return { fontPt: 24, fontFamily: FAMILY, fontWeight: 400, align: "center", anchor: "middle" };
    case "body":
      // Body / content: smaller, left-aligned, top-anchored — matches
      // a bullet at outline level 1.
      return { fontPt: 18, fontFamily: FAMILY, fontWeight: 400, align: "left", anchor: "top" };
    case "ftr":
      return {
        fontPt: 12,
        fontFamily: FAMILY,
        fontWeight: 400,
        align: "center",
        anchor: "middle",
        label: "Footer",
      };
    case "hdr":
      return {
        fontPt: 12,
        fontFamily: FAMILY,
        fontWeight: 400,
        align: "center",
        anchor: "middle",
        label: "Header",
      };
    case "dt":
      return {
        fontPt: 12,
        fontFamily: FAMILY,
        fontWeight: 400,
        align: "left",
        anchor: "middle",
        label: "Date",
      };
    case "sldNum":
      return {
        fontPt: 12,
        fontFamily: FAMILY,
        fontWeight: 400,
        align: "right",
        anchor: "middle",
        label: "‹#›",
      };
    case "chart":
    case "tbl":
    case "dgm":
    case "media":
    case "pic":
      // Asset placeholders: the icon dominates, hint sits centered as a
      // caption below it. Rendered specially in `renderPlaceholderHint`.
      return { fontPt: 14, fontFamily: FAMILY, fontWeight: 400, align: "center", anchor: "middle" };
    default:
      return { fontPt: 18, fontFamily: FAMILY, fontWeight: 400, align: "left", anchor: "top" };
  }
}

function horizontalAnchorX(align: "left" | "center" | "right", w: number, padX: number): number {
  switch (align) {
    case "left":
      return padX;
    case "center":
      return w / 2;
    case "right":
      return w - padX;
  }
}

function verticalAnchorY(
  anchor: "top" | "middle" | "bottom",
  h: number,
  padY: number,
  fontSize: number
): number {
  switch (anchor) {
    case "top":
      // Baseline near the top — push down by ~80% of the cap height
      // so the glyph tops sit visually flush with the top inset.
      return padY + fontSize * 0.8;
    case "middle":
      return h / 2;
    case "bottom":
      return h - padY;
  }
}

function svgTextAnchor(align: "left" | "center" | "right"): string {
  switch (align) {
    case "left":
      return "start";
    case "center":
      return "middle";
    case "right":
      return "end";
  }
}

function svgDominantBaseline(anchor: "top" | "middle" | "bottom"): string {
  switch (anchor) {
    case "top":
      return "alphabetic";
    case "middle":
      return "middle";
    case "bottom":
      return "alphabetic";
  }
}

/**
 * PowerPoint-style placeholder prompt text. Mirrors the wording in
 * `BUILTIN_LAYOUTS` (which is dropped during `clonePlaceholdersIntoSlide`
 * because we don't want the prompt to round-trip as real text content).
 * Falls back to a generic "Click to add content" for placeholder types
 * we don't have a specific prompt for.
 */
function placeholderHintLabel(type: string): string {
  switch (type) {
    case "title":
    case "ctrTitle":
      return "Click to add title";
    case "subTitle":
      return "Click to add subtitle";
    case "body":
      return "Click to add text";
    case "pic":
      return "Click to add picture";
    case "chart":
      return "Click to add chart";
    case "tbl":
      return "Click to add table";
    case "dgm":
      return "Click to add diagram";
    case "media":
      return "Click to add media";
    case "ftr":
      return "Footer";
    case "hdr":
      return "Header";
    case "dt":
      return "Date";
    case "sldNum":
      return "Slide number";
    default:
      return "Click to add content";
  }
}

/** Compact "image" glyph (frame + sun + mountain) drawn at the given anchor. */
function renderPictureIcon(x: number, y: number, size: number, color: string): string {
  const stroke = `stroke="${color}" stroke-width="${Math.max(1, size / 24)}" fill="none" stroke-linejoin="round" stroke-linecap="round"`;
  const left = x;
  const top = y;
  const right = x + size;
  const bottom = y + size;
  const sunR = size * 0.1;
  const sunCx = left + size * 0.32;
  const sunCy = top + size * 0.32;
  const mountainBase = bottom - size * 0.12;
  return [
    `<rect x="${left}" y="${top}" width="${size}" height="${size}" rx="${size * 0.08}" ry="${size * 0.08}" ${stroke}/>`,
    `<circle cx="${sunCx}" cy="${sunCy}" r="${sunR}" ${stroke}/>`,
    `<polyline points="${left + size * 0.12},${mountainBase} ${left + size * 0.42},${top + size * 0.55} ${left + size * 0.62},${top + size * 0.72} ${left + size * 0.78},${top + size * 0.5} ${right - size * 0.06},${mountainBase}" ${stroke}/>`,
  ].join("");
}

/**
 * Render a text shape's body as a `<foreignObject>` containing styled
 * HTML so the browser word-wraps long runs to the shape's width. The
 * older `<text>`/`<tspan>` path produced one unbroken line per
 * paragraph and overflowed the geometry on anything longer than the
 * box — exactly what users notice as "no automatic line breaks".
 *
 * Paragraph alignment, vertical anchor (`<a:bodyPr anchor>`), and per-
 * run formatting (bold/italic/underline/strike, font-family/size,
 * fill, highlight) are all honoured so the wrapped output matches the
 * editing overlay's flow.
 */
function renderWrappedTextHtml(
  shape: TextShape,
  cxEmu: number,
  cyEmu: number,
  shapeFillHex: string | null,
  theme: ThemeColorScheme
): string {
  const w = u(cxEmu);
  const h = u(cyEmu);
  const defaultColor = pickContrastingTextColor(shapeFillHex, theme);
  const baseFontSizePx = u(estimateFontSizeEmu(shape.txBody.paragraphs[0]));
  const anchor = readBodyAnchor(shape.txBody.bodyPrRaw);
  const insets = readBodyInsets(shape.txBody.bodyPrRaw);
  // `wrap="none"` (rare) means "let the text overflow". Default in
  // PowerPoint's `<a:bodyPr>` is "square" → wrap.
  const wrapMode = readBodyWrap(shape.txBody.bodyPrRaw);

  const justifyContent = anchor === "ctr" ? "center" : anchor === "b" ? "flex-end" : "flex-start";

  const paragraphs = shape.txBody.paragraphs.map((p) => paragraphToHtml(p, theme)).join("");

  // foreignObject sits inside an SVG that's itself inside a `<g>` with
  // the shape's `translate(x y)` already applied — so we anchor at
  // (0, 0) here and use the shape's intrinsic cx/cy as our box.
  const containerStyle = [
    "width:100%",
    "height:100%",
    "display:flex",
    "flex-direction:column",
    `justify-content:${justifyContent}`,
    `padding:${u(insets.t)}px ${u(insets.r)}px ${u(insets.b)}px ${u(insets.l)}px`,
    "box-sizing:border-box",
    `color:#${defaultColor}`,
    "font-family:sans-serif",
    `font-size:${baseFontSizePx}px`,
    "line-height:1.2",
    wrapMode === "none" ? "white-space:pre" : "white-space:pre-wrap",
    "word-wrap:break-word",
    "overflow:hidden",
  ].join(";");
  return [
    `<foreignObject x="0" y="0" width="${w}" height="${h}">`,
    `<div xmlns="http://www.w3.org/1999/xhtml" style="${containerStyle}">`,
    paragraphs,
    `</div>`,
    `</foreignObject>`,
  ].join("");
}

function paragraphToHtml(p: TextParagraph, theme: ThemeColorScheme): string {
  const align =
    p.properties.alignment === "center"
      ? "center"
      : p.properties.alignment === "right"
        ? "right"
        : p.properties.alignment === "justify"
          ? "justify"
          : "left";
  const flatLen = p.runs.reduce((acc, r) => acc + (r.isLineBreak ? 0 : r.text.length), 0);
  if (flatLen === 0) {
    // Empty paragraph — emit a non-breaking space so the line takes up
    // a row (matching how PowerPoint renders blank paragraphs).
    return `<div style="text-align:${align}">&#160;</div>`;
  }
  const runs = p.runs.map((r) => runToHtml(r, theme)).join("");
  return `<div style="text-align:${align}">${runs}</div>`;
}

function runToHtml(r: TextRun, theme: ThemeColorScheme): string {
  if (r.isLineBreak) return "<br/>";
  const styles: string[] = [];
  if (r.properties.bold) styles.push("font-weight:bold");
  if (r.properties.italic) styles.push("font-style:italic");
  const decorations: string[] = [];
  if (r.properties.underline) decorations.push("underline");
  if (r.properties.strike) decorations.push("line-through");
  if (decorations.length > 0) styles.push(`text-decoration:${decorations.join(" ")}`);
  if (r.properties.fontFamily) styles.push(`font-family:${escXml(r.properties.fontFamily)}`);
  styles.push(`color:#${resolveRunFill(r, theme)}`);
  if (r.properties.fontSizeHundredths !== undefined) {
    const pt = r.properties.fontSizeHundredths / 100;
    styles.push(`font-size:${pt}pt`);
  }
  if (r.properties.highlight) styles.push(`background-color:#${escXml(r.properties.highlight)}`);
  return `<span style="${styles.join(";")}">${escXml(r.text)}</span>`;
}

function readBodyAnchor(bodyPr: OpaqueXml | undefined): "t" | "ctr" | "b" {
  const v = bodyPr?.attrs?.anchor ?? bodyPr?.rawAttrs?.["@_anchor"];
  if (v === "ctr") return "ctr";
  if (v === "b") return "b";
  return "t";
}

function readBodyWrap(bodyPr: OpaqueXml | undefined): "square" | "none" {
  const v = bodyPr?.attrs?.wrap ?? bodyPr?.rawAttrs?.["@_wrap"];
  return v === "none" ? "none" : "square";
}

function readBodyInsets(bodyPr: OpaqueXml | undefined): { l: number; r: number; t: number; b: number } {
  // PowerPoint defaults (ECMA-376 21.1.2.1.1 bodyPr): lIns/rIns 91440,
  // tIns/bIns 45720 (in EMU). Honour explicit overrides when present.
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

/**
 * Build the SVG primitive(s) for a given `<a:prstGeom>` preset. Accepts
 * EMU dimensions and returns scaled, ready-to-emit XML. The `prst` value
 * comes from `readPrstGeom` and is `null` when the shape didn't declare
 * one — we then default to a rectangle so untyped placeholders still
 * have a visible bounding box.
 *
 * Only the presets the editor's "Insert shape" menu can produce are
 * special-cased; everything else (`hexagon`, `cloud`, …) renders as a
 * plain rect, which is enough to preserve the layout while we wait to
 * port the full preset geometry library from `presetShapeDefinitions`.
 */
function renderGeometry(
  prst: string | null,
  cxEmu: number,
  cyEmu: number,
  fill: string,
  strokeAttrs: string
): string {
  const w = u(cxEmu);
  const h = u(cyEmu);
  switch (prst) {
    case "ellipse":
      return `<ellipse cx="${w / 2}" cy="${h / 2}" rx="${w / 2}" ry="${h / 2}" fill="${fill}"${strokeAttrs}/>`;
    case "roundRect": {
      // ~12% corner radius matches PowerPoint's default `<a:avLst>` value.
      const r = Math.min(w, h) * 0.12;
      return `<rect width="${w}" height="${h}" rx="${r}" ry="${r}" fill="${fill}"${strokeAttrs}/>`;
    }
    case "triangle":
      return `<polygon points="${w / 2},0 ${w},${h} 0,${h}" fill="${fill}"${strokeAttrs}/>`;
    case "rtTriangle":
      return `<polygon points="0,0 0,${h} ${w},${h}" fill="${fill}"${strokeAttrs}/>`;
    case "diamond":
      return `<polygon points="${w / 2},0 ${w},${h / 2} ${w / 2},${h} 0,${h / 2}" fill="${fill}"${strokeAttrs}/>`;
    case "rightArrow": {
      // Body height = 60% of total; head occupies the right 30% width.
      const bodyH = h * 0.6;
      const bodyTop = (h - bodyH) / 2;
      const bodyBot = bodyTop + bodyH;
      const headStart = w * 0.7;
      return `<polygon points="0,${bodyTop} ${headStart},${bodyTop} ${headStart},0 ${w},${h / 2} ${headStart},${h} ${headStart},${bodyBot} 0,${bodyBot}" fill="${fill}"${strokeAttrs}/>`;
    }
    case "rect":
    case null:
    default:
      return `<rect width="${w}" height="${h}" fill="${fill}"${strokeAttrs}/>`;
  }
}

interface StrokeStyle {
  readonly color: string;
  readonly widthEmu: number;
}

function readStrokeFromOpaque(
  children: ReadonlyArray<OpaqueXml>,
  theme: ThemeColorScheme
): StrokeStyle | null {
  for (const c of children) {
    if (c.tag !== "a:ln") continue;
    // `<a:ln>` may carry `noFill` (explicit "no stroke") — in which case
    // we honour it even if a width attribute is present.
    let hasNoFill = false;
    let nestedFill: string | null = null;
    for (const inner of c.subtree) {
      if (!inner || typeof inner !== "object" || Array.isArray(inner)) continue;
      const obj = inner as Record<string, unknown>;
      const keys = Object.keys(obj).filter((k) => k !== ":@");
      if (keys.length !== 1) continue;
      const tag = keys[0];
      if (tag === "a:noFill") {
        hasNoFill = true;
        continue;
      }
      if (tag === "a:solidFill") {
        nestedFill = readFillFromOpaque([toOpaqueChild(obj)], theme);
      }
    }
    if (hasNoFill) return null;
    const wRaw = c.attrs?.w ?? c.rawAttrs?.["@_w"];
    const w = typeof wRaw === "string" ? Number(wRaw) : NaN;
    const widthEmu = Number.isFinite(w) && w > 0 ? w : 6350;
    if (nestedFill) return { color: nestedFill, widthEmu };
    // No nested fill: still draw a thin black hairline if a width was set
    // explicitly (matches PowerPoint's "outline by default").
    if (Number.isFinite(w) && w > 0) {
      return { color: theme.tx1, widthEmu };
    }
  }
  return null;
}

function toOpaqueChild(obj: Record<string, unknown>): OpaqueXml {
  const keys = Object.keys(obj).filter((k) => k !== ":@");
  const tag = keys[0] ?? "";
  const attrsRaw = (obj[":@"] as Record<string, unknown> | undefined) ?? {};
  const attrs: Record<string, string> = {};
  for (const [k, v] of Object.entries(attrsRaw)) {
    if (typeof v === "string" && k.startsWith("@_")) attrs[k.slice(2)] = v;
  }
  const subtreeNode = obj[tag];
  const subtree = Array.isArray(subtreeNode) ? subtreeNode : [];
  return { tag, attrs, subtree, rawAttrs: attrsRaw } as OpaqueXml;
}

function readPrstGeom(children: ReadonlyArray<OpaqueXml>): string | null {
  for (const c of children) {
    if (c.tag !== "a:prstGeom") continue;
    const prst = c.attrs?.prst ?? c.rawAttrs?.["@_prst"];
    if (typeof prst === "string") return prst;
  }
  return null;
}

/**
 * If a shape has a dark fill, the body-text default (`tx1`, usually black)
 * is unreadable. Pick a light text color in that case so the renderer
 * remains a useful preview even when an authoring tool relied on master
 * placeholders we don't yet inherit. Per-run fills (typed `properties.color`
 * or opaque `solidFill`) still win — this only affects runs that fell back
 * to the default.
 */
function pickContrastingTextColor(fillHex: string | null, theme: ThemeColorScheme): string {
  if (!fillHex) return theme.tx1;
  const v = fillHex.replace(/^#/, "");
  if (v.length !== 6) return theme.tx1;
  const r = parseInt(v.slice(0, 2), 16);
  const g = parseInt(v.slice(2, 4), 16);
  const b = parseInt(v.slice(4, 6), 16);
  const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
  return luminance < 0.5 ? theme.bg1 : theme.tx1;
}

/**
 * Resolve the fill color for a text run. Order of precedence:
 *  1. Typed `properties.color` (already extracted from `a:solidFill > a:srgbClr`).
 *  2. `a:solidFill > a:schemeClr|a:srgbClr|a:sysClr` captured in
 *     `properties.opaqueChildren`. Scheme refs resolve through `theme`.
 *  3. `theme.tx1` (body text default).
 */
function resolveRunFill(r: TextRun, theme: ThemeColorScheme): string {
  if (r.properties.color) return escXml(r.properties.color);
  const fromOpaque = readFillFromOpaque(r.properties.opaqueChildren ?? [], theme);
  if (fromOpaque) return escXml(fromOpaque);
  return theme.tx1;
}

function readFillFromOpaque(children: ReadonlyArray<OpaqueXml>, theme: ThemeColorScheme): string | null {
  for (const c of children) {
    if (c.tag !== "a:solidFill") continue;
    for (const inner of c.subtree) {
      const color = resolveSchemeOrLiteralColor(inner, theme);
      if (color) return color;
    }
  }
  return null;
}

/**
 * Resolve a single fast-xml-parser node like `{ "a:srgbClr": [], ":@": { "@_val": "FF0000" }}`
 * to a 6-char hex colour, taking literal/sys/scheme colours into account.
 * Exported so the slide-level `<p:bg>` walker can share the same logic.
 */
export function resolveSchemeOrLiteralColor(inner: unknown, theme: ThemeColorScheme): string | null {
  if (!inner || typeof inner !== "object" || Array.isArray(inner)) return null;
  const obj = inner as Record<string, unknown>;
  const keys = Object.keys(obj).filter((k) => k !== ":@");
  if (keys.length !== 1) return null;
  const tag = keys[0];
  const attrs = obj[":@"] as Record<string, unknown> | undefined;
  const val = attrs && typeof attrs === "object" ? attrs["@_val"] : undefined;
  if (typeof val !== "string") return null;
  if (tag === "a:srgbClr") return val;
  if (tag === "a:sysClr") {
    const last = attrs && typeof attrs === "object" ? attrs["@_lastClr"] : undefined;
    return typeof last === "string" ? last : val;
  }
  if (tag === "a:schemeClr") {
    const mapped = mapSchemeName(val);
    if (mapped) return theme[mapped];
  }
  return null;
}

function mapSchemeName(name: string): keyof ThemeColorScheme | null {
  switch (name) {
    case "accent1":
    case "accent2":
    case "accent3":
    case "accent4":
    case "accent5":
    case "accent6":
    case "tx1":
    case "tx2":
    case "bg1":
    case "bg2":
    case "hlink":
    case "folHlink":
      return name;
    case "dk1":
      return "tx1";
    case "lt1":
      return "bg1";
    case "dk2":
      return "tx2";
    case "lt2":
      return "bg2";
    default:
      return null;
  }
}

function estimateFontSizeEmu(p: TextParagraph | undefined): number {
  if (!p) return 18 * 12700; // 18pt default
  const r = p.runs.find((x) => !x.isLineBreak);
  if (r?.properties.fontSizeHundredths !== undefined) {
    return (r.properties.fontSizeHundredths / 100) * 12700;
  }
  return 18 * 12700;
}

function pictureToSvg(shape: Picture, ctx: SvgRenderCtx): string {
  const box = shapeBoundingBox(shape);
  if (!box) return groupOpen("pic", shape.id) + groupClose();
  const url = ctx.mediaUrls?.get(shape.mediaPartPath);
  const x = u(box.x);
  const y = u(box.y);
  const cx = u(box.cx);
  const cy = u(box.cy);
  if (!url) {
    return [
      groupOpen("pic", shape.id),
      `<rect x="${x}" y="${y}" width="${cx}" height="${cy}" fill="#f4f4f5" stroke="#d4d4d8"/>`,
      `<text x="${x + cx / 2}" y="${y + cy / 2}" text-anchor="middle" font-size="${u(estimateLabelSizeEmu(box.cx, box.cy))}" fill="#71717a">image</text>`,
      groupClose(),
    ].join("");
  }
  return [
    groupOpen("pic", shape.id),
    `<image href="${escXml(url)}" x="${x}" y="${y}" width="${cx}" height="${cy}" preserveAspectRatio="xMidYMid meet"/>`,
    groupClose(),
  ].join("");
}

function groupShapeToSvg(shape: GroupShape, ctx: SvgRenderCtx): string {
  const inner = shape.children.map((c) => shapeToSvg(c, ctx)).join("");
  const tx = shape.position?.xEmu ?? 0;
  const ty = shape.position?.yEmu ?? 0;
  return [
    groupOpen("group", shape.id, { transform: `translate(${u(tx)} ${u(ty)})` }),
    inner,
    groupClose(),
  ].join("");
}

/**
 * Render a `TableShape` as an SVG `<g>` containing per-cell rectangles
 * and centered text. Width per column comes from `columnWidths`; row
 * heights distribute the table-bbox height equally if the row's stored
 * height is `0` (typical when authoring tools leave layout to the
 * renderer). Visual fidelity is intentionally simple — borders and fills
 * are not rendered yet (P2 work). The point of F2.4 is that the
 * renderer never crashes on table shapes and shows the cell text.
 */
function tableToSvg(shape: TableShape, ctx: SvgRenderCtx): string {
  const box = shapeBoundingBox(shape);
  if (!box) return groupOpen("table", shape.id) + groupClose();
  const theme = ctx.theme ?? DEFAULT_THEME;

  const colCount = shape.columnWidths.length;
  const totalColWidth = shape.columnWidths.reduce((a, b) => a + b, 0) || box.cx;
  const colXs: number[] = [];
  let acc = 0;
  for (const w of shape.columnWidths) {
    colXs.push(acc);
    acc += w;
  }

  // Determine per-row heights: prefer stored height, fall back to even split.
  const storedTotal = shape.rows.reduce((a, r) => a + r.height, 0);
  const rowHeights = shape.rows.map((r) =>
    storedTotal > 0 ? r.height : Math.floor(box.cy / Math.max(1, shape.rows.length))
  );

  const parts: string[] = [];
  parts.push(groupOpen("table", shape.id, { transform: `translate(${u(box.x)} ${u(box.y)})` }));
  parts.push(`<rect x="0" y="0" width="${u(box.cx)}" height="${u(box.cy)}" fill="white" stroke="#9CA3AF"/>`);

  let yAcc = 0;
  for (let r = 0; r < shape.rows.length; r++) {
    const row = shape.rows[r]!;
    const rowH = rowHeights[r]!;
    for (let c = 0; c < Math.min(row.cells.length, colCount); c++) {
      const cell = row.cells[c]!;
      const cx = colXs[c]!;
      const cw = shape.columnWidths[c]!;
      parts.push(
        `<rect x="${u(cx)}" y="${u(yAcc)}" width="${u(cw)}" height="${u(rowH)}" fill="transparent" stroke="#9CA3AF"/>`
      );
      const text = cellToFlatText(cell.txBody.paragraphs);
      if (text.length > 0) {
        const fontSize = u(estimateFontSizeEmu(cell.txBody.paragraphs[0]));
        const fillColor = theme.tx1;
        parts.push(
          `<text x="${u(cx + cw / 2)}" y="${u(yAcc + rowH / 2)}" text-anchor="middle" dominant-baseline="middle" font-family="sans-serif" font-size="${fontSize}" fill="#${fillColor}" xml:space="preserve">${escXml(text)}</text>`
        );
      }
    }
    yAcc += rowH;
  }
  parts.push(groupClose());
  void totalColWidth;
  return parts.join("");
}

function cellToFlatText(paragraphs: ReadonlyArray<TextParagraph>): string {
  const lines = paragraphs.map((p) =>
    p.runs
      .filter((r) => !r.isLineBreak)
      .map((r) => r.text)
      .join("")
  );
  return lines.filter((s) => s.length > 0).join(" / ");
}

/**
 * Render a `ChartShape` as native SVG. Bar / line / pie / area chart
 * types get a minimal native rendering; unknown types fall back to a
 * labeled placeholder rectangle. Visual fidelity is intentionally
 * simple — the goal is "you can tell at a glance which kind of chart
 * this is and what the magnitude of each series looks like", not
 * pixel-perfect parity with PowerPoint's renderer.
 */
function chartToSvg(shape: ChartShape, ctx: SvgRenderCtx): string {
  const box = shapeBoundingBox(shape);
  if (!box) return groupOpen("chart", shape.id) + groupClose();
  const part = ctx.charts?.get(shape.chartPartPath);
  if (!part) return chartPlaceholder(shape, box, "chart");

  const palette = chartPalette(ctx.theme ?? DEFAULT_THEME);
  switch (part.chartType) {
    case "bar":
      return chartBarSvg(shape, box, part, palette);
    case "line":
      return chartLineSvg(shape, box, part, palette);
    case "area":
      return chartAreaSvg(shape, box, part, palette);
    case "pie":
      return chartPieSvg(shape, box, part, palette);
    case "unsupported":
      return chartPlaceholder(shape, box, `${part.title ?? "chart"} · unsupported`);
  }
}

interface ChartBox {
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
}

function chartPlaceholder(shape: ChartShape, box: ChartBox, label: string): string {
  return [
    groupOpen("chart", shape.id),
    `<rect x="${u(box.x)}" y="${u(box.y)}" width="${u(box.cx)}" height="${u(box.cy)}" fill="#f9fafb" stroke="#9CA3AF"/>`,
    `<text x="${u(box.x + box.cx / 2)}" y="${u(box.y + box.cy / 2)}" text-anchor="middle" dominant-baseline="middle" font-family="sans-serif" font-size="${u(estimateLabelSizeEmu(box.cx, box.cy))}" fill="#374151">${escXml(label)}</text>`,
    groupClose(),
  ].join("");
}

function chartPalette(theme: ThemeColorScheme): ReadonlyArray<string> {
  return [
    `#${theme.accent1}`,
    `#${theme.accent2}`,
    `#${theme.accent3}`,
    `#${theme.accent4}`,
    `#${theme.accent5}`,
    `#${theme.accent6}`,
  ];
}

interface PlotArea {
  readonly inner: ChartBox;
  readonly titleHeight: number;
  readonly titleY: number;
  readonly valueMax: number;
  readonly valueMin: number;
}

function plotAreaFor(box: ChartBox, part: ChartPart): PlotArea {
  const padX = box.cx * 0.06;
  const padY = box.cy * 0.06;
  const titleHeight = part.title ? box.cy * 0.12 : 0;
  const inner: ChartBox = {
    x: box.x + padX,
    y: box.y + padY + titleHeight,
    cx: box.cx - 2 * padX,
    cy: box.cy - 2 * padY - titleHeight,
  };
  let max = 0;
  let min = 0;
  for (const s of part.series) {
    for (const v of s.values) {
      if (v > max) max = v;
      if (v < min) min = v;
    }
  }
  if (max === min) max = min + 1;
  return { inner, titleHeight, titleY: box.y + padY, valueMax: max, valueMin: min };
}

function chartTitleSvg(box: ChartBox, part: ChartPart): string {
  if (!part.title) return "";
  const fs = estimateLabelSizeEmu(box.cx, box.cy);
  return `<text x="${u(box.x + box.cx / 2)}" y="${u(box.y + box.cy * 0.06 + fs)}" text-anchor="middle" font-family="sans-serif" font-size="${u(fs)}" fill="#111827">${escXml(part.title)}</text>`;
}

function chartBarSvg(
  shape: ChartShape,
  box: ChartBox,
  part: ChartPart,
  palette: ReadonlyArray<string>
): string {
  const pa = plotAreaFor(box, part);
  const out: string[] = [groupOpen("chart", shape.id)];
  out.push(
    `<rect x="${u(box.x)}" y="${u(box.y)}" width="${u(box.cx)}" height="${u(box.cy)}" fill="white" stroke="#E5E7EB"/>`
  );
  out.push(chartTitleSvg(box, part));
  const groupCount = Math.max(1, part.categories.length || part.series[0]?.values.length || 1);
  const seriesCount = Math.max(1, part.series.length);
  const groupGap = pa.inner.cx / groupCount;
  const barGap = groupGap / (seriesCount + 1);
  const barWidth = barGap * 0.8;
  const range = pa.valueMax - pa.valueMin;
  const baselineY = pa.inner.y + pa.inner.cy;
  for (let g = 0; g < groupCount; g++) {
    const groupX = pa.inner.x + g * groupGap;
    for (let si = 0; si < seriesCount; si++) {
      const v = part.series[si]?.values[g] ?? 0;
      const h = (Math.max(0, v) / range) * pa.inner.cy;
      const x = groupX + barGap * (si + 0.5) - barWidth / 2;
      const y = baselineY - h;
      const fill = palette[si % palette.length];
      out.push(
        `<rect x="${u(x)}" y="${u(y)}" width="${u(barWidth)}" height="${u(h)}" fill="${escXml(fill)}"/>`
      );
    }
  }
  out.push(
    `<line x1="${u(pa.inner.x)}" y1="${u(baselineY)}" x2="${u(pa.inner.x + pa.inner.cx)}" y2="${u(baselineY)}" stroke="#9CA3AF"/>`
  );
  out.push(groupClose());
  return out.join("");
}

function chartLineSvg(
  shape: ChartShape,
  box: ChartBox,
  part: ChartPart,
  palette: ReadonlyArray<string>
): string {
  return chartLineOrAreaSvg(shape, box, part, palette, false);
}

function chartAreaSvg(
  shape: ChartShape,
  box: ChartBox,
  part: ChartPart,
  palette: ReadonlyArray<string>
): string {
  return chartLineOrAreaSvg(shape, box, part, palette, true);
}

function chartLineOrAreaSvg(
  shape: ChartShape,
  box: ChartBox,
  part: ChartPart,
  palette: ReadonlyArray<string>,
  filled: boolean
): string {
  const pa = plotAreaFor(box, part);
  const out: string[] = [groupOpen("chart", shape.id)];
  out.push(
    `<rect x="${u(box.x)}" y="${u(box.y)}" width="${u(box.cx)}" height="${u(box.cy)}" fill="white" stroke="#E5E7EB"/>`
  );
  out.push(chartTitleSvg(box, part));
  const range = pa.valueMax - pa.valueMin;
  const baselineY = pa.inner.y + pa.inner.cy;
  for (let si = 0; si < part.series.length; si++) {
    const series = part.series[si]!;
    const n = series.values.length;
    if (n === 0) continue;
    const stepX = n === 1 ? 0 : pa.inner.cx / (n - 1);
    const points: string[] = [];
    for (let i = 0; i < n; i++) {
      const v = series.values[i] ?? 0;
      const x = pa.inner.x + (n === 1 ? pa.inner.cx / 2 : i * stepX);
      const y = baselineY - ((Math.max(0, v) - pa.valueMin) / range) * pa.inner.cy;
      points.push(`${u(x)},${u(y)}`);
    }
    const stroke = palette[si % palette.length];
    if (filled) {
      const polyPoints = [
        ...points,
        `${u(pa.inner.x + pa.inner.cx)},${u(baselineY)}`,
        `${u(pa.inner.x)},${u(baselineY)}`,
      ].join(" ");
      out.push(
        `<polygon points="${polyPoints}" fill="${escXml(stroke)}" fill-opacity="0.35" stroke="${escXml(stroke)}" stroke-width="${u(Math.max(2, pa.inner.cy / 200))}"/>`
      );
    } else {
      out.push(
        `<polyline points="${points.join(" ")}" fill="none" stroke="${escXml(stroke)}" stroke-width="${u(Math.max(2, pa.inner.cy / 150))}"/>`
      );
    }
  }
  out.push(
    `<line x1="${u(pa.inner.x)}" y1="${u(baselineY)}" x2="${u(pa.inner.x + pa.inner.cx)}" y2="${u(baselineY)}" stroke="#9CA3AF"/>`
  );
  out.push(groupClose());
  return out.join("");
}

function chartPieSvg(
  shape: ChartShape,
  box: ChartBox,
  part: ChartPart,
  palette: ReadonlyArray<string>
): string {
  const out: string[] = [groupOpen("chart", shape.id)];
  out.push(
    `<rect x="${u(box.x)}" y="${u(box.y)}" width="${u(box.cx)}" height="${u(box.cy)}" fill="white" stroke="#E5E7EB"/>`
  );
  out.push(chartTitleSvg(box, part));
  const series = part.series[0];
  const titleHeight = part.title ? box.cy * 0.12 : 0;
  const padX = box.cx * 0.06;
  const padY = box.cy * 0.06;
  const innerCx = box.cx - 2 * padX;
  const innerCy = box.cy - 2 * padY - titleHeight;
  const r = Math.min(innerCx, innerCy) / 2;
  const cxc = box.x + padX + innerCx / 2;
  const cyc = box.y + padY + titleHeight + innerCy / 2;
  if (!series || series.values.length === 0) {
    out.push(`<circle cx="${u(cxc)}" cy="${u(cyc)}" r="${u(r)}" fill="#F3F4F6" stroke="#9CA3AF"/>`);
    out.push(groupClose());
    return out.join("");
  }
  const total = series.values.reduce((a, b) => a + Math.max(0, b), 0) || 1;
  let startAngle = -Math.PI / 2;
  for (let i = 0; i < series.values.length; i++) {
    const v = Math.max(0, series.values[i] ?? 0);
    const sweep = (v / total) * Math.PI * 2;
    const endAngle = startAngle + sweep;
    const x1 = cxc + r * Math.cos(startAngle);
    const y1 = cyc + r * Math.sin(startAngle);
    const x2 = cxc + r * Math.cos(endAngle);
    const y2 = cyc + r * Math.sin(endAngle);
    const largeArc = sweep > Math.PI ? 1 : 0;
    const fill = palette[i % palette.length];
    if (sweep >= Math.PI * 2 - 1e-9) {
      out.push(`<circle cx="${u(cxc)}" cy="${u(cyc)}" r="${u(r)}" fill="${escXml(fill)}"/>`);
    } else {
      out.push(
        `<path d="M ${u(cxc)} ${u(cyc)} L ${u(x1)} ${u(y1)} A ${u(r)} ${u(r)} 0 ${largeArc} 1 ${u(x2)} ${u(y2)} Z" fill="${escXml(fill)}"/>`
      );
    }
    startAngle = endAngle;
  }
  out.push(groupClose());
  return out.join("");
}

function opaqueShapeToSvg(shape: OpaqueShape): string {
  const box = shapeBoundingBox(shape);
  if (!box) {
    return groupOpen("opaque", shape.id) + groupClose();
  }
  return [
    groupOpen("opaque", shape.id),
    `<rect class="placeholder" x="${u(box.x)}" y="${u(box.y)}" width="${u(box.cx)}" height="${u(box.cy)}" fill="#fafafa" stroke="#a1a1aa" stroke-dasharray="${u(50000)},${u(30000)}"/>`,
    `<text x="${u(box.x + box.cx / 2)}" y="${u(box.y + box.cy / 2)}" text-anchor="middle" font-size="${u(estimateLabelSizeEmu(box.cx, box.cy))}" fill="#71717a">${escXml(shape.tag)}</text>`,
    groupClose(),
  ].join("");
}

function estimateLabelSizeEmu(cx: number, cy: number): number {
  return Math.max(60000, Math.floor(Math.min(cx, cy) / 8));
}

function groupOpen(cls: string, id: string, extra: Record<string, string> = {}): string {
  const a: string[] = [`class="shape ${cls}"`, `data-shape-id="${escXml(id)}"`];
  for (const [k, v] of Object.entries(extra)) a.push(`${k}="${escXml(v)}"`);
  return `<g ${a.join(" ")}>`;
}

function groupClose(): string {
  return "</g>";
}
