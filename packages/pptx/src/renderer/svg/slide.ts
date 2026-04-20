import type { OpaqueXml, Shape, Slide } from "../../model/types.js";
import { DEFAULT_THEME } from "../layout/color.js";
import { slideViewBox } from "../layout/slide.js";
import { collectObstacles } from "../connector-router/index.js";
import { resolveSchemeOrLiteralColor, shapeToSvg, type SvgRenderCtx } from "./shapes.js";

export function slideToSvgString(slide: Slide, ctx: SvgRenderCtx): string {
  const bg = resolveSlideBackgroundColor(slide.cSldHead, ctx.theme ?? DEFAULT_THEME);
  const fillAttr = bg ? `#${bg}` : "white";
  const shapesByCNvPrId = ctx.shapesByCNvPrId ?? buildShapesByCNvPrId(slide.shapes);
  // Pre-collect the slide-wide obstacle list once; each connector
  // filters out its own anchored target shapes downstream. We skip
  // collection when the caller already supplied obstacles or the
  // slide has no non-connector shapes (cheap predicate).
  const connectorObstacles = ctx.connectorObstacles ?? collectObstacles(slide.shapes, EMPTY_NUMBER_SET);
  const ctxWithExtras: SvgRenderCtx =
    ctx.shapesByCNvPrId && ctx.connectorObstacles ? ctx : { ...ctx, shapesByCNvPrId, connectorObstacles };
  const parts: string[] = [
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${slideViewBox(ctx.slideSize)}" preserveAspectRatio="xMidYMid meet">`,
    CONNECTOR_DEFS_SVG,
    `<rect width="100%" height="100%" fill="${fillAttr}"/>`,
  ];
  for (const s of slide.shapes) parts.push(shapeToSvg(s, ctxWithExtras));
  parts.push(`</svg>`);
  return parts.join("");
}

const EMPTY_NUMBER_SET: ReadonlySet<number> = new Set();

/**
 * `<defs>` block that's emitted once per slide. We declare ONE marker
 * per `ConnectorEndShape` value so the renderer can pick the right
 * one via `marker-end="url(#cxn-${shape})"`. Marker units default to
 * `strokeWidth` so arrowheads scale with the line weight, matching
 * PowerPoint's behaviour.
 *
 * Geometry rationale:
 *   - cxn-arrow:    open chevron — PowerPoint's "Arrow" preset.
 *   - cxn-triangle: filled triangle — PowerPoint's "Triangle" / "stealth".
 *   - cxn-oval:     filled disc — PowerPoint's "Oval".
 *   - cxn-none:     empty marker so `marker-end` references resolve
 *     even when the model says "no end shape" (avoids a flicker if
 *     the value is toggled on/off).
 *
 * `viewBox` 0..10 / `refX=10` keeps the tip aligned with the path
 * endpoint. `refX=5` on the oval centres the disc so the line ends
 * inside the dot rather than peeking out.
 */
const CONNECTOR_DEFS_SVG = [
  `<defs>`,
  `<marker id="cxn-arrow" viewBox="0 0 10 10" refX="10" refY="5" markerWidth="10" markerHeight="10" orient="auto-start-reverse" markerUnits="strokeWidth"><path d="M 0 0 L 10 5 L 0 10 z" fill="context-stroke"/></marker>`,
  `<marker id="cxn-triangle" viewBox="0 0 10 10" refX="10" refY="5" markerWidth="9" markerHeight="9" orient="auto-start-reverse" markerUnits="strokeWidth"><path d="M 0 0 L 10 5 L 0 10 L 2 5 z" fill="context-stroke"/></marker>`,
  `<marker id="cxn-oval" viewBox="0 0 10 10" refX="5" refY="5" markerWidth="6" markerHeight="6" orient="auto" markerUnits="strokeWidth"><circle cx="5" cy="5" r="4" fill="context-stroke"/></marker>`,
  `<marker id="cxn-none" viewBox="0 0 1 1" refX="0" refY="0" markerWidth="0.01" markerHeight="0.01" markerUnits="strokeWidth"></marker>`,
  `</defs>`,
].join("");

export function buildShapesByCNvPrId(shapes: ReadonlyArray<Shape>): Map<number, Shape> {
  const out = new Map<number, Shape>();
  walk(shapes, out);
  return out;
}

function walk(shapes: ReadonlyArray<Shape>, out: Map<number, Shape>): void {
  for (const s of shapes) {
    if (s.cNvPrId > 0) out.set(s.cNvPrId, s);
    if (s.kind === "group") walk(s.children, out);
  }
}

/**
 * Walk `<p:bg>` → `<p:bgPr>` → `<a:solidFill>` and resolve the colour
 * (literal `srgbClr`, `sysClr` with `lastClr`, or `schemeClr` against the
 * theme). Returns null when no background fill is declared (slides that
 * inherit from layout/master fall back to white in our renderer).
 */
export function resolveSlideBackgroundColor(
  cSldHead: ReadonlyArray<OpaqueXml>,
  theme: typeof DEFAULT_THEME
): string | null {
  for (const node of cSldHead) {
    if (node.tag !== "p:bg") continue;
    for (const inner of node.subtree) {
      if (!inner || typeof inner !== "object" || Array.isArray(inner)) continue;
      const obj = inner as Record<string, unknown>;
      const keys = Object.keys(obj).filter((k) => k !== ":@");
      if (keys.length !== 1) continue;
      const tag = keys[0];
      const sub = obj[tag];
      if (tag !== "p:bgPr" || !Array.isArray(sub)) continue;
      for (const child of sub) {
        if (!child || typeof child !== "object" || Array.isArray(child)) continue;
        const cobj = child as Record<string, unknown>;
        const ckeys = Object.keys(cobj).filter((k) => k !== ":@");
        if (ckeys.length !== 1) continue;
        const ctag = ckeys[0];
        if (ctag !== "a:solidFill") continue;
        const csub = cobj[ctag];
        if (!Array.isArray(csub)) continue;
        for (const fillChild of csub) {
          const color = resolveSchemeOrLiteralColor(fillChild, theme);
          if (color) return color;
        }
      }
    }
  }
  return null;
}
