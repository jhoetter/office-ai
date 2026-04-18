import type { OpaqueXml, Shape, Slide } from "../../model/types.js";
import { DEFAULT_THEME } from "../layout/color.js";
import { slideViewBox } from "../layout/slide.js";
import { resolveSchemeOrLiteralColor, shapeToSvg, type SvgRenderCtx } from "./shapes.js";

export function slideToSvgString(slide: Slide, ctx: SvgRenderCtx): string {
  const bg = resolveSlideBackgroundColor(slide.cSldHead, ctx.theme ?? DEFAULT_THEME);
  const fillAttr = bg ? `#${bg}` : "white";
  const shapesByCNvPrId = ctx.shapesByCNvPrId ?? buildShapesByCNvPrId(slide.shapes);
  const ctxWithMap: SvgRenderCtx = ctx.shapesByCNvPrId ? ctx : { ...ctx, shapesByCNvPrId };
  const parts: string[] = [
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${slideViewBox(ctx.slideSize)}" preserveAspectRatio="xMidYMid meet">`,
    CONNECTOR_DEFS_SVG,
    `<rect width="100%" height="100%" fill="${fillAttr}"/>`,
  ];
  for (const s of slide.shapes) parts.push(shapeToSvg(s, ctxWithMap));
  parts.push(`</svg>`);
  return parts.join("");
}

/**
 * `<defs>` block that's emitted once per slide — currently just the
 * arrowhead marker every connector references via `marker-end` /
 * `marker-start`. Marker units default to `strokeWidth` so arrow size
 * scales with the line weight, matching PowerPoint's behaviour.
 */
const CONNECTOR_DEFS_SVG = `<defs><marker id="cxn-arrow" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="6" markerHeight="6" orient="auto-start-reverse"><path d="M 0 0 L 10 5 L 0 10 z" fill="context-stroke"/></marker></defs>`;

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
