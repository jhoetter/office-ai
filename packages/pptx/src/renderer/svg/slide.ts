import type { OpaqueXml, Slide } from "../../model/types.js";
import { DEFAULT_THEME } from "../layout/color.js";
import { slideViewBox } from "../layout/slide.js";
import { resolveSchemeOrLiteralColor, shapeToSvg, type SvgRenderCtx } from "./shapes.js";

export function slideToSvgString(slide: Slide, ctx: SvgRenderCtx): string {
  const bg = resolveSlideBackgroundColor(slide.cSldHead, ctx.theme ?? DEFAULT_THEME);
  const fillAttr = bg ? `#${bg}` : "white";
  const parts: string[] = [
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${slideViewBox(ctx.slideSize)}" preserveAspectRatio="xMidYMid meet">`,
    `<rect width="100%" height="100%" fill="${fillAttr}"/>`,
  ];
  for (const s of slide.shapes) parts.push(shapeToSvg(s, ctx));
  parts.push(`</svg>`);
  return parts.join("");
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
