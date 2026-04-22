import type { OpaqueXml } from "../../model/types.js";
import { readFillSpec, type FillSpec, type GradientFillSpec } from "../../model/fill.js";
import type { ThemeColorScheme } from "../layout/color.js";
import { resolveSchemeOrLiteralColor } from "./shapes.js";

/**
 * SVG-render-side representation of a fill: the `paintRef` you stick
 * into a `fill="…"` attribute, plus any `<defs>` content (linearGradient,
 * radialGradient, pattern, …) that must be emitted somewhere ancestor
 * to the painted element. The defs are valid stand-alone SVG so the
 * caller can either inline them into the shape's `<g>` or hoist them
 * into a slide-level `<defs>` block — both work.
 */
export interface ResolvedPaint {
  readonly defs: string;
  readonly paintRef: string;
}

const TRANSPARENT: ResolvedPaint = { defs: "", paintRef: "transparent" };

/** Sanitise a NodeId / string into a valid SVG id segment. */
function sanitizeId(input: string): string {
  return input.replace(/[^a-zA-Z0-9_-]/g, "_");
}

/**
 * Try to resolve a `FillSpec` to an SVG paint. Returns `null` when the
 * fill kind cannot be expressed as native SVG paint (e.g. `picture`
 * fills require the caller to embed a real `<image>` since SVG patterns
 * for tiled bitmaps are out of scope for now). Caller-side fallback:
 * treat as transparent and (optionally) draw the picture as a separate
 * layer.
 */
export function fillSpecToPaint(spec: FillSpec, idHint: string): ResolvedPaint | null {
  switch (spec.type) {
    case "none":
      return TRANSPARENT;
    case "solid": {
      const opacity = spec.alpha !== undefined ? spec.alpha : 1;
      // For < 1 opacity emit a defs-less paint with `fill-opacity` baked
      // into a CSS variable on the paint? Simpler: return rgba via a
      // gradient-of-one trick? Instead, callers that care about alpha
      // pull it via `fill-opacity` themselves; we surface alpha via the
      // hex's 8th-channel approximation when the renderer asks for it.
      // For now keep RGB only; alpha is implicit via stroke/fill-opacity
      // which the shape renderer applies separately.
      void opacity;
      return { defs: "", paintRef: `#${spec.color}` };
    }
    case "gradient": {
      const id = `paint-${sanitizeId(idHint)}`;
      return { defs: gradientDefs(id, spec), paintRef: `url(#${id})` };
    }
    case "pattern":
      // SVG `<pattern>` would be the right approach; out-of-scope for
      // this iteration. Render as the foreground colour so the user at
      // least sees *something* until pattern support lands.
      return { defs: "", paintRef: `#${spec.fgColor}` };
    case "picture":
      // Picture fills need an actual `<image>` (or `<pattern>` wrapping
      // an `<image>`) — caller's responsibility.
      return null;
    default: {
      const _exhaustive: never = spec;
      void _exhaustive;
      return null;
    }
  }
}

function gradientDefs(id: string, spec: GradientFillSpec): string {
  const stops = spec.stops
    .map((s) => {
      const offset = `${(s.pos * 100).toFixed(2)}%`;
      const opacity = s.alpha !== undefined ? ` stop-opacity="${s.alpha.toFixed(3)}"` : "";
      return `<stop offset="${offset}" stop-color="#${s.color}"${opacity}/>`;
    })
    .join("");
  if (spec.kind === "linear") {
    // PowerPoint angles are clockwise from "3 o'clock" (positive x axis).
    // SVG default linearGradient is left → right (0deg in our terms).
    // Convert by computing endpoints on the unit square. We use
    // `gradientUnits="objectBoundingBox"` so the same defs fits any
    // shape — coords are relative to the shape's bounding box (0..1).
    const rad = (spec.angleDeg * Math.PI) / 180;
    const cx = 0.5;
    const cy = 0.5;
    const dx = Math.cos(rad) / 2;
    const dy = Math.sin(rad) / 2;
    const x1 = (cx - dx).toFixed(4);
    const y1 = (cy - dy).toFixed(4);
    const x2 = (cx + dx).toFixed(4);
    const y2 = (cy + dy).toFixed(4);
    return `<linearGradient id="${id}" x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}" gradientUnits="objectBoundingBox">${stops}</linearGradient>`;
  }
  // Radial: concentric circles, centre at the bounding box centre.
  return `<radialGradient id="${id}" cx="0.5" cy="0.5" r="0.5" fx="0.5" fy="0.5" gradientUnits="objectBoundingBox">${stops}</radialGradient>`;
}

/**
 * Resolve a shape-fill from `spPrTail`. Used by `textShapeToSvg` to
 * paint the shape's preset geometry. Falls back to `null` if no fill
 * node is present (caller treats as inherit / transparent).
 *
 * `idHint` is mixed into any generated `<defs>` id so per-shape
 * gradients don't collide with each other.
 */
export function resolveShapePaint(
  spPrTail: ReadonlyArray<OpaqueXml>,
  theme: ThemeColorScheme,
  idHint: string
): ResolvedPaint | null {
  const spec = readFillSpec(spPrTail);
  if (spec) {
    // Solid fills with a literal `srgbClr` round-trip cleanly through
    // `readFillSpec`. Solid fills with a `schemeClr` come back as the
    // `808080` placeholder; re-resolve through the theme so theme
    // colours actually paint correctly.
    if (spec.type === "solid") {
      const themed = resolveSolidThemeColor(spPrTail, theme);
      if (themed) return { defs: "", paintRef: `#${themed}` };
    }
    return fillSpecToPaint(spec, idHint);
  }
  return null;
}

/**
 * Re-walk `<a:solidFill>` looking for the underlying colour child and
 * resolve it through the theme. Handles `srgbClr`, `sysClr`, and
 * `schemeClr` (the case `readFillSpec` collapses to grey).
 */
function resolveSolidThemeColor(
  spPrTail: ReadonlyArray<OpaqueXml>,
  theme: ThemeColorScheme
): string | null {
  for (const c of spPrTail) {
    if (c.tag !== "a:solidFill") continue;
    for (const inner of c.subtree) {
      const color = resolveSchemeOrLiteralColor(inner, theme);
      if (color) return color;
    }
  }
  return null;
}

/**
 * Resolve a slide background fill into a renderable SVG paint.
 * Returns `null` when no `<p:bg>` is declared (slide inherits — caller
 * paints white).
 */
export function resolveSlideBackgroundPaint(
  cSldHead: ReadonlyArray<OpaqueXml>,
  theme: ThemeColorScheme,
  idHint: string
): ResolvedPaint | null {
  // `<p:bg>` wraps `<p:bgPr>` which carries the actual fill choice. We
  // walk it manually instead of `readFillSpec(cSldHead)` because the
  // fill nodes live one nesting level deeper.
  for (const node of cSldHead) {
    if (node.tag !== "p:bg") continue;
    const bgPr = walkInto(node.subtree, "p:bgPr");
    if (!bgPr) continue;
    const fillNodes = collectFillChildren(bgPr);
    const spec = readFillSpec(fillNodes);
    if (!spec) continue;
    if (spec.type === "solid") {
      const themed = resolveSolidThemeColor(fillNodes, theme);
      if (themed) return { defs: "", paintRef: `#${themed}` };
    }
    return fillSpecToPaint(spec, `bg-${sanitizeId(idHint)}`);
  }
  return null;
}

function walkInto(subtree: ReadonlyArray<unknown>, tag: string): ReadonlyArray<unknown> | null {
  for (const inner of subtree) {
    if (!inner || typeof inner !== "object" || Array.isArray(inner)) continue;
    const obj = inner as Record<string, unknown>;
    if (Array.isArray(obj[tag])) return obj[tag] as unknown[];
  }
  return null;
}

function collectFillChildren(subtree: ReadonlyArray<unknown>): OpaqueXml[] {
  const out: OpaqueXml[] = [];
  for (const inner of subtree) {
    if (!inner || typeof inner !== "object" || Array.isArray(inner)) continue;
    const obj = inner as Record<string, unknown>;
    const keys = Object.keys(obj).filter((k) => k !== ":@");
    if (keys.length !== 1) continue;
    const tag = keys[0];
    if (
      tag !== "a:solidFill" &&
      tag !== "a:noFill" &&
      tag !== "a:gradFill" &&
      tag !== "a:pattFill" &&
      tag !== "a:blipFill"
    ) {
      continue;
    }
    const sub = obj[tag];
    const attrs = (obj[":@"] as Record<string, unknown> | undefined) ?? {};
    const rawAttrs: Record<string, string> = {};
    const flatAttrs: Record<string, string> = {};
    for (const [k, v] of Object.entries(attrs)) {
      const sv = String(v);
      rawAttrs[k] = sv;
      const stripped = k.startsWith("@_") ? k.slice(2) : k;
      flatAttrs[stripped] = sv;
    }
    out.push({
      tag,
      attrs: flatAttrs,
      rawAttrs,
      subtree: Array.isArray(sub) ? (sub as unknown[]) : [],
    });
  }
  return out;
}
