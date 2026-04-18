import type {
  GroupShape,
  OpaqueShape,
  OpaqueXml,
  Picture,
  Shape,
  SlideSize,
  TextParagraph,
  TextRun,
  TextShape,
} from "../../model/types.js";
import { DEFAULT_THEME, type ThemeColorScheme } from "../layout/color.js";
import { shapeBoundingBox } from "../layout/shape.js";
import { escXml } from "./escape.js";

export interface SvgRenderCtx {
  readonly slideSize: SlideSize;
  readonly theme?: ThemeColorScheme;
  /** Map from media partPath → URL (object URL or data URL). */
  readonly mediaUrls?: ReadonlyMap<string, string>;
}

export function shapeToSvg(shape: Shape, ctx: SvgRenderCtx): string {
  switch (shape.kind) {
    case "text":
      return textShapeToSvg(shape, ctx);
    case "pic":
      return pictureToSvg(shape, ctx);
    case "group":
      return groupShapeToSvg(shape, ctx);
    case "opaque":
      return opaqueShapeToSvg(shape);
  }
}

function textShapeToSvg(shape: TextShape, ctx: SvgRenderCtx): string {
  const box = shapeBoundingBox(shape);
  if (!box) return groupOpen("text", shape.id) + groupClose();
  const theme = ctx.theme ?? DEFAULT_THEME;

  const lines = shape.txBody.paragraphs.map((p) => paragraphToTSpan(p, theme));
  const fontSizePx = estimateFontSizeEmu(shape.txBody.paragraphs[0]);

  return [
    groupOpen("text", shape.id, { transform: `translate(${box.x} ${box.y})` }),
    `<rect width="${box.cx}" height="${box.cy}" fill="transparent"/>`,
    `<text x="0" y="${fontSizePx}" font-family="sans-serif" font-size="${fontSizePx}" fill="#${theme.tx1}" xml:space="preserve">`,
    lines.join(""),
    `</text>`,
    groupClose(),
  ].join("");
}

function paragraphToTSpan(p: TextParagraph, theme: ThemeColorScheme): string {
  if (p.runs.length === 0) {
    return `<tspan x="0" dy="1em"></tspan>`;
  }
  const tspans = p.runs
    .map((r) => runToTSpan(r, theme))
    .join("");
  return `<tspan x="0" dy="1em">${tspans}</tspan>`;
}

function runToTSpan(r: TextRun, theme: ThemeColorScheme): string {
  if (r.isLineBreak) return "<tspan>\n</tspan>";
  const attrs: string[] = [];
  if (r.properties.bold) attrs.push('font-weight="bold"');
  if (r.properties.italic) attrs.push('font-style="italic"');
  if (r.properties.underline) attrs.push('text-decoration="underline"');
  if (r.properties.strike) attrs.push('text-decoration="line-through"');
  if (r.properties.fontFamily) attrs.push(`font-family="${escXml(r.properties.fontFamily)}"`);
  attrs.push(`fill="#${resolveRunFill(r, theme)}"`);
  if (r.properties.fontSizeHundredths !== undefined) {
    // sz is hundredths of a point. EMU per point = 12700.
    const sizeEmu = (r.properties.fontSizeHundredths / 100) * 12700;
    attrs.push(`font-size="${sizeEmu}"`);
  }
  return `<tspan ${attrs.join(" ")}>${escXml(r.text)}</tspan>`;
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

function readFillFromOpaque(
  children: ReadonlyArray<OpaqueXml>,
  theme: ThemeColorScheme
): string | null {
  for (const c of children) {
    if (c.tag !== "a:solidFill") continue;
    for (const inner of c.subtree) {
      if (!inner || typeof inner !== "object" || Array.isArray(inner)) continue;
      const obj = inner as Record<string, unknown>;
      const keys = Object.keys(obj).filter((k) => k !== ":@");
      if (keys.length !== 1) continue;
      const tag = keys[0];
      const attrs = obj[":@"] as Record<string, unknown> | undefined;
      const val =
        attrs && typeof attrs === "object" ? attrs["@_val"] : undefined;
      if (typeof val !== "string") continue;
      if (tag === "a:srgbClr") {
        return val;
      }
      if (tag === "a:sysClr") {
        const last =
          attrs && typeof attrs === "object" ? attrs["@_lastClr"] : undefined;
        return typeof last === "string" ? last : val;
      }
      if (tag === "a:schemeClr") {
        const mapped = mapSchemeName(val);
        if (mapped) return theme[mapped];
      }
    }
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
  if (!url) {
    return [
      groupOpen("pic", shape.id),
      `<rect x="${box.x}" y="${box.y}" width="${box.cx}" height="${box.cy}" fill="#f4f4f5" stroke="#d4d4d8"/>`,
      `<text x="${box.x + box.cx / 2}" y="${box.y + box.cy / 2}" text-anchor="middle" font-size="${estimateLabelSizeEmu(box.cx, box.cy)}" fill="#71717a">image</text>`,
      groupClose(),
    ].join("");
  }
  return [
    groupOpen("pic", shape.id),
    `<image href="${escXml(url)}" x="${box.x}" y="${box.y}" width="${box.cx}" height="${box.cy}" preserveAspectRatio="xMidYMid meet"/>`,
    groupClose(),
  ].join("");
}

function groupShapeToSvg(shape: GroupShape, ctx: SvgRenderCtx): string {
  const inner = shape.children.map((c) => shapeToSvg(c, ctx)).join("");
  // Group transform — translate to position; we don't yet implement chOff/chExt scaling.
  const tx = shape.position?.xEmu ?? 0;
  const ty = shape.position?.yEmu ?? 0;
  return [
    groupOpen("group", shape.id, { transform: `translate(${tx} ${ty})` }),
    inner,
    groupClose(),
  ].join("");
}

function opaqueShapeToSvg(shape: OpaqueShape): string {
  const box = shapeBoundingBox(shape);
  if (!box) {
    return groupOpen("opaque", shape.id) + groupClose();
  }
  return [
    groupOpen("opaque", shape.id),
    `<rect class="placeholder" x="${box.x}" y="${box.y}" width="${box.cx}" height="${box.cy}" fill="#fafafa" stroke="#a1a1aa" stroke-dasharray="50000,30000"/>`,
    `<text x="${box.x + box.cx / 2}" y="${box.y + box.cy / 2}" text-anchor="middle" font-size="${estimateLabelSizeEmu(box.cx, box.cy)}" fill="#71717a">${escXml(shape.tag)}</text>`,
    groupClose(),
  ].join("");
}

function estimateLabelSizeEmu(cx: number, cy: number): number {
  return Math.max(60000, Math.floor(Math.min(cx, cy) / 8));
}

function groupOpen(
  cls: string,
  id: string,
  extra: Record<string, string> = {}
): string {
  const a: string[] = [
    `class="shape ${cls}"`,
    `data-shape-id="${escXml(id)}"`,
  ];
  for (const [k, v] of Object.entries(extra)) a.push(`${k}="${escXml(v)}"`);
  return `<g ${a.join(" ")}>`;
}

function groupClose(): string {
  return "</g>";
}
