import type {
  ChartPart,
  ChartShape,
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
import { DEFAULT_THEME, type ThemeColorScheme } from "../layout/color.js";
import { shapeBoundingBox } from "../layout/shape.js";
import { SVG_UNIT_PER_EMU } from "../layout/slide.js";
import { escXml } from "./escape.js";

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
    case "opaque":
      return opaqueShapeToSvg(shape);
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
  const strokeAttrs = stroke
    ? ` stroke="#${stroke.color}" stroke-width="${u(stroke.widthEmu)}"`
    : "";

  const hasText = shape.txBody.paragraphs.some((p) =>
    p.runs.some((r) => !r.isLineBreak && r.text.length > 0)
  );

  const out = [
    groupOpen("text", shape.id, { transform: `translate(${u(box.x)} ${u(box.y)})` }),
    `<rect width="${u(box.cx)}" height="${u(box.cy)}" fill="${rectFill}"${strokeAttrs}/>`,
  ];
  if (hasText) {
    const lines = shape.txBody.paragraphs.map((p) => paragraphToTSpan(p, theme));
    const fontSizePx = u(estimateFontSizeEmu(shape.txBody.paragraphs[0]));
    const defaultColor = pickContrastingTextColor(fill, theme);
    out.push(
      `<text x="0" y="${fontSizePx}" font-family="sans-serif" font-size="${fontSizePx}" fill="#${defaultColor}" xml:space="preserve">`,
      lines.join(""),
      `</text>`
    );
  }
  out.push(groupClose());
  return out.join("");
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
function pickContrastingTextColor(
  fillHex: string | null,
  theme: ThemeColorScheme
): string {
  if (!fillHex) return theme.tx1;
  const v = fillHex.replace(/^#/, "");
  if (v.length !== 6) return theme.tx1;
  const r = parseInt(v.slice(0, 2), 16);
  const g = parseInt(v.slice(2, 4), 16);
  const b = parseInt(v.slice(4, 6), 16);
  const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
  return luminance < 0.5 ? theme.bg1 : theme.tx1;
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
    attrs.push(`font-size="${u(sizeEmu)}"`);
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
export function resolveSchemeOrLiteralColor(
  inner: unknown,
  theme: ThemeColorScheme
): string | null {
  if (!inner || typeof inner !== "object" || Array.isArray(inner)) return null;
  const obj = inner as Record<string, unknown>;
  const keys = Object.keys(obj).filter((k) => k !== ":@");
  if (keys.length !== 1) return null;
  const tag = keys[0];
  const attrs = obj[":@"] as Record<string, unknown> | undefined;
  const val =
    attrs && typeof attrs === "object" ? attrs["@_val"] : undefined;
  if (typeof val !== "string") return null;
  if (tag === "a:srgbClr") return val;
  if (tag === "a:sysClr") {
    const last =
      attrs && typeof attrs === "object" ? attrs["@_lastClr"] : undefined;
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
  parts.push(
    `<rect x="0" y="0" width="${u(box.cx)}" height="${u(box.cy)}" fill="white" stroke="#9CA3AF"/>`
  );

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
    p.runs.filter((r) => !r.isLineBreak).map((r) => r.text).join("")
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
    out.push(
      `<circle cx="${u(cxc)}" cy="${u(cyc)}" r="${u(r)}" fill="#F3F4F6" stroke="#9CA3AF"/>`
    );
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
