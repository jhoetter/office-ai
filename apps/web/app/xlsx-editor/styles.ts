/**
 * Cross-cuts between the editor surface (Toolbar / Grid) and the
 * `@officeai/xlsx` style table. Two responsibilities:
 *
 *   1. Translate an `EffectiveStyle` (the resolved per-cell style
 *      flattened from the workbook's `cellXfs`) into the small set of
 *      CSS properties the Grid actually renders. Excel's full style
 *      surface is huge; we cherry-pick the ones a user can toggle from
 *      the toolbar (font weight/style/decoration, font + fill colour,
 *      horizontal alignment) so what they pick reflects in the cell
 *      immediately.
 *   2. Format a raw `CellValue` with the cell's `numberFormat` for
 *      display. We support the handful of P0 number-format buckets the
 *      toolbar exposes (general, number, currency, percent, date) plus
 *      a graceful fallback for anything custom.
 */

import type { CSSProperties } from "react";
import type { CellValue, EffectiveStyle, StyleBorderSide, StyleColor, StyleTable } from "@officeai/xlsx";
import { flattenCellXf } from "@officeai/xlsx";
import { wrapFontFamily } from "@officeai/text-formatting";

export function styleForCell(
  table: StyleTable,
  styleId: number | undefined
): { css: CSSProperties; effective: EffectiveStyle } {
  const effective = flattenCellXf(table, styleId);
  const css: CSSProperties = {};
  if (effective.font.bold) css.fontWeight = 700;
  if (effective.font.italic) css.fontStyle = "italic";
  const decorations: string[] = [];
  if (effective.font.underline) decorations.push("underline");
  if (effective.font.strike) decorations.push("line-through");
  if (decorations.length > 0) css.textDecoration = decorations.join(" ");
  if (effective.font.color) {
    const c = renderColor(effective.font.color);
    if (c) css.color = c;
  }
  // Font family ↔ OOXML `<font><name val="..."/>` and font size in
  // points ↔ `<font><sz val="..."/>`. Both are part of the shared
  // `TextFormat` contract; without these the toolbar's font/size
  // pickers dispatch correctly but the cell never repaints.
  //
  // `wrapFontFamily` appends a `system-ui, sans-serif` tail so unknown
  // families fall through to a sane sans-serif rather than the UA's
  // default serif. Common Microsoft families (Calibri, Aptos, Cambria,
  // Times New Roman, Arial, …) are also redefined via `@font-face` in
  // `apps/web/app/globals.css` so they resolve to bundled metric-equivalent
  // open-source twins (Carlito / Caladea / Tinos / Arimo / Cousine) on
  // systems without Office installed.
  const family = wrapFontFamily(effective.font.name);
  if (family) css.fontFamily = family;
  if (typeof effective.font.size === "number" && effective.font.size > 0) {
    css.fontSize = `${effective.font.size}pt`;
  }
  if (effective.fill.kind === "pattern" && effective.fill.patternType === "solid") {
    const fg = renderColor(effective.fill.fgColor);
    if (fg) css.background = fg;
  }
  const align = effective.alignment;
  if (align?.horizontal) {
    switch (align.horizontal) {
      case "left":
      case "center":
      case "right":
      case "justify":
        css.justifyContent = horizontalToJustify(align.horizontal);
        css.textAlign = align.horizontal;
        break;
      default:
        break;
    }
  }
  // OOXML `vertical` alignment maps onto flexbox's `align-items`
  // (the Grid's per-cell <div> is `display: flex`). We only override
  // when the workbook is explicit — leaving unstyled cells on the
  // class's default (`center`) avoids a sweeping visual change for
  // sheets that never specified an alignment.
  if (align?.vertical) {
    const v = verticalToAlignItems(align.vertical);
    if (v) css.alignItems = v;
  }
  // `wrapText="1"` switches the cell from single-line truncation to
  // CSS multi-line wrapping. The Grid's per-cell <div> needs both
  // `whiteSpace: pre-wrap` (so explicit newlines are honoured) and
  // an unset `text-overflow` to actually re-flow. We also drop
  // `overflow: hidden`'s ellipsis hint by setting `textOverflow`
  // back to `clip` — Grid still keeps `overflow: hidden` for the
  // row-height clip.
  if (align?.wrapText) {
    css.whiteSpace = "pre-wrap";
    css.textOverflow = "clip";
    css.wordBreak = "break-word";
  }
  // OOXML `indent` is in "indent units" (≈ one character of the
  // default font, or roughly 8 px). Apply only on left/right-aligned
  // cells — Excel ignores indent for centre / fill / general.
  if (typeof align?.indent === "number" && align.indent > 0) {
    const indentPx = align.indent * 8;
    if (align.horizontal === "right") {
      css.paddingRight = (typeof css.paddingRight === "number" ? css.paddingRight : 4) + indentPx;
    } else {
      css.paddingLeft = (typeof css.paddingLeft === "number" ? css.paddingLeft : 4) + indentPx;
    }
  }
  // Borders: Excel stores per-side `<border><left><color/></left>…`
  // entries. We render them via the cell's own CSS borders so they
  // sit on the cell box and overlay the Grid's default 1px guideline
  // (which doubles the border width when both are present — Excel
  // accepts that visually). Skip borders that resolve to `none`.
  const borderCss = bordersToCss(effective.border);
  Object.assign(css, borderCss);
  return { css, effective };
}

function horizontalToJustify(h: "left" | "center" | "right" | "justify"): CSSProperties["justifyContent"] {
  switch (h) {
    case "left":
      return "flex-start";
    case "center":
      return "center";
    case "right":
      return "flex-end";
    case "justify":
      return "space-between";
  }
}

function verticalToAlignItems(v: string): CSSProperties["alignItems"] | undefined {
  switch (v) {
    case "top":
      return "flex-start";
    case "center":
    case "middle":
      return "center";
    case "bottom":
      return "flex-end";
    case "justify":
    case "distributed":
      return "stretch";
    default:
      return undefined;
  }
}

/**
 * Translate an OOXML border style name into a CSS `border-*` shorthand.
 * Excel has 13 named styles; we map families to their nearest CSS
 * equivalent. `none` / unrecognised → `undefined` so we don't paint a
 * border the file didn't ask for.
 */
function borderSideToCss(side: StyleBorderSide | undefined): string | undefined {
  if (!side?.style || side.style === "none") return undefined;
  const color = renderColor(side.color) ?? "#000";
  switch (side.style) {
    case "thin":
    case "hair":
      return `1px solid ${color}`;
    case "medium":
    case "mediumDashed":
    case "mediumDashDot":
    case "mediumDashDotDot":
    case "slantDashDot":
      return `2px solid ${color}`;
    case "thick":
      return `3px solid ${color}`;
    case "double":
      return `3px double ${color}`;
    case "dotted":
      return `1px dotted ${color}`;
    case "dashed":
    case "dashDot":
    case "dashDotDot":
      return `1px dashed ${color}`;
    default:
      return `1px solid ${color}`;
  }
}

function bordersToCss(border: EffectiveStyle["border"]): CSSProperties {
  const out: CSSProperties = {};
  const top = borderSideToCss(border.top);
  const right = borderSideToCss(border.right);
  const bottom = borderSideToCss(border.bottom);
  const left = borderSideToCss(border.left);
  if (top) out.borderTop = top;
  if (right) out.borderRight = right;
  if (bottom) out.borderBottom = bottom;
  if (left) out.borderLeft = left;
  return out;
}

/**
 * Render a `StyleColor` into a CSS colour string. We only handle ARGB
 * hex (the 99% case for the toolbar) — themed/indexed colours fall
 * through to `undefined` so the cell paints with the document
 * default rather than a wrong colour.
 */
function renderColor(c: StyleColor | undefined): string | undefined {
  if (!c) return undefined;
  if (c.rgb) {
    const hex = c.rgb;
    if (hex.length === 8) return `#${hex.slice(2)}`;
    if (hex.length === 6) return `#${hex}`;
  }
  return undefined;
}

/**
 * Catalogue of agent-facing number-format presets exposed by the
 * toolbar. The `code` is what we send into
 * `xlsx:set-cell-format → format.numberFormat`; the handler interns
 * unknown codes via `internNumberFormat` and reuses built-in IDs
 * 0..49 when they match.
 */
export const NUMBER_FORMAT_PRESETS = [
  { id: "general", label: "General", code: "General" },
  { id: "number", label: "Number (1,234.56)", code: "#,##0.00" },
  { id: "currency-eur", label: "Currency (€)", code: '"€"#,##0.00' },
  { id: "currency-usd", label: "Currency ($)", code: '"$"#,##0.00' },
  { id: "percent", label: "Percent (12.34%)", code: "0.00%" },
  { id: "date", label: "Date (yyyy-mm-dd)", code: "yyyy-mm-dd" },
] as const;

export type NumberFormatPresetId = (typeof NUMBER_FORMAT_PRESETS)[number]["id"];

/** Lookup the numFmtId for the built-in presets we ship in the toolbar. */
const BUILTIN_NUM_FMT_IDS: Record<string, number> = {
  General: 0,
  "0.00%": 10,
  "yyyy-mm-dd": 14,
  "#,##0.00": 4,
};

/**
 * Render a `CellValue` for display, applying the resolved
 * number-format when the value is numeric. Non-numbers fall through
 * to the existing string renderer (`true`/`false`, error codes, etc).
 *
 * This is intentionally a *display* approximation — full Excel
 * number-format semantics belong in the engine, not the renderer.
 * The cases below cover everything the toolbar can produce.
 */
export function formatCellValue(value: CellValue, numFmtId: number): string {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value;
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  if (typeof value === "number") return formatNumber(value, numFmtId);
  switch (value.kind) {
    case "error":
      return value.code;
    default: {
      const _exhaustive: never = value.kind;
      void _exhaustive;
      return "";
    }
  }
}

function formatNumber(n: number, numFmtId: number): string {
  switch (numFmtId) {
    case 0:
      return String(n);
    case 1: // 0
      return n.toFixed(0);
    case 2: // 0.00
      return n.toFixed(2);
    case 3: // #,##0
      return Math.round(n).toLocaleString("en-US");
    case 4: // #,##0.00
      return n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
    case 9: // 0%
      return `${Math.round(n * 100)}%`;
    case 10: // 0.00%
      return `${(n * 100).toFixed(2)}%`;
    case 14: // m/d/yy → render ISO for predictability
      return excelSerialToIso(n);
    default:
      return String(n);
  }
}

/**
 * Convert an Excel date serial to `yyyy-mm-dd`. Uses the 1900 epoch
 * with the well-known leap-year bug accounted for (serial 60 ==
 * fictitious 29-Feb-1900 in Excel).
 */
function excelSerialToIso(serial: number): string {
  if (!Number.isFinite(serial) || serial < 1) return String(serial);
  const adjusted = serial < 60 ? serial + 1 : serial;
  const ms = (adjusted - 25569) * 86_400_000;
  const d = new Date(ms);
  if (Number.isNaN(d.getTime())) return String(serial);
  const yyyy = d.getUTCFullYear().toString().padStart(4, "0");
  const mm = (d.getUTCMonth() + 1).toString().padStart(2, "0");
  const dd = d.getUTCDate().toString().padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

/** True iff this preset id maps to a built-in numFmtId already in the workbook's defaults. */
export function presetNumFmtId(presetId: NumberFormatPresetId): number | undefined {
  const code = NUMBER_FORMAT_PRESETS.find((p) => p.id === presetId)?.code;
  return code ? BUILTIN_NUM_FMT_IDS[code] : undefined;
}
