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
import type { CellValue, EffectiveStyle, StyleColor, StyleTable } from "@officeai/xlsx";
import { flattenCellXf } from "@officeai/xlsx";

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
  if (effective.fill.kind === "pattern" && effective.fill.patternType === "solid") {
    const fg = renderColor(effective.fill.fgColor);
    if (fg) css.background = fg;
  }
  if (effective.alignment?.horizontal) {
    switch (effective.alignment.horizontal) {
      case "left":
      case "center":
      case "right":
      case "justify":
        css.justifyContent = horizontalToJustify(effective.alignment.horizontal);
        css.textAlign = effective.alignment.horizontal;
        break;
      default:
        break;
    }
  }
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
