import { parseRange } from "@officeai/xlsx";

/**
 * Geometry of a chart's `dataRange` in 0-based row/col coordinates.
 * Always normalized so `r1 <= r2` and `c1 <= c2`.
 */
export interface ChartRangeShape {
  readonly r1: number;
  readonly r2: number;
  readonly c1: number;
  readonly c2: number;
}

/**
 * Parse a chart's stored A1 dataRange or a dialog input string into
 * the rectangle the renderer uses. Delegates to the canonical
 * `parseRange` from `@officeai/xlsx` so we accept the same set of
 * inputs as `addChartHandler` (including `$`-anchored refs and
 * single-cell ranges). Strips an optional sheet prefix because chart
 * `dataRange`s are always sheet-local in the model, but a user might
 * have typed `Sheet1!A1:B5` in the dialog.
 */
export function parseChartRangeShape(range: string): ChartRangeShape | null {
  const trimmed = range.trim();
  if (!trimmed) return null;
  const bang = trimmed.lastIndexOf("!");
  const refOnly = bang === -1 ? trimmed : trimmed.slice(bang + 1);
  try {
    const { start, end } = parseRange(refOnly);
    return {
      r1: Math.min(start.row, end.row),
      r2: Math.max(start.row, end.row),
      c1: Math.min(start.col, end.col),
      c2: Math.max(start.col, end.col),
    };
  } catch {
    return null;
  }
}

/**
 * Strip Excel's `$` anchor markers from an A1 ref so the stored
 * `dataRange` is the bare form the renderer expects. We keep an
 * optional sheet prefix intact (`Sheet1!A1:B5` → `Sheet1!A1:B5`)
 * because the model treats sheet-local refs verbatim.
 */
export function normalizeRangeForStorage(range: string): string {
  return range.trim().toUpperCase().replace(/\$/g, "");
}

/**
 * Pick sensible header / category defaults from a freshly opened
 * selection. A single-column or single-row selection almost never
 * has both a label header *and* a label gutter — defaulting both
 * toggles to `true` collapses the value series to zero and shows
 * the "No data in selected range" empty state, which is the bug
 * that motivated this. Multi-row + multi-col selections keep the
 * Excel-parity defaults (header row + category column on).
 */
export function pickToggleDefaults(shape: ChartRangeShape | null): {
  readonly hasHeaderRow: boolean;
  readonly hasCategoryColumn: boolean;
} {
  if (!shape) return { hasHeaderRow: true, hasCategoryColumn: true };
  const singleRow = shape.r1 === shape.r2;
  const singleCol = shape.c1 === shape.c2;
  if (singleRow || singleCol) return { hasHeaderRow: false, hasCategoryColumn: false };
  return { hasHeaderRow: true, hasCategoryColumn: true };
}

/**
 * Result of validating a `(range, hasHeaderRow, hasCategoryColumn)`
 * tuple against the chart renderer's expectations. Used by the
 * Insert/Edit dialog to disable submit + surface a hint when the
 * combination would render an empty chart.
 */
export type ChartShapeValidation =
  | { readonly kind: "empty" }
  | { readonly kind: "invalid" }
  | { readonly kind: "single-cell" }
  | { readonly kind: "no-values"; readonly axis: "row" | "column" }
  | { readonly kind: "ok" };

/**
 * Mirror the value-series / body-row math in
 * `ChartOverlay.extractSeries` so the dialog can warn the user
 * *before* submitting — never let the user produce a chart that
 * silently renders the "No data in selected range" empty state.
 */
export function validateChartShape(
  range: string,
  shape: ChartRangeShape | null,
  hasHeaderRow: boolean,
  hasCategoryColumn: boolean
): ChartShapeValidation {
  if (!range.trim()) return { kind: "empty" };
  if (!shape) return { kind: "invalid" };
  if (shape.r1 === shape.r2 && shape.c1 === shape.c2) return { kind: "single-cell" };
  const valStart = hasCategoryColumn ? shape.c1 + 1 : shape.c1;
  const bodyStart = hasHeaderRow ? shape.r1 + 1 : shape.r1;
  if (valStart > shape.c2) return { kind: "no-values", axis: "column" };
  if (bodyStart > shape.r2) return { kind: "no-values", axis: "row" };
  return { kind: "ok" };
}
