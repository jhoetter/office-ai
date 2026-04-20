import type { DocxAgent } from "@officeai/docx";
import type { PptxAgent } from "@officeai/pptx";
import type { XlsxClipboardSnapshot } from "@officeai/xlsx";

import { applyXlsxRangeToDocx } from "./applyXlsxRangeToDocx";
import { applyXlsxRangeToPptx } from "./applyXlsxRangeToPptx";

/**
 * Three modes — mirrors PowerPoint / Word's "Paste Special" submenu:
 *   - `materialized`: a typed table (`<w:tbl>` / `<a:tbl>`). Best
 *     fidelity inside the host document; cells stay editable as
 *     cells.
 *   - `live`: an OLE-embedded `.xlsx` package. Office activates Excel
 *     on double-click; the embedded bytes round-trip with the
 *     document.
 *   - `chart`: a typed chart (`<c:chartSpace>`) using the snapshot's
 *     first row as series names and first column as categories.
 *     Office's "Edit Data" round-trips against a real workbook.
 */
export type XlsxEmbedMode = "materialized" | "live" | "chart";

/**
 * Discriminated target for the cross-format dispatcher. Carries the
 * format-specific anchor (paragraph index for DOCX, slide index for
 * PPTX) alongside the agent the helper should mutate.
 */
export type XlsxEmbedTarget =
  | { readonly kind: "docx"; readonly agent: DocxAgent; readonly paragraphIndex: number }
  | { readonly kind: "pptx"; readonly agent: PptxAgent; readonly slideIndex: number };

/**
 * Single dispatcher used by the editor paste/drop handlers and the
 * "Insert from xlsx" actions. Delegates to the format-specific
 * helper so the per-format wiring stays in one place.
 */
export async function applyXlsxEmbed(args: {
  readonly target: XlsxEmbedTarget;
  readonly snapshot: XlsxClipboardSnapshot;
  readonly mode?: XlsxEmbedMode;
}): Promise<void> {
  const { target, snapshot, mode } = args;
  switch (target.kind) {
    case "docx":
      await applyXlsxRangeToDocx({
        agent: target.agent,
        snapshot,
        paragraphIndex: target.paragraphIndex,
        ...(mode ? { mode } : {}),
      });
      return;
    case "pptx":
      await applyXlsxRangeToPptx({
        agent: target.agent,
        snapshot,
        slideIndex: target.slideIndex,
        ...(mode ? { mode } : {}),
      });
      return;
    default: {
      const _exhaustive: never = target;
      throw new Error(`unhandled embed target: ${JSON.stringify(_exhaustive)}`);
    }
  }
}

/**
 * Project a clipboard snapshot onto a 2D grid of primitive values
 * suitable for both `docx:insert-spreadsheet` and `pptx:insert-spreadsheet`.
 * Formula cells fall back to their evaluated `value` (Office shows
 * the value, not the formula, until the user activates Excel and
 * double-clicks). Errors surface as their `code` (e.g. `"#REF!"`)
 * so the cell isn't silently empty in the preview.
 */
export function snapshotToGrid(
  snapshot: XlsxClipboardSnapshot
): ReadonlyArray<ReadonlyArray<string | number | null>> {
  return snapshot.cells.map((row) =>
    row.map((cell) => {
      if (!cell) return null;
      const v = cell.value;
      if (v === null || v === undefined) return null;
      if (typeof v === "string") return v;
      if (typeof v === "number") return Number.isFinite(v) ? v : null;
      if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
      return v.code;
    })
  );
}

export interface ChartProjection {
  readonly categories: ReadonlyArray<string>;
  readonly series: ReadonlyArray<{
    readonly name?: string;
    readonly values: ReadonlyArray<number>;
  }>;
}

/**
 * Project a clipboard snapshot onto an Office-style chart spec:
 *   - Row 0, columns ≥ 1 → series names
 *   - Column 0, rows ≥ 1 → category labels
 *   - (r, c) for r ≥ 1, c ≥ 1 → numeric value for category (r-1)
 *     of series (c-1). Non-numeric / blank cells are coerced to 0
 *     so the bar/line renderer doesn't drop the data point silently.
 *
 * Returns `null` for grids too small to plot (< 2 rows or < 2
 * columns) — callers fall back to materialised tables in that
 * case so the paste isn't a no-op.
 */
export function snapshotToChartSpec(snapshot: XlsxClipboardSnapshot): ChartProjection | null {
  if (snapshot.height < 2 || snapshot.width < 2) return null;
  const grid = snapshotToGrid(snapshot);
  const headerRow = grid[0]!;
  const categories: string[] = [];
  for (let r = 1; r < snapshot.height; r++) {
    categories.push(stringify(grid[r]?.[0]));
  }
  const series: Array<{ name?: string; values: number[] }> = [];
  for (let c = 1; c < snapshot.width; c++) {
    const name = stringify(headerRow[c]);
    const values: number[] = [];
    for (let r = 1; r < snapshot.height; r++) {
      values.push(toNumeric(grid[r]?.[c]));
    }
    series.push(name ? { name, values } : { values });
  }
  return { categories, series };
}

function stringify(v: string | number | null | undefined): string {
  if (v === null || v === undefined) return "";
  return String(v);
}

function toNumeric(v: string | number | null | undefined): number {
  if (typeof v === "number" && Number.isFinite(v)) return v;
  if (typeof v === "string") {
    const n = Number(v);
    return Number.isFinite(n) ? n : 0;
  }
  return 0;
}
