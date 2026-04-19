/**
 * Selection model for the virtualized grid.
 *
 * `anchor` is where the user started the selection (single click, or
 * the "fixed" end of a Shift-extend); `focus` is the moving end. Both
 * are inclusive and 0-based. A single-cell selection has
 * `anchor === focus` (per-coordinate equality, not reference).
 *
 * The grid renders a single bounding-box marquee around
 * `normalizeSelection(sel)` so the highlight reads like Excel even
 * when the user drags up/left.
 */

import { formatA1, formatRange, type CellRange } from "@officeai/xlsx";

export interface CellPos {
  readonly row: number;
  readonly col: number;
}

export interface Selection {
  readonly anchor: CellPos;
  readonly focus: CellPos;
}

export interface NormalizedSelection {
  readonly r0: number;
  readonly c0: number;
  readonly r1: number;
  readonly c1: number;
}

export function singleSelection(pos: CellPos): Selection {
  return { anchor: pos, focus: pos };
}

export function normalizeSelection(sel: Selection): NormalizedSelection {
  const r0 = Math.min(sel.anchor.row, sel.focus.row);
  const r1 = Math.max(sel.anchor.row, sel.focus.row);
  const c0 = Math.min(sel.anchor.col, sel.focus.col);
  const c1 = Math.max(sel.anchor.col, sel.focus.col);
  return { r0, c0, r1, c1 };
}

export function isSingle(sel: Selection): boolean {
  return sel.anchor.row === sel.focus.row && sel.anchor.col === sel.focus.col;
}

export function containsCell(sel: Selection, row: number, col: number): boolean {
  const n = normalizeSelection(sel);
  return row >= n.r0 && row <= n.r1 && col >= n.c0 && col <= n.c1;
}

export function selectionToRange(sel: Selection): CellRange {
  const n = normalizeSelection(sel);
  return { start: { row: n.r0, col: n.c0 }, end: { row: n.r1, col: n.c1 } };
}

/**
 * Render the selection as Excel would: `A1` for single cells, `A1:C3`
 * for ranges. Used by the cell-ref pill in the formula bar.
 */
export function formatSelection(sel: Selection): string {
  if (isSingle(sel)) return formatA1({ row: sel.anchor.row, col: sel.anchor.col });
  return formatRange(selectionToRange(sel));
}

/**
 * C13 — Multi-area selection helpers.
 *
 * Excel lets you Ctrl-click (Cmd-click on Mac) to add disjoint
 * rectangles to the selection. We model this as the existing single
 * `selection` (the *active* area, where the formula-bar editing,
 * fill-handle, paste destination etc. live) plus an array of
 * `extraAreas`. The union is just `[...extraAreas, selection]`.
 */
export function allAreas(
  selection: Selection | null,
  extras: ReadonlyArray<Selection>
): ReadonlyArray<Selection> {
  return selection ? [...extras, selection] : extras;
}

export function areasContainCell(areas: ReadonlyArray<Selection>, row: number, col: number): boolean {
  for (const a of areas) {
    if (containsCell(a, row, col)) return true;
  }
  return false;
}

export function selectionEquals(a: Selection, b: Selection): boolean {
  return (
    a.anchor.row === b.anchor.row &&
    a.anchor.col === b.anchor.col &&
    a.focus.row === b.focus.row &&
    a.focus.col === b.focus.col
  );
}

/**
 * Returns true if `inner` is a subset (rectangle-wise) of `outer`.
 * Used when Ctrl-clicking inside an existing area to *deselect* it.
 */
export function selectionCovers(outer: Selection, inner: Selection): boolean {
  const o = normalizeSelection(outer);
  const i = normalizeSelection(inner);
  return i.r0 >= o.r0 && i.r1 <= o.r1 && i.c0 >= o.c0 && i.c1 <= o.c1;
}

/**
 * Render the union as Excel does in the Name Box: comma-joined A1
 * refs, active area last. Single-cell areas collapse to `A1` form.
 */
export function formatAreas(selection: Selection | null, extras: ReadonlyArray<Selection>): string {
  return allAreas(selection, extras).map(formatSelection).join(",");
}

/**
 * Iterate every (row, col) covered by the union *exactly once*.
 * Streams via a callback so we don't materialise enormous arrays for
 * Ctrl+A-style selections on a 1M × 16K sheet.
 *
 * Cells in overlapping areas are visited only once thanks to a
 * dedup `Set<string>` keyed by `r:c`. The set is local to the call,
 * not retained.
 *
 * SAFETY: For column-header / row-header / Ctrl+A selections the
 * naive (r, c) walk would be 1M × 16K = ~16 billion iterations and
 * lock the main thread. Callers that don't need to visit empty
 * cells (e.g. aggregates, clearing) should use
 * {@link forEachUnionSparseCell} instead, which iterates a sparse
 * `cells` map and only visits populated coordinates inside the
 * selection. This helper bails out with a no-op if the union span
 * exceeds {@link DENSE_CELL_BUDGET} so a misuse can never freeze
 * the UI.
 */
const DENSE_CELL_BUDGET = 1_000_000;

export function forEachUnionCell(
  areas: ReadonlyArray<Selection>,
  fn: (row: number, col: number) => void
): void {
  if (unionSpanUpperBound(areas) > DENSE_CELL_BUDGET) {
    if (typeof console !== "undefined") {
      console.warn(
        "forEachUnionCell: refusing to walk %d cells; use forEachUnionSparseCell for whole-row/col selections.",
        unionSpanUpperBound(areas)
      );
    }
    return;
  }
  const seen = new Set<string>();
  for (const a of areas) {
    const n = normalizeSelection(a);
    for (let r = n.r0; r <= n.r1; r++) {
      for (let c = n.c0; c <= n.c1; c++) {
        const k = `${r}:${c}`;
        if (seen.has(k)) continue;
        seen.add(k);
        fn(r, c);
      }
    }
  }
}

/**
 * Upper bound on the number of cells covered by the union of `areas`,
 * counting overlapping coordinates more than once (cheap rectangle
 * sum). Used as a fast guardrail before any dense (r, c) walk.
 */
export function unionSpanUpperBound(areas: ReadonlyArray<Selection>): number {
  let total = 0;
  for (const a of areas) {
    const n = normalizeSelection(a);
    total += (n.r1 - n.r0 + 1) * (n.c1 - n.c0 + 1);
    if (total >= Number.MAX_SAFE_INTEGER) return Number.MAX_SAFE_INTEGER;
  }
  return total;
}

/**
 * Iterate the populated cells of a sparse cell map that fall inside
 * the union of `areas`. This is O(numPopulatedCells) regardless of
 * how big the selection rectangle is — the only safe walk for a
 * "select entire column / row / sheet" gesture.
 */
export function forEachUnionSparseCell<T extends { readonly row: number; readonly col: number }>(
  cells: Iterable<T>,
  areas: ReadonlyArray<Selection>,
  fn: (cell: T) => void
): void {
  if (areas.length === 0) return;
  for (const cell of cells) {
    if (areasContainCell(areas, cell.row, cell.col)) fn(cell);
  }
}
