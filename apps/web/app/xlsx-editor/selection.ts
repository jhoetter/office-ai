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
