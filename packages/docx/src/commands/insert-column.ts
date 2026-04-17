import { CommandError, type CommandHandler, type IdMinter } from "@officeai/core";
import type { DocxSnapshot, Table, TableCell, TableGridCol, TableRow } from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import { findTable, makeBlankCell, withoutRaw } from "./insert-table.js";

/**
 * Add one column to a typed table.
 *
 * `at` is the 0-based column index where the new column lands. `at ===
 * grid.length` appends at the right edge. The handler:
 *
 * 1. Inserts a new `<w:gridCol>` at the same index in `table.grid`.
 * 2. Inserts a fresh `<w:tc>` (single empty paragraph) at the same index in
 *    every row.
 *
 * Merge protection: every row must have a clean column boundary at `at` —
 * i.e. the running sum of cell `gridSpan` values up to (but not including)
 * the cell at `at` must equal `at`. If a cell with `gridSpan > 1` straddles
 * `at`, we reject with `merged-cell-not-supported` rather than corrupt the
 * span. Boundary insertions (`at === 0` or `at === grid.length`) are always
 * safe because they sit at the table's own edges.
 *
 * Default width: when `width` is omitted, we equal-split the existing
 * declared column widths. If the existing grid has no widths at all, we
 * fall back to 1000 twips (Word treats this as "auto" once it opens the
 * file, so the visible default rendering remains reasonable).
 */
export interface InsertColumnPayload {
  readonly tableId: string;
  readonly at: number;
  /** Column width in twips. */
  readonly width?: number;
}

export const insertColumnHandler: CommandHandler<InsertColumnPayload, DocxSnapshot> = {
  type: "docx:insert-column",
  apply(snapshot, payload, ctx) {
    const { tableId, at, width } = payload;
    if (!tableId) {
      throw new CommandError("unknown-target", "tableId must be a non-empty string");
    }
    const located = findTable(snapshot.root, tableId);
    if (!located) {
      throw new CommandError("unknown-target", `no table with id "${tableId}"`);
    }
    const table = located.table;
    const cols = table.grid.length;
    if (!Number.isInteger(at) || at < 0 || at > cols) {
      throw new CommandError(
        "invalid-position",
        `column index ${at} out of range [0, ${cols}] for table "${tableId}"`
      );
    }

    // Boundary-safety check: every row must split cleanly at `at`. Skip when
    // we're inserting at one of the table's own edges (always safe).
    if (at !== 0 && at !== cols) {
      for (let r = 0; r < table.rows.length; r++) {
        const row = table.rows[r];
        if (!hasBoundaryAt(row, at)) {
          throw new CommandError(
            "merged-cell-not-supported",
            `row ${r} of table "${tableId}" has a horizontal-span cell crossing column index ${at}; reflow not supported this round`
          );
        }
      }
    }

    const colWidth = width ?? defaultColumnWidth(table.grid);
    const newGrid = insertGridCol(table.grid, at, colWidth);
    const newRows: TableRow[] = table.rows.map((row) => insertCellAt(row, at, colWidth, ctx.mintNodeId));
    const newTable: Table = withoutRaw({ ...table, grid: newGrid, rows: newRows });
    const nextDoc = located.replace(newTable);
    const next = evolveSnapshot(snapshot, nextDoc, { body: true });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-inserted",
        nodeId: newTable.id,
        path: ["body", located.bodyIndex, "grid", at],
        summary: `+column at ${at} (w=${colWidth}, table "${tableId}")`,
      }),
    };
  },
};

function hasBoundaryAt(row: TableRow, at: number): boolean {
  let consumed = 0;
  for (const cell of row.cells) {
    if (consumed === at) return true;
    const span = cell.properties.gridSpan ?? 1;
    if (consumed < at && consumed + span > at) {
      return false;
    }
    consumed += span;
  }
  // Reached the end without finding a boundary mid-row → safe iff at sits
  // exactly at the row's right edge (handled by callers, but guard anyway).
  return consumed === at;
}

function insertGridCol(grid: ReadonlyArray<TableGridCol>, at: number, width: number): TableGridCol[] {
  const next = grid.slice();
  next.splice(at, 0, { w: width });
  return next;
}

function insertCellAt(row: TableRow, at: number, width: number, mintNodeId: IdMinter): TableRow {
  const newCell: TableCell = makeBlankCell(width, mintNodeId);
  // Map column index → cells array index by walking gridSpans. At a clean
  // boundary the index where the new cell lands is the count of preceding
  // cells whose cumulative span ≤ `at`.
  const cells = row.cells.slice();
  const insertIndex = cellIndexForColumn(row, at);
  cells.splice(insertIndex, 0, newCell);
  return { ...row, cells };
}

function cellIndexForColumn(row: TableRow, at: number): number {
  let consumed = 0;
  for (let i = 0; i < row.cells.length; i++) {
    if (consumed === at) return i;
    consumed += row.cells[i].properties.gridSpan ?? 1;
  }
  return row.cells.length;
}

/**
 * Equal-split the existing declared widths. When the grid has zero declared
 * widths we fall back to a sensible 1000-twip default (~0.69") that
 * Word/LibreOffice both render reasonably.
 */
function defaultColumnWidth(grid: ReadonlyArray<TableGridCol>): number {
  const declared = grid.filter((g): g is { w: number } => g.w !== undefined);
  if (declared.length === 0) return 1000;
  const total = declared.reduce((sum, g) => sum + g.w, 0);
  // Equal split across the new (cols + 1) layout — keeps total table width
  // roughly constant.
  return Math.max(100, Math.round(total / (declared.length + 1)));
}
