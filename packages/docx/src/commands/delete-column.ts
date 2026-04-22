import { CommandError, type CommandHandler } from "@officeai/core";
import type { DocxSnapshot, Table, TableGridCol, TableRow } from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import { findTable, withoutRaw } from "./insert-table.js";

/**
 * Remove a column from a typed table — drops the matching `<w:gridCol>`
 * entry and the `tc` at index `at` from every row.
 *
 * Horizontal-merge guard: a column whose cells declare a `<w:gridSpan>`
 * spanning into adjacent columns can't be removed by a per-cell splice
 * without rebalancing every spanned cell. We refuse rather than emit
 * malformed OOXML — callers can split the cell first.
 *
 * Tables that drop to zero columns are illegal in OOXML; for those,
 * dispatch `docx:delete-table` instead.
 */
export interface DeleteColumnPayload {
  readonly tableId: string;
  /** 0-based column index. */
  readonly at: number;
}

export const deleteColumnHandler: CommandHandler<DeleteColumnPayload, DocxSnapshot> = {
  type: "docx:delete-column",
  apply(snapshot, payload) {
    const { tableId, at } = payload;
    if (!tableId) {
      throw new CommandError("unknown-target", "tableId must be a non-empty string");
    }
    const located = findTable(snapshot.root, tableId);
    if (!located) {
      throw new CommandError("unknown-target", `no table with id "${tableId}"`);
    }
    const table = located.table;
    const colCount = table.grid.length;
    if (!Number.isInteger(at) || at < 0 || at >= colCount) {
      throw new CommandError(
        "invalid-position",
        `column index ${at} out of range [0, ${colCount - 1}] for table "${tableId}"`
      );
    }
    if (colCount <= 1) {
      throw new CommandError(
        "invalid-position",
        `cannot delete the last column of table "${tableId}" — dispatch docx:delete-table to remove the whole table`
      );
    }

    // Horizontal-merge guard.
    for (let r = 0; r < table.rows.length; r++) {
      const cell = table.rows[r].cells[at];
      if (cell && cell.properties.gridSpan && cell.properties.gridSpan > 1) {
        throw new CommandError(
          "merged-cell-not-supported",
          `column ${at} carries a horizontal merge in row ${r} of table "${tableId}"`
        );
      }
    }

    const newGrid: TableGridCol[] = table.grid.slice();
    newGrid.splice(at, 1);
    const newRows: TableRow[] = table.rows.map((row) => {
      const cells = row.cells.slice();
      // The cell array can be shorter than the grid (when a preceding
      // gridSpan eats this column); guard the splice accordingly.
      if (at < cells.length) cells.splice(at, 1);
      return { ...row, cells };
    });
    const newTable: Table = withoutRaw({ ...table, grid: newGrid, rows: newRows });
    const nextDoc = located.replace(newTable);
    const next = evolveSnapshot(snapshot, nextDoc, { body: true });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: table.id,
        path: ["body", located.bodyIndex],
        field: "grid",
        summary: `-column at ${at} (table "${tableId}")`,
      }),
    };
  },
};
