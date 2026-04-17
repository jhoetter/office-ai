import { CommandError, type CommandHandler } from "@officeai/core";
import type { DocxSnapshot, Table } from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import { findTable, makeBlankRow, withoutRaw } from "./insert-table.js";

/**
 * Insert a fresh row into a typed table at the given index.
 *
 * The new row carries one cell per existing grid column, each with one
 * empty paragraph and a `<w:tcW>` width inherited from `table.grid`. The
 * row's `header` flag is intentionally left unset so it sits in the body
 * regardless of where it lands; if the brief eventually needs "promote
 * inserted row to header" semantics, callers can follow up with a typed
 * mutation in their own command.
 *
 * Merge protection: insertion is rejected when it would split a vertical
 * merge region. We detect this by looking at the cell at `(at, col)` — if
 * it carries `vMerge: "continue"`, inserting a fresh row before it breaks
 * the merge chain (a continuation cell with no preceding restart is
 * malformed).
 */
export interface InsertRowPayload {
  readonly tableId: string;
  /** 0-based row index. `at === rows.length` appends. */
  readonly at: number;
}

export const insertRowHandler: CommandHandler<InsertRowPayload, DocxSnapshot> = {
  type: "docx:insert-row",
  apply(snapshot, payload, ctx) {
    const { tableId, at } = payload;
    if (!tableId) {
      throw new CommandError("unknown-target", "tableId must be a non-empty string");
    }
    const located = findTable(snapshot.root, tableId);
    if (!located) {
      throw new CommandError("unknown-target", `no table with id "${tableId}"`);
    }
    const table = located.table;
    const rowCount = table.rows.length;
    if (!Number.isInteger(at) || at < 0 || at > rowCount) {
      throw new CommandError(
        "invalid-position",
        `row index ${at} out of range [0, ${rowCount}] for table "${tableId}"`
      );
    }

    if (at < rowCount) {
      const succeeding = table.rows[at];
      for (const cell of succeeding.cells) {
        if (cell.properties.vMerge === "continue") {
          throw new CommandError(
            "merged-cell-not-supported",
            `inserting a row at index ${at} would split a vertical merge region in table "${tableId}"`
          );
        }
      }
    }

    const cols = Math.max(table.grid.length, 1);
    const newRow = makeBlankRow(cols, table.grid, ctx.mintNodeId);
    const newRows = table.rows.slice();
    newRows.splice(at, 0, newRow);
    const newTable: Table = withoutRaw({ ...table, rows: newRows });
    const nextDoc = located.replace(newTable);
    const next = evolveSnapshot(snapshot, nextDoc, { body: true });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-inserted",
        nodeId: newRow.id,
        path: ["body", located.bodyIndex, "rows", at],
        summary: `+row at ${at} (table "${tableId}")`,
      }),
    };
  },
};
