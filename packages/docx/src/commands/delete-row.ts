import { CommandError, type CommandHandler } from "@officeai/core";
import type { DocxSnapshot, Table } from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import { findTable, withoutRaw } from "./insert-table.js";

/**
 * Remove one row from a typed table.
 *
 * Mirrors {@link insertRowHandler}'s vertical-merge protection: removing
 * a row whose first cell is the *anchor* of a vMerge (i.e. carries
 * `vMerge: "restart"` while a successor row carries `vMerge: "continue"`)
 * would leave the continuation cell dangling, which OOXML rejects.
 *
 * Tables that drop to zero rows are illegal in OOXML (Word refuses to
 * open them). Callers wanting to remove the table entirely should
 * dispatch `docx:delete-table` instead.
 */
export interface DeleteRowPayload {
  readonly tableId: string;
  /** 0-based row index. */
  readonly at: number;
}

export const deleteRowHandler: CommandHandler<DeleteRowPayload, DocxSnapshot> = {
  type: "docx:delete-row",
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
    const rowCount = table.rows.length;
    if (!Number.isInteger(at) || at < 0 || at >= rowCount) {
      throw new CommandError(
        "invalid-position",
        `row index ${at} out of range [0, ${rowCount - 1}] for table "${tableId}"`
      );
    }
    if (rowCount <= 1) {
      throw new CommandError(
        "invalid-position",
        `cannot delete the last row of table "${tableId}" — dispatch docx:delete-table to remove the whole table`
      );
    }

    // Vertical-merge guard: deleting a row that anchors a merge
    // chain leaves dangling continuations.
    const removed = table.rows[at];
    for (let c = 0; c < removed.cells.length; c++) {
      if (removed.cells[c].properties.vMerge === "restart") {
        const next = table.rows[at + 1];
        if (next && next.cells[c]?.properties.vMerge === "continue") {
          throw new CommandError(
            "merged-cell-not-supported",
            `deleting row ${at} would orphan a vertical-merge continuation in table "${tableId}"`
          );
        }
      }
    }

    const newRows = table.rows.slice();
    newRows.splice(at, 1);
    const newTable: Table = withoutRaw({ ...table, rows: newRows });
    const nextDoc = located.replace(newTable);
    const next = evolveSnapshot(snapshot, nextDoc, { body: true });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-deleted",
        nodeId: removed.id,
        path: ["body", located.bodyIndex, "rows", at],
        summary: `-row at ${at} (table "${tableId}")`,
      }),
    };
  },
};
