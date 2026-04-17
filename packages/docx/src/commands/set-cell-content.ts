import { CommandError, type CommandHandler } from "@officeai/core";
import type { BlockNode, DocxSnapshot, Table, TableCell, TableRow } from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import { findTable, withoutRaw } from "./insert-table.js";
import type { SetCellContentPayload } from "./payloads.js";

/**
 * Replace the body of a single typed cell wholesale.
 *
 * The handler intentionally rejects writes that would land inside a merged
 * region — the brief explicitly defers span reflow to a later session. We
 * only check the target cell itself for `vMerge: "continue"`; horizontal
 * spans (`gridSpan`) on the target cell are fine, the cell is just wider.
 *
 * Cycle protection: `payload.content` may not contain a `Table` whose `id`
 * collides with any of the target table's ancestor ids. This catches the
 * common mistake of trying to nest a table inside one of its own ancestors.
 */
export const setCellContentHandler: CommandHandler<SetCellContentPayload, DocxSnapshot> = {
  type: "docx:set-cell-content",
  apply(snapshot, payload) {
    const { tableId, row, col, content } = payload;
    if (!tableId) {
      throw new CommandError("unknown-target", "tableId must be a non-empty string");
    }
    const located = findTable(snapshot.root, tableId);
    if (!located) {
      throw new CommandError("unknown-target", `no table with id "${tableId}"`);
    }
    const table = located.table;
    if (!Number.isInteger(row) || row < 0 || row >= table.rows.length) {
      throw new CommandError(
        "unknown-target",
        `row index ${row} out of range (table "${tableId}" has ${table.rows.length} rows)`
      );
    }
    const targetRow = table.rows[row];
    if (!Number.isInteger(col) || col < 0 || col >= targetRow.cells.length) {
      throw new CommandError(
        "unknown-target",
        `col index ${col} out of range (row ${row} of table "${tableId}" has ${targetRow.cells.length} cells)`
      );
    }
    const targetCell = targetRow.cells[col];
    if (targetCell.properties.vMerge === "continue") {
      throw new CommandError(
        "merged-cell-not-supported",
        `cell at (row=${row}, col=${col}) of table "${tableId}" is a vertical-merge continuation; reflow not supported this round`
      );
    }
    const forbiddenIds = new Set<string>([tableId, ...located.ancestorIds]);
    assertNoTableCycle(content, forbiddenIds);

    const newBody: BlockNode[] =
      content.length > 0 ? content.slice() : [makeEmptyParagraphFallback(targetCell)];
    const newCell: TableCell = { ...targetCell, body: newBody };
    const newCells = targetRow.cells.slice();
    newCells[col] = newCell;
    const newRow: TableRow = { ...targetRow, cells: newCells };
    const newRows = table.rows.slice();
    newRows[row] = newRow;
    const newTable: Table = withoutRaw({ ...table, rows: newRows });
    const nextDoc = located.replace(newTable);
    const next = evolveSnapshot(snapshot, nextDoc, { body: true });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: targetCell.id,
        path: ["body", located.bodyIndex, "rows", row, "cells", col],
        field: "body",
        summary: `cell (${row},${col}) ← ${content.length} block${content.length === 1 ? "" : "s"}`,
      }),
    };
  },
};

/**
 * Recursive sweep for `Table` nodes whose `id` collides with the target's
 * ancestor chain. Throws `unknown-target` (not `invalid-payload`) so the
 * error code matches the brief's wording for "attempt to write a `Table`
 * that contains itself".
 */
function assertNoTableCycle(blocks: ReadonlyArray<BlockNode>, forbidden: ReadonlySet<string>): void {
  for (const block of blocks) {
    if (block.kind !== "table") continue;
    if (forbidden.has(block.id)) {
      throw new CommandError(
        "unknown-target",
        `cell content contains table id "${block.id}" which is the target or an ancestor table — would create a cycle`
      );
    }
    for (const row of block.rows) {
      for (const cell of row.cells) {
        assertNoTableCycle(cell.body, forbidden);
      }
    }
  }
}

/**
 * If the caller passes an empty `content` array we still need to leave the
 * cell with at least one block — OOXML requires every `<w:tc>` to contain
 * a paragraph. We synthesize a fresh empty paragraph carrying a fresh id
 * derived from the target cell's id (best-effort; the agent will typically
 * pass a non-empty content array so this is the rare edge case path).
 */
function makeEmptyParagraphFallback(targetCell: TableCell): BlockNode {
  const baseId = `${targetCell.id}-empty`;
  return {
    kind: "paragraph",
    id: baseId,
    properties: {},
    children: [{ kind: "run", id: `${baseId}-r0`, properties: {}, children: [] }],
  };
}
