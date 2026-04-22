import { CommandError, type CommandHandler } from "@officeai/core";
import type { DocxSnapshot } from "../model/types.js";
import { buildDiff, evolveSnapshot, removeBlocks } from "./helpers.js";
import { findTable } from "./insert-table.js";

/**
 * Remove a top-level table from the body. Nested tables (tables inside
 * cells of other tables) are out of scope today — deleting one would
 * require a typed splice inside the parent cell, which the brief
 * defers to a follow-up.
 */
export interface DeleteTablePayload {
  readonly tableId: string;
}

export const deleteTableHandler: CommandHandler<DeleteTablePayload, DocxSnapshot> = {
  type: "docx:delete-table",
  apply(snapshot, payload) {
    const { tableId } = payload;
    if (!tableId) {
      throw new CommandError("unknown-target", "tableId must be a non-empty string");
    }
    const located = findTable(snapshot.root, tableId);
    if (!located) {
      throw new CommandError("unknown-target", `no table with id "${tableId}"`);
    }
    if (located.ancestorIds.length > 0) {
      throw new CommandError(
        "invalid-position",
        `cannot delete nested table "${tableId}" — only top-level tables are supported`
      );
    }
    const idx = located.bodyIndex;
    const block = snapshot.root.body[idx];
    if (!block || block.kind !== "table" || block.id !== tableId) {
      throw new CommandError("unknown-target", `body[${idx}] is not the requested table`);
    }
    const nextDoc = removeBlocks(snapshot.root, idx, idx + 1);
    const next = evolveSnapshot(snapshot, nextDoc, { body: true });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-deleted",
        nodeId: tableId,
        path: ["body", idx],
        summary: `-table "${tableId}"`,
      }),
    };
  },
};
