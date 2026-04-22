import { CommandError, type CommandHandler } from "@officeai/core";
import { cellKey } from "../model/refs.js";
import type { Cell, CellValue, Sheet, XlsxSnapshot } from "../model/types.js";
import { recomputeHiddenRows } from "./auto-filter-eval.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type { RemoveDuplicatesPayload } from "./payloads.js";
import { parseRangeRef, resolveSheet } from "./validation.js";

/**
 * `xlsx:remove-duplicates` — Excel's "Data ▸ Remove Duplicates"
 * command. Walks the body rows of `range` (skipping the header row),
 * collapses rows whose values across `keyCols` match a previously-seen
 * row, and emits the survivors back into the original block. Trailing
 * rows are cleared.
 *
 * Behavioural mirrors of Excel's dialog:
 *   - The first row of the range is the header and never moves.
 *   - When `keyCols` is empty we fall back to "every column in the
 *     range" — same default Excel surfaces when the user clicks
 *     "Select All" in the dedupe dialog.
 *   - Comparison is value-equality with case-insensitive string
 *     matching (matches Excel: "ALPHA" and "alpha" collapse).
 *   - Numbers, booleans and errors compare strictly; `null`/empty
 *     cells collapse together.
 *   - Surviving rows preserve their cell formats — only the row
 *     positions change. Cells outside `range` are untouched, even
 *     when on a surviving / removed row, matching Excel's
 *     "selection-only" semantics.
 *
 * The handler returns a {@link RemoveDuplicatesResultMeta} via the
 * diff.summary so callers can surface "N duplicates removed" toasts.
 */
export const removeDuplicatesHandler: CommandHandler<RemoveDuplicatesPayload, XlsxSnapshot> = {
  type: "xlsx:remove-duplicates",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    const range = parseRangeRef(payload.range);
    const colSpan = range.end.col - range.start.col;
    if (range.end.row <= range.start.row) {
      // Header-only or empty range — nothing to dedupe.
      const noop = evolveSnapshot(snapshot, snapshot.root, {});
      return { next: noop, diff: buildDiff(snapshot.revision, noop.revision, []) };
    }

    const keyOffsets = (payload.keyCols && payload.keyCols.length > 0
      ? payload.keyCols
      : Array.from({ length: colSpan + 1 }, (_, i) => i)
    ).slice();
    for (const off of keyOffsets) {
      if (!Number.isInteger(off) || off < 0 || off > colSpan) {
        throw new CommandError(
          "invalid-payload",
          `keyCols offset ${off} out of range; range spans 0..${colSpan}`
        );
      }
    }

    const headerRow = range.start.row;
    const firstBody = headerRow + 1;
    const lastBody = range.end.row;

    type RowSnapshot = { cells: Map<number, Cell> };
    const rows: RowSnapshot[] = [];
    for (let r = firstBody; r <= lastBody; r++) {
      const cells = new Map<number, Cell>();
      for (let c = range.start.col; c <= range.end.col; c++) {
        const cell = sheet.cells.get(cellKey(r, c));
        if (cell) cells.set(c, cell);
      }
      rows.push({ cells });
    }

    const seen = new Set<string>();
    const survivors: RowSnapshot[] = [];
    for (const row of rows) {
      const key = keyOffsets
        .map((off) => normaliseValue(row.cells.get(range.start.col + off)?.value ?? null))
        .join("\u0001");
      if (seen.has(key)) continue;
      seen.add(key);
      survivors.push(row);
    }
    const removed = rows.length - survivors.length;
    if (removed === 0) {
      const noop = evolveSnapshot(snapshot, snapshot.root, {});
      return { next: noop, diff: buildDiff(snapshot.revision, noop.revision, []) };
    }

    const cells = new Map(sheet.cells);
    for (let r = firstBody; r <= lastBody; r++) {
      for (let c = range.start.col; c <= range.end.col; c++) {
        cells.delete(cellKey(r, c));
      }
    }
    for (let i = 0; i < survivors.length; i++) {
      const targetRow = firstBody + i;
      const src = survivors[i]!.cells;
      for (const [c, cell] of src) {
        cells.set(cellKey(targetRow, c), { ...cell, row: targetRow, col: c });
      }
    }

    const nextSheet: Sheet = { ...sheet, cells };
    const finalSheet: Sheet = nextSheet.autoFilter
      ? {
          ...nextSheet,
          hiddenRows: recomputeHiddenRows(nextSheet, snapshot.root.styles, nextSheet.autoFilter),
        }
      : nextSheet;

    const nextWorkbook = replaceSheet(snapshot.root, finalSheet);
    const next = evolveSnapshot(snapshot, nextWorkbook, { sheets: [sheet.partPath] });
    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, [
        {
          kind: "node-updated",
          nodeId: sheet.id,
          path: ["sheets", sheet.index, "cells"],
          field: "remove-duplicates",
          summary: `${sheet.name} ${payload.range}: removed ${removed} duplicate row${removed === 1 ? "" : "s"} (${survivors.length} kept)`,
        },
      ]),
    };
  },
};

function normaliseValue(v: CellValue | null): string {
  if (v === null || v === undefined || v === "") return "<blank>";
  if (typeof v === "string") return `s:${v.toLocaleLowerCase()}`;
  if (typeof v === "number") return `n:${v}`;
  if (typeof v === "boolean") return `b:${v ? 1 : 0}`;
  return `e:${String(v)}`;
}
