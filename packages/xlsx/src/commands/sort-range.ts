import { CommandError, type CommandHandler } from "@officeai/core";
import { cellKey } from "../model/refs.js";
import type { Cell, CellValue, Sheet, XlsxSnapshot } from "../model/types.js";
import { recomputeHiddenRows } from "./auto-filter-eval.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type { SortRangePayload } from "./payloads.js";
import { parseRangeRef, resolveSheet } from "./validation.js";

/**
 * `xlsx:sort-range` — sort the body rows inside `range` by one
 * column. The first row of the range is treated as the header and
 * never moves (matches the dropdown's "Sort A→Z" / "Sort Z→A").
 *
 * Implementation notes:
 *  - Sort is stable.
 *  - Numbers sort numerically; strings sort case-insensitively
 *    (locale-aware); booleans (FALSE < TRUE) and errors sort last.
 *  - Whole rows move — not just the sort column — so adjacent data
 *    stays attached, mirroring Excel's default expand-selection sort.
 *  - We rewrite cells in place; `Sheet.cells` gets a fresh map and
 *    the dirty bit fires on the sheet part. AutoFilter hidden rows
 *    are recomputed since their row indices may now point at
 *    different values.
 */
export const sortRangeHandler: CommandHandler<SortRangePayload, XlsxSnapshot> = {
  type: "xlsx:sort-range",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    const range = parseRangeRef(payload.range);
    const colSpan = range.end.col - range.start.col;
    if (
      !Number.isInteger(payload.sortBy.colId) ||
      payload.sortBy.colId < 0 ||
      payload.sortBy.colId > colSpan
    ) {
      throw new CommandError(
        "invalid-payload",
        `sortBy.colId ${payload.sortBy.colId} out of range; range spans 0..${colSpan}`
      );
    }
    if (range.start.row === range.end.row) {
      const noop = evolveSnapshot(snapshot, snapshot.root, {});
      return { next: noop, diff: buildDiff(snapshot.revision, noop.revision, []) };
    }

    const headerRow = range.start.row;
    const firstBody = headerRow + 1;
    const lastBody = range.end.row;
    const sortCol = range.start.col + payload.sortBy.colId;
    const order = payload.sortBy.order === "desc" ? -1 : 1;

    // Snapshot each body row's cells (only those inside the range).
    type RowSnapshot = { key: CellValue; cells: Map<number, Cell> };
    const rowSnapshots: RowSnapshot[] = [];
    for (let r = firstBody; r <= lastBody; r++) {
      const cells = new Map<number, Cell>();
      for (let c = range.start.col; c <= range.end.col; c++) {
        const cell = sheet.cells.get(cellKey(r, c));
        if (cell) cells.set(c, cell);
      }
      rowSnapshots.push({ key: cells.get(sortCol)?.value ?? null, cells });
    }

    // Stable sort — pair each row with its original index.
    const indexed = rowSnapshots.map((row, idx) => ({ row, idx }));
    indexed.sort((a, b) => {
      const cmp = compareValues(a.row.key, b.row.key) * order;
      return cmp !== 0 ? cmp : a.idx - b.idx;
    });

    // Did anything actually move?
    let moved = false;
    for (let i = 0; i < indexed.length; i++) {
      if (indexed[i]!.idx !== i) {
        moved = true;
        break;
      }
    }
    if (!moved) {
      const noop = evolveSnapshot(snapshot, snapshot.root, {});
      return { next: noop, diff: buildDiff(snapshot.revision, noop.revision, []) };
    }

    // Rewrite the body rows in place.
    const cells = new Map(sheet.cells);
    for (let r = firstBody; r <= lastBody; r++) {
      for (let c = range.start.col; c <= range.end.col; c++) {
        cells.delete(cellKey(r, c));
      }
    }
    for (let i = 0; i < indexed.length; i++) {
      const targetRow = firstBody + i;
      const src = indexed[i]!.row.cells;
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
          field: "sort",
          summary: `${sheet.name} sorted ${payload.range} by col ${payload.sortBy.colId} ${payload.sortBy.order}`,
        },
      ]),
    };
  },
};

/**
 * Three-way comparator with Excel-flavoured semantics: blanks last,
 * numbers before strings, booleans before errors.
 */
function compareValues(a: CellValue | null, b: CellValue | null): number {
  const ar = rank(a);
  const br = rank(b);
  if (ar !== br) return ar - br;
  switch (ar) {
    case 0:
      return 0; // both blank
    case 1: {
      const an = a as number;
      const bn = b as number;
      return an === bn ? 0 : an < bn ? -1 : 1;
    }
    case 2:
      return (a as string).localeCompare(b as string, undefined, { sensitivity: "base" });
    case 3:
      return Number(a as boolean) - Number(b as boolean);
    case 4:
      return 0;
  }
  return 0;
}

function rank(v: CellValue | null): number {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return 1;
  if (typeof v === "string") return 2;
  if (typeof v === "boolean") return 3;
  return 4; // error sentinel
}
