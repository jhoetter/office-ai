import { CommandError, type CommandHandler, type DiffChange } from "@officeai/core";
import { cellKey, formatA1, formatRange, rangeArea } from "../model/refs.js";
import type { MergedCell, Sheet, XlsxSnapshot } from "../model/types.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type { MergeCellsPayload } from "./payloads.js";
import { parseRangeRef, resolveSheet } from "./validation.js";

export const mergeCellsHandler: CommandHandler<MergeCellsPayload, XlsxSnapshot> = {
  type: "xlsx:merge-cells",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    const range = parseRangeRef(payload.range);

    if (rangeArea(range) < 2) {
      throw new CommandError("invalid-range", `Range "${payload.range}" must cover ≥ 2 cells (e.g. "A1:B2")`);
    }

    const newMerge: MergedCell = {
      r1: range.start.row,
      c1: range.start.col,
      r2: range.end.row,
      c2: range.end.col,
    };

    for (const existing of sheet.merges) {
      const overlaps =
        !(existing.r2 < newMerge.r1 || existing.r1 > newMerge.r2) &&
        !(existing.c2 < newMerge.c1 || existing.c1 > newMerge.c2);
      if (!overlaps) continue;
      const sameRange =
        existing.r1 === newMerge.r1 &&
        existing.r2 === newMerge.r2 &&
        existing.c1 === newMerge.c1 &&
        existing.c2 === newMerge.c2;
      throw new CommandError(
        "overlap-with-existing-merge",
        sameRange
          ? `range "${payload.range}" is already merged`
          : `range "${payload.range}" overlaps existing merge ${formatA1({ row: existing.r1, col: existing.c1 })}:${formatA1({ row: existing.r2, col: existing.c2 })}; call xlsx:unmerge-cells first`
      );
    }

    const cells = new Map(sheet.cells);
    const changes: DiffChange[] = [];

    for (let r = newMerge.r1; r <= newMerge.r2; r++) {
      for (let c = newMerge.c1; c <= newMerge.c2; c++) {
        if (r === newMerge.r1 && c === newMerge.c1) continue;
        const key = cellKey(r, c);
        const before = cells.get(key);
        if (!before) continue;
        cells.delete(key);
        changes.push({
          kind: "node-updated",
          nodeId: sheet.id,
          path: ["sheets", sheet.index, "cells", formatA1({ row: r, col: c })],
          field: "value",
          summary: `${formatA1({ row: r, col: c })}: cleared by merge`,
        });
      }
    }

    const merges: MergedCell[] = [...sheet.merges, newMerge];
    const nextSheet: Sheet = { ...sheet, cells, merges };
    const nextWorkbook = replaceSheet(snapshot.root, nextSheet);
    const next = evolveSnapshot(snapshot, nextWorkbook, { sheets: [sheet.partPath] });

    changes.unshift({
      kind: "node-inserted",
      nodeId: sheet.id,
      path: ["sheets", sheet.index, "merges", formatRange(range)],
      summary: `merge ${formatRange(range)}`,
    });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, changes),
    };
  },
};
