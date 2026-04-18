import { CommandError, type CommandHandler, type DiffChange } from "@officeai/core";
import { cellKey, formatA1, formatRange, rangeArea } from "../model/refs.js";
import type { Cell, Sheet, XlsxSnapshot } from "../model/types.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type { SetRangeValuesPayload } from "./payloads.js";
import { assertNotFormulaString, findContainingMerge, parseRangeRef, resolveSheet } from "./validation.js";

const CELL_CAP = 100_000;

export const setRangeValuesHandler: CommandHandler<SetRangeValuesPayload, XlsxSnapshot> = {
  type: "xlsx:set-range-values",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    const range = parseRangeRef(payload.range);

    const rows = range.end.row - range.start.row + 1;
    const cols = range.end.col - range.start.col + 1;

    if (payload.values.length !== rows) {
      throw new CommandError(
        "dimension-mismatch",
        `range "${payload.range}" expects ${rows} rows; values has ${payload.values.length}`
      );
    }
    for (let i = 0; i < rows; i++) {
      const row = payload.values[i];
      if (!Array.isArray(row) || row.length !== cols) {
        throw new CommandError(
          "dimension-mismatch",
          `range "${payload.range}" row ${i} expects ${cols} columns; got ${row?.length ?? 0}`
        );
      }
      for (const v of row) assertNotFormulaString(v);
    }

    const area = rangeArea(range);
    if (area > CELL_CAP) {
      throw new CommandError(
        "cell-cap-exceeded",
        `range covers ${area} cells which exceeds the ${CELL_CAP} cell cap; split into smaller batches`
      );
    }

    assertNoPartialMergeOverlap(sheet, range);

    const cells = new Map(sheet.cells);
    const changes: DiffChange[] = [];

    for (let r = 0; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        const row = range.start.row + r;
        const col = range.start.col + c;
        const value = payload.values[r][c];
        const key = cellKey(row, col);
        const before = cells.get(key);

        if (value === null) {
          if (!before) continue;
          cells.delete(key);
        } else {
          const next: Cell = { row, col, value };
          if (cellsValueEqual(before, next)) continue;
          cells.set(key, next);
        }

        changes.push({
          kind: "node-updated",
          nodeId: sheet.id,
          path: ["sheets", sheet.index, "cells", formatA1({ row, col })],
          field: "value",
          summary: `${formatA1({ row, col })}: ${formatValue(before?.value ?? null)} → ${formatValue(value)}`,
        });
      }
    }

    const nextSheet: Sheet = { ...sheet, cells };
    const nextWorkbook = replaceSheet(snapshot.root, nextSheet);
    const next = evolveSnapshot(snapshot, nextWorkbook, { sheets: [sheet.partPath] });

    if (changes.length === 0) {
      changes.push({
        kind: "node-updated",
        nodeId: sheet.id,
        path: ["sheets", sheet.index],
        field: "noop",
        summary: `${formatRange(range)}: no value changes`,
      });
    }

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, changes),
    };
  },
};

function assertNoPartialMergeOverlap(sheet: Sheet, range: ReturnType<typeof parseRangeRef>): void {
  for (const merge of sheet.merges) {
    const overlaps =
      !(merge.r2 < range.start.row || merge.r1 > range.end.row) &&
      !(merge.c2 < range.start.col || merge.c1 > range.end.col);
    if (!overlaps) continue;
    const fullyContained =
      merge.r1 >= range.start.row &&
      merge.r2 <= range.end.row &&
      merge.c1 >= range.start.col &&
      merge.c2 <= range.end.col;
    if (fullyContained) continue;
    throw new CommandError(
      "merge-overlap",
      `range overlaps merge ${formatA1({ row: merge.r1, col: merge.c1 })}:${formatA1({ row: merge.r2, col: merge.c2 })} partially; call xlsx:unmerge-cells first or shrink the range`
    );
  }
  void findContainingMerge;
}

function cellsValueEqual(a: Cell | undefined, b: Cell): boolean {
  if (!a) return false;
  if (a.formula || b.formula) return false;
  const av = a.value;
  const bv = b.value;
  if (av === bv) return true;
  if (typeof av === "object" && av && typeof bv === "object" && bv) {
    return (av as { code?: string }).code === (bv as { code?: string }).code;
  }
  return false;
}

function formatValue(v: unknown): string {
  if (v === null || v === undefined) return "∅";
  if (typeof v === "string") return JSON.stringify(v);
  if (typeof v === "object" && v && "kind" in v) return String((v as unknown as { code: string }).code);
  return String(v);
}
