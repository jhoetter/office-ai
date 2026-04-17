import type { CommandHandler } from "@officeai/core";
import { cellKey, formatA1 } from "../model/refs.js";
import type { Cell, Sheet, XlsxSnapshot } from "../model/types.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type { SetCellValuePayload } from "./payloads.js";
import {
  assertNotFormulaString,
  assertNotMergedNonAnchor,
  parseCellRef,
  resolveSheet,
} from "./validation.js";

export const setCellValueHandler: CommandHandler<SetCellValuePayload, XlsxSnapshot> = {
  type: "xlsx:set-cell-value",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    const addr = parseCellRef(payload.ref);
    assertNotFormulaString(payload.value);
    assertNotMergedNonAnchor(sheet, addr);

    const key = cellKey(addr.row, addr.col);
    const cells = new Map(sheet.cells);
    const before = sheet.cells.get(key);

    if (payload.value === null) {
      cells.delete(key);
    } else {
      const nextCell: Cell = { row: addr.row, col: addr.col, value: payload.value };
      cells.set(key, nextCell);
    }

    const nextSheet: Sheet = { ...sheet, cells };
    const nextWorkbook = replaceSheet(snapshot.root, nextSheet);
    const next = evolveSnapshot(snapshot, nextWorkbook, { sheets: [sheet.partPath] });
    const summary = describeChange(payload.ref, before?.value ?? null, payload.value);

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, [
        {
          kind: "node-updated",
          nodeId: sheet.id,
          path: ["sheets", sheet.index, "cells", formatA1(addr)],
          field: "value",
          summary,
        },
      ]),
    };
  },
};

function describeChange(ref: string, before: unknown, after: unknown): string {
  return `${ref}: ${formatValue(before)} → ${formatValue(after)}`;
}

function formatValue(v: unknown): string {
  if (v === null || v === undefined) return "∅";
  if (typeof v === "string") return JSON.stringify(v);
  if (typeof v === "object" && v && "kind" in v) return String((v as unknown as { code: string }).code);
  return String(v);
}
