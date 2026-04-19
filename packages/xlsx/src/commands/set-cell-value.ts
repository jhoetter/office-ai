import type { CommandHandler, DiffChange } from "@officeai/core";
import { bindEngineToWorkbook, fromEngineValue } from "../formula/workbook-host.js";
import { parseCellKey as parseEngineCellKey } from "../formula/references.js";
import { cellKey, formatA1 } from "../model/refs.js";
import type { Cell, CellValue, Sheet, XlsxSnapshot } from "../model/types.js";
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

    let nextSheet: Sheet = { ...sheet, cells };
    let nextWorkbook = replaceSheet(snapshot.root, nextSheet);

    // Recalc just the formulas that transitively depend on the cell we
    // touched. Without this, any formula referencing the cell would
    // keep its stale cached value (e.g. B2 := 44, but E8 := =B2 still
    // displays 42). Critically, we *only* refresh cells in the
    // dependency closure of the changed ref: if the workbook contains
    // unrelated formulas with stored cached values that happen to
    // disagree with what the engine would compute, we leave them
    // alone — touching them would break OOXML round-trip identity for
    // parts the user didn't actually edit.
    const { engine, host } = bindEngineToWorkbook(nextWorkbook);
    void host.seedFormulas(engine);
    const changedRef = { sheet: sheet.name, row: addr.row, col: addr.col, abs: 0 as const };
    const dependentKeys = engine.graph.collectDependents(changedRef);

    const dependentChanges: DiffChange[] = [];
    const dirtySheetPaths = new Set<string>([sheet.partPath]);
    if (dependentKeys.size > 0) {
      const result = engine.recalc();
      const sheetUpdates = new Map<string, Map<string, Cell>>();
      for (const [k, value] of result.values) {
        if (!dependentKeys.has(k)) continue;
        const parts = parseEngineCellKey(k);
        const ws = nextWorkbook.sheets.find((s) => s.name === parts.sheet);
        if (!ws) continue;
        let map = sheetUpdates.get(parts.sheet);
        if (!map) {
          map = new Map(ws.cells);
          sheetUpdates.set(parts.sheet, map);
        }
        const existing = map.get(cellKey(parts.row, parts.col));
        // Only refresh cached values for actual formula cells. A pure
        // value cell that ended up in `result.values` would otherwise
        // have its typed value silently rewritten.
        if (!existing || !existing.formula) continue;
        const nextValue = fromEngineValue(value);
        if (cellValuesEqual(existing.value, nextValue)) continue;
        const updated: Cell = { ...existing, value: nextValue };
        map.set(cellKey(parts.row, parts.col), updated);
        dirtySheetPaths.add(ws.partPath);
        dependentChanges.push({
          kind: "node-updated",
          nodeId: ws.id,
          path: [
            "sheets",
            ws.index,
            "cells",
            `${parts.sheet}!${formatA1({ row: parts.row, col: parts.col })}`,
          ],
          field: "cachedValue",
          summary: `${parts.sheet}!${formatA1({ row: parts.row, col: parts.col })}: ${formatValue(existing.value)} → ${formatValue(nextValue)}`,
        });
      }

      if (sheetUpdates.size > 0) {
        for (const [sName, map] of sheetUpdates) {
          const ws = nextWorkbook.sheets.find((s) => s.name === sName);
          if (!ws) continue;
          const updated: Sheet = sName === sheet.name ? { ...nextSheet, cells: map } : { ...ws, cells: map };
          nextWorkbook = replaceSheet(nextWorkbook, updated);
          if (sName === sheet.name) nextSheet = updated;
        }
      }
    }

    const next = evolveSnapshot(snapshot, nextWorkbook, { sheets: [...dirtySheetPaths] });
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
        ...dependentChanges,
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

function cellValuesEqual(a: CellValue, b: CellValue): boolean {
  if (a === b) return true;
  if (a === null || b === null) return false;
  if (typeof a === "object" && typeof b === "object" && "kind" in a && "kind" in b) {
    return a.kind === b.kind && a.code === b.code;
  }
  return false;
}
