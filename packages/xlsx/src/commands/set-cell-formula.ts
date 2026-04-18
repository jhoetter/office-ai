import { CommandError, type CommandHandler, type DiffChange } from "@officeai/core";
import { bindEngineToWorkbook, fromEngineValue } from "../formula/workbook-host.js";
import { parseCellKey as parseEngineCellKey, type CellRef } from "../formula/references.js";
import { cellKey, formatA1 } from "../model/refs.js";
import type { Cell, CellValue, Sheet, XlsxSnapshot } from "../model/types.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type { SetCellFormulaPayload } from "./payloads.js";
import { assertNotMergedNonAnchor, parseCellRef, resolveSheet } from "./validation.js";

/**
 * `xlsx:set-cell-formula` — author or replace a formula on a single
 * cell, then run a workbook-wide recalc and write the cached values
 * for every dependent that changed.
 *
 * Spec: `spec/xlsx/agent-commands.md` §2 + `formula-engine.md` §15.
 *
 * Pipeline:
 *   1. Validate sheet / ref / merge anchor.
 *   2. Strip a leading `=`. Empty body becomes a `set-cell-value(null)`.
 *   3. Parse via the engine. Parse errors reject loudly with
 *      `formula-parse-error` (per `EC-F3`); cells with bad formulas
 *      authored by an LLM are rejected so the LLM gets feedback,
 *      whereas import-time parse errors *are* preserved as `#NAME?`.
 *   4. Build a fresh engine bound to the workbook, seed it with every
 *      existing formula, add the new formula, and recalc.
 *   5. Apply the new formula + cached value to the target cell.
 *      For every other formula cell whose cached value changed, write
 *      the new cached value back into the typed model.
 *   6. Diff: one `formula-updated` for the source cell + one
 *      `cell-updated` per dependent that changed. If the recalc
 *      surfaced cycles, attach a `circular` change with `meta.cycle`.
 */
export const setCellFormulaHandler: CommandHandler<SetCellFormulaPayload, XlsxSnapshot> = {
  type: "xlsx:set-cell-formula",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    const addr = parseCellRef(payload.ref);
    assertNotMergedNonAnchor(sheet, addr);

    const text = stripLeadingEquals(payload.formula);
    if (text.length === 0) {
      // Empty formula collapses to clear-the-cell semantics, mirroring
      // §2.2's "empty-formula" branch.
      return clearCell(snapshot, sheet, addr, payload);
    }

    const { engine, host } = bindEngineToWorkbook(snapshot.root);
    const seedFailures = host.seedFormulas(engine);
    // Seed-time parse failures are non-fatal — they reflect formulas
    // that were already in the workbook (per `EC-F3`); we keep them
    // literal-only and proceed with the requested edit.
    void seedFailures;

    const targetRef: CellRef = { sheet: sheet.name, row: addr.row, col: addr.col, abs: 0 };
    let parsed;
    try {
      parsed = engine.parse(text, targetRef);
    } catch (e) {
      const err = e as Error & { kind?: string };
      const kindHint = err.kind ? ` [${err.kind}]` : "";
      throw new CommandError(
        "formula-parse-error",
        `Failed to parse formula ${JSON.stringify(payload.formula)} for ${payload.ref}${kindHint}: ${err.message}`,
        { cause: e }
      );
    }

    engine.addCell(targetRef, parsed, null);
    const result = engine.recalc();

    // 5a. Capture before-state for the target cell + the cells whose
    // cached value changed, so the diff carries before→after deltas.
    const before = sheet.cells.get(cellKey(addr.row, addr.col));
    const sheetUpdates = new Map<string, ReadonlyMap<string, Cell>>();
    sheetUpdates.set(sheet.name, new Map(sheet.cells));

    // 5b. Write target cell.
    const targetCachedValue = result.values.get(`${sheet.name}!${addr.row}:${addr.col}`);
    const newCellValue: CellValue =
      targetCachedValue !== undefined ? fromEngineValue(targetCachedValue) : null;
    const newCell: Cell = {
      row: addr.row,
      col: addr.col,
      value: newCellValue,
      formula: { text },
    };
    const targetCells = sheetUpdates.get(sheet.name)!;
    (targetCells as Map<string, Cell>).set(cellKey(addr.row, addr.col), newCell);

    // 5c. Apply downstream cached-value updates. The `result.values`
    // map carries every cell touched by the recalc (excluding the
    // target). Iterate, look up the existing typed cell, and rewrite
    // its `value` while preserving its `formula`.
    const dependentChanges: DiffChange[] = [];
    for (const [key, value] of result.values) {
      if (key === `${sheet.name}!${addr.row}:${addr.col}`) continue;
      const parts = parseEngineCellKey(key);
      const sName = parts.sheet;
      const ref: { row: number; col: number } = { row: parts.row, col: parts.col };
      const existingSheet = snapshot.root.sheets.find((s) => s.name === sName);
      if (!existingSheet) continue;
      let cellsForSheet = sheetUpdates.get(sName);
      if (!cellsForSheet) {
        cellsForSheet = new Map(existingSheet.cells);
        sheetUpdates.set(sName, cellsForSheet);
      }
      const existing = cellsForSheet.get(cellKey(ref.row, ref.col));
      if (!existing || !existing.formula) continue;
      const nextValue = fromEngineValue(value);
      if (cellValuesEqual(existing.value, nextValue)) continue;
      const updated: Cell = { ...existing, value: nextValue };
      (cellsForSheet as Map<string, Cell>).set(cellKey(ref.row, ref.col), updated);
      dependentChanges.push({
        kind: "node-updated",
        nodeId: existingSheet.id,
        path: ["sheets", existingSheet.index, "cells", `${sName}!${formatA1(ref)}`],
        field: "cachedValue",
        summary: `${sName}!${formatA1(ref)}: ${formatVal(existing.value)} → ${formatVal(nextValue)}`,
      });
    }

    // 6a. Stitch the workbook back together with all updated sheets.
    let nextWorkbook = snapshot.root;
    const dirtySheetPaths: string[] = [];
    for (const [sName, cells] of sheetUpdates) {
      const s = nextWorkbook.sheets.find((x) => x.name === sName);
      if (!s) continue;
      const nextSheet: Sheet = { ...s, cells };
      nextWorkbook = replaceSheet(nextWorkbook, nextSheet);
      dirtySheetPaths.push(s.partPath);
    }
    const next = evolveSnapshot(snapshot, nextWorkbook, { sheets: dirtySheetPaths });

    // 6b. Build the diff.
    const changes: DiffChange[] = [];
    changes.push({
      kind: "node-updated",
      nodeId: sheet.id,
      path: ["sheets", sheet.index, "cells", `${sheet.name}!${formatA1(addr)}`],
      field: "formula",
      summary: `${sheet.name}!${formatA1(addr)}: =${text} (${formatVal(newCellValue)})`,
      meta: {
        before: before ? { value: before.value, formula: before.formula?.text ?? null } : null,
        after: { value: newCellValue, formula: text },
      },
    });
    changes.push(...dependentChanges);
    if (result.cycles.length > 0) {
      changes.push({
        kind: "node-updated",
        nodeId: sheet.id,
        path: ["sheets", sheet.index, "cells", `${sheet.name}!${formatA1(addr)}`],
        field: "circular",
        summary: `Circular reference involving ${result.cycles[0].length} cells`,
        meta: { cycle: result.cycles.flat() },
      });
    }

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, changes),
    };
  },
};

function stripLeadingEquals(text: string): string {
  return text.startsWith("=") ? text.slice(1) : text;
}

function clearCell(
  snapshot: XlsxSnapshot,
  sheet: Sheet,
  addr: { row: number; col: number },
  payload: SetCellFormulaPayload
): { next: XlsxSnapshot; diff: ReturnType<typeof buildDiff> } {
  const cells = new Map(sheet.cells);
  const before = sheet.cells.get(cellKey(addr.row, addr.col));
  cells.delete(cellKey(addr.row, addr.col));
  const nextSheet: Sheet = { ...sheet, cells };
  const nextWorkbook = replaceSheet(snapshot.root, nextSheet);
  const next = evolveSnapshot(snapshot, nextWorkbook, { sheets: [sheet.partPath] });
  return {
    next,
    diff: buildDiff(snapshot.revision, next.revision, [
      {
        kind: "node-updated",
        nodeId: sheet.id,
        path: ["sheets", sheet.index, "cells", `${sheet.name}!${formatA1(addr)}`],
        field: "value",
        summary: `${payload.ref}: ${formatVal(before?.value ?? null)} → ∅ (empty formula)`,
      },
    ]),
  };
}

function cellValuesEqual(a: CellValue, b: CellValue): boolean {
  if (a === b) return true;
  if (a === null || b === null) return false;
  if (typeof a === "object" && typeof b === "object" && "kind" in a && "kind" in b) {
    return a.kind === b.kind && a.code === b.code;
  }
  return false;
}

function formatVal(v: CellValue): string {
  if (v === null) return "∅";
  if (typeof v === "string") return JSON.stringify(v);
  if (typeof v === "object" && v && "kind" in v) return v.code;
  return String(v);
}
