import { CommandError, type CommandHandler, type DiffChange } from "@officeai/core";
import type { XlsxClipboardCell, XlsxClipboardSnapshot } from "../clipboard/snapshot.js";
import { bindEngineToWorkbook, fromEngineValue } from "../formula/workbook-host.js";
import { rewriteFormulaRefs, type AdjustFn } from "../formula/rewrite-refs.js";
import { type CellRef, type RangeRef } from "../formula/references.js";
import { cellKey, formatA1, formatRange } from "../model/refs.js";
import type { Cell, MergedCell, Sheet, XlsxSnapshot } from "../model/types.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type { PasteRangePayload } from "./payloads.js";
import { assertNotMergedNonAnchor, parseCellRef, resolveSheet } from "./validation.js";

/**
 * `xlsx:paste-range` — write a {@link XlsxClipboardSnapshot} at the
 * given top-left target. The snapshot may carry values, formulas
 * (with relative-shift), per-cell style ids, and merged-region
 * offsets relative to its own top-left.
 *
 * Spec: `spec/xlsx/agent-commands.md` §14.
 *
 * Pipeline:
 *   1. Validate target sheet + ref + bounds.
 *   2. Pick effective dimensions (transpose flips width/height).
 *   3. For every clipboard cell:
 *      - in `mode = "values"|"all"`: write `value` (or relative-shifted formula).
 *      - in `mode = "formats"|"all"`: write `styleId`.
 *   4. Recalc once at the end so dependents pick up the new formulas.
 *   5. Rewrite the merges to absolute coords + apply (skip overlapping
 *      pre-existing merges).
 *   6. Diff: one `node-updated` per touched cell + merge add.
 */
export const pasteRangeHandler: CommandHandler<PasteRangePayload, XlsxSnapshot> = {
  type: "xlsx:paste-range",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    const target = parseCellRef(payload.target);
    assertNotMergedNonAnchor(sheet, target);

    const source = payload.source;
    if (!source || source.height <= 0 || source.width <= 0) {
      throw new CommandError("invalid-clipboard", "Paste source is empty");
    }

    const transpose = !!payload.transpose;
    const height = transpose ? source.width : source.height;
    const width = transpose ? source.height : source.width;
    const mode = payload.mode ?? "all";

    if (target.row + height > 1048576 || target.col + width > 16384) {
      throw new CommandError(
        "out-of-bounds",
        `Paste of ${height}×${width} at ${payload.target} extends past the sheet edge`
      );
    }

    // Stage cell updates in a fresh map; the recalc at the end may
    // mutate the same map again with cached formula values.
    const cells = new Map(sheet.cells);
    const writeCell = mode === "all" || mode === "values" || mode === "formulas";
    const writeStyle = mode === "all" || mode === "formats";
    // In "values" mode formulas collapse to their cached value so the
    // destination ends up with literals only. "formulas" keeps the
    // formula text (relative-shifted); "all" does both.
    const keepFormula = mode === "all" || mode === "formulas";

    const cellChanges: DiffChange[] = [];
    const formulaTargets: Array<{ row: number; col: number; text: string }> = [];

    for (let dr = 0; dr < height; dr++) {
      for (let dc = 0; dc < width; dc++) {
        const sr = transpose ? dc : dr;
        const sc = transpose ? dr : dc;
        const src = source.cells[sr]?.[sc];
        const r = target.row + dr;
        const c = target.col + dc;
        const key = cellKey(r, c);
        const before = cells.get(key);

        if (!src) {
          // Empty source position. In `values` / `all` mode we still
          // overwrite the destination with empty (matches Excel's
          // behaviour: paste a blank cell, get a blank cell). Keep
          // `formats` mode non-destructive.
          if (writeCell && before) {
            cells.delete(key);
            cellChanges.push(
              cellChange(sheet, r, c, before.value ?? null, null, "value", before.formula?.text ?? null, null)
            );
          }
          continue;
        }

        // Re-shift formula refs against the destination anchor.
        let formulaText: string | null = null;
        if (writeCell && keepFormula && src.formula) {
          formulaText = shiftFormula(
            src.formula,
            source.origin.sheet || sheet.name,
            sheet.name,
            sr,
            sc,
            r,
            c
          );
          formulaTargets.push({ row: r, col: c, text: formulaText });
        }

        const nextStyleId = writeStyle ? src.styleId : before?.styleId;
        // The cached value gets overwritten by the recalc later when a
        // formula is present; for value-only sources we trust `src.value`.
        const nextValue = writeCell ? src.value : (before?.value ?? null);

        const nextCell: Cell = {
          row: r,
          col: c,
          value: nextValue,
          ...(formulaText !== null ? { formula: { text: formulaText } } : {}),
          ...(nextStyleId !== undefined ? { styleId: nextStyleId } : {}),
        };

        if (cellsEqual(before, nextCell)) continue;
        cells.set(key, nextCell);
        cellChanges.push(
          cellChange(
            sheet,
            r,
            c,
            before?.value ?? null,
            nextCell.value,
            formulaText !== null ? "formula" : "value",
            before?.formula?.text ?? null,
            formulaText
          )
        );
      }
    }

    // Apply merges from the snapshot, translated to absolute coords.
    // We add new merges; we DO NOT touch pre-existing merges that
    // happen to live in the destination — that would surprise the user.
    // Instead, partial overlap aborts the paste with `merge-overlap`.
    let nextMerges = sheet.merges;
    if (source.merges.length > 0 && mode === "all") {
      const incoming: MergedCell[] = source.merges.map((m) => {
        const dr0 = transpose ? m.c0 : m.r0;
        const dc0 = transpose ? m.r0 : m.c0;
        const dr1 = transpose ? m.c1 : m.r1;
        const dc1 = transpose ? m.r1 : m.c1;
        return {
          r1: target.row + Math.min(dr0, dr1),
          c1: target.col + Math.min(dc0, dc1),
          r2: target.row + Math.max(dr0, dr1),
          c2: target.col + Math.max(dc0, dc1),
        };
      });
      assertNoMergeOverlap(sheet.merges, incoming);
      nextMerges = sheet.merges.concat(incoming);
    }

    let nextSheet: Sheet = { ...sheet, cells, merges: nextMerges };
    let nextWorkbook = replaceSheet(snapshot.root, nextSheet);

    // Recalc. We seed every formula in the workbook (including the
    // freshly pasted ones) and let the engine produce cached values.
    if (formulaTargets.length > 0) {
      const { engine, host } = bindEngineToWorkbook(nextWorkbook);
      void host.seedFormulas(engine);
      const result = engine.recalc();

      // Walk every formula cell whose cached value differs and
      // update the typed model accordingly.
      const sheetUpdates = new Map<string, Map<string, Cell>>();
      sheetUpdates.set(sheet.name, new Map(cells));

      for (const [key, value] of result.values) {
        const sep = key.indexOf("!");
        if (sep === -1) continue;
        const sName = key.slice(0, sep);
        const coord = key.slice(sep + 1);
        const colon = coord.indexOf(":");
        if (colon === -1) continue;
        const row = Number.parseInt(coord.slice(0, colon), 10);
        const col = Number.parseInt(coord.slice(colon + 1), 10);
        if (!Number.isFinite(row) || !Number.isFinite(col)) continue;

        const ws = nextWorkbook.sheets.find((s) => s.name === sName);
        if (!ws) continue;
        let map = sheetUpdates.get(sName);
        if (!map) {
          map = new Map(ws.cells);
          sheetUpdates.set(sName, map);
        }
        const existing = map.get(cellKey(row, col));
        if (!existing || !existing.formula) continue;
        const nextValue = fromEngineValue(value);
        if (cellValuesEqual(existing.value, nextValue)) continue;
        map.set(cellKey(row, col), { ...existing, value: nextValue });
      }

      const dirtySheetPaths = new Set<string>([sheet.partPath]);
      for (const [sName, m] of sheetUpdates) {
        const ws = nextWorkbook.sheets.find((s) => s.name === sName);
        if (!ws) continue;
        const updated: Sheet = { ...ws, cells: m };
        if (sName === sheet.name) nextSheet = { ...updated, merges: nextMerges };
        nextWorkbook = replaceSheet(nextWorkbook, sName === sheet.name ? nextSheet : updated);
        dirtySheetPaths.add(ws.partPath);
      }

      const next = evolveSnapshot(snapshot, nextWorkbook, { sheets: [...dirtySheetPaths] });
      return finish(snapshot, next, sheet, target, source, cellChanges, mode);
    }

    const next = evolveSnapshot(snapshot, nextWorkbook, { sheets: [sheet.partPath] });
    return finish(snapshot, next, sheet, target, source, cellChanges, mode);
  },
};

function finish(
  prev: XlsxSnapshot,
  next: XlsxSnapshot,
  sheet: Sheet,
  target: { row: number; col: number },
  source: XlsxClipboardSnapshot,
  changes: ReadonlyArray<DiffChange>,
  mode: PasteRangePayload["mode"]
): { next: XlsxSnapshot; diff: ReturnType<typeof buildDiff> } {
  const summary: DiffChange[] = [...changes];
  if (changes.length === 0) {
    summary.push({
      kind: "node-updated",
      nodeId: sheet.id,
      path: ["sheets", sheet.index],
      field: "noop",
      summary: `paste at ${formatA1(target)} (${source.height}×${source.width}, mode=${mode}): no changes`,
    });
  }
  return { next, diff: buildDiff(prev.revision, next.revision, summary) };
}

function shiftFormula(
  text: string,
  sourceSheet: string,
  destSheet: string,
  sr: number,
  sc: number,
  dr: number,
  dc: number
): string {
  const dRow = dr - sr;
  const dCol = dc - sc;
  const adjust: AdjustFn = (ref) => {
    if ("row" in ref) {
      const next: CellRef = {
        sheet: ref.sheet,
        row: ref.abs & 1 ? ref.row : ref.row + dRow,
        col: ref.abs & 2 ? ref.col : ref.col + dCol,
        abs: ref.abs,
      };
      return next;
    }
    const r0Abs = ref.abs0 & 1;
    const c0Abs = ref.abs0 & 2;
    const r1Abs = ref.abs1 & 1;
    const c1Abs = ref.abs1 & 2;
    const next: RangeRef = {
      sheet: ref.sheet,
      r0: r0Abs ? ref.r0 : ref.r0 + dRow,
      c0: c0Abs ? ref.c0 : ref.c0 + dCol,
      r1: r1Abs ? ref.r1 : ref.r1 + dRow,
      c1: c1Abs ? ref.c1 : ref.c1 + dCol,
      abs0: ref.abs0,
      abs1: ref.abs1,
    };
    return next;
  };
  // The anchor we pass governs sheet-prefix elision and the parser's
  // default-sheet resolution. Use the destination cell so refs that
  // were sheet-implicit at the source stay sheet-implicit at the
  // destination, but cross-sheet refs keep their explicit prefix.
  const anchor: CellRef = { sheet: destSheet, row: dr, col: dc, abs: 0 };
  void sourceSheet;
  return rewriteFormulaRefs(text, anchor, adjust).text;
}

function assertNoMergeOverlap(
  existing: ReadonlyArray<MergedCell>,
  incoming: ReadonlyArray<MergedCell>
): void {
  for (const i of incoming) {
    for (const e of existing) {
      const overlap = !(e.r2 < i.r1 || e.r1 > i.r2 || e.c2 < i.c1 || e.c1 > i.c2);
      if (!overlap) continue;
      const exact = e.r1 === i.r1 && e.r2 === i.r2 && e.c1 === i.c1 && e.c2 === i.c2;
      if (exact) continue;
      throw new CommandError(
        "merge-overlap",
        `Pasted merge ${formatA1({ row: i.r1, col: i.c1 })}:${formatA1({ row: i.r2, col: i.c2 })} overlaps existing merge ${formatA1({ row: e.r1, col: e.c1 })}:${formatA1({ row: e.r2, col: e.c2 })}`
      );
    }
  }
}

function cellsEqual(a: Cell | undefined, b: Cell): boolean {
  if (!a) return false;
  if ((a.formula?.text ?? null) !== (b.formula?.text ?? null)) return false;
  if ((a.styleId ?? -1) !== (b.styleId ?? -1)) return false;
  return cellValuesEqual(a.value, b.value);
}

function cellValuesEqual(a: Cell["value"], b: Cell["value"]): boolean {
  if (a === b) return true;
  if (a === null || b === null) return false;
  if (typeof a === "object" && typeof b === "object" && "kind" in a && "kind" in b) {
    return a.kind === b.kind && a.code === b.code;
  }
  return false;
}

function cellChange(
  sheet: Sheet,
  row: number,
  col: number,
  beforeVal: Cell["value"],
  afterVal: Cell["value"],
  field: "value" | "formula",
  beforeFormula: string | null,
  afterFormula: string | null
): DiffChange {
  return {
    kind: "node-updated",
    nodeId: sheet.id,
    path: ["sheets", sheet.index, "cells", `${sheet.name}!${formatA1({ row, col })}`],
    field,
    summary: `${formatA1({ row, col })}: ${formatVal(beforeVal)} → ${formatVal(afterVal)}${
      afterFormula !== null ? ` (=${afterFormula})` : ""
    }`,
    meta: {
      before: { value: beforeVal, formula: beforeFormula },
      after: { value: afterVal, formula: afterFormula },
    },
  };
}

function formatVal(v: Cell["value"]): string {
  if (v === null || v === undefined) return "∅";
  if (typeof v === "string") return JSON.stringify(v);
  if (typeof v === "object" && v && "kind" in v) return v.code;
  return String(v);
}

// Re-export for the agent surface.
export type { XlsxClipboardCell, XlsxClipboardSnapshot };
// Avoid an unused-import warning on the helper formatRange.
void formatRange;
