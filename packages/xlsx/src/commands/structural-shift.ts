import { CommandError, type DiffChange } from "@officeai/core";
import {
  adjustForDeleteColumn,
  adjustForDeleteRow,
  adjustForInsertColumn,
  adjustForInsertRow,
  type CellRef,
  type RangeRef,
} from "../formula/references.js";
import { rewriteFormulaRefs } from "../formula/rewrite-refs.js";
import { bindEngineToWorkbook, fromEngineValue } from "../formula/workbook-host.js";
import type { CellError } from "../formula/errors.js";
import { cellKey, formatA1 } from "../model/refs.js";
import type { Cell, CellValue, MergedCell, Sheet, XlsxSnapshot, XlsxWorkbook } from "../model/types.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import { resolveSheet } from "./validation.js";

/**
 * Backbone for the four structural reshape commands
 * (`xlsx:insert-row`, `xlsx:insert-column`, `xlsx:delete-row`,
 * `xlsx:delete-column`).
 *
 * Spec: `spec/xlsx/agent-commands.md` §§5–8.
 *
 * Pipeline (per spec):
 *   1. Validate sheet + `at`/`count` against Excel limits.
 *   2. Walk the target sheet's merges; reject `merge-boundary-crossed`.
 *   3. Build a new cell map: shift every surviving cell's row/col,
 *      drop cells that fall inside a deletion band.
 *   4. Build a new merges array: shift, expand, or drop per the
 *      operation.
 *   5. Across the whole workbook, walk every formula's text through
 *      `rewriteFormulaRefs` with the matching `adjustForXxx`. Track
 *      casualties (cells whose formula now contains `#REF!`).
 *   6. Build a fresh engine over the post-mutation workbook and
 *      recalc so cached values reflect the new dependency layout.
 *   7. Stitch updated cells back in, emit a single summary diff plus
 *      per-cell formula/cached-value changes.
 *
 * Defined names + comments + hyperlinks are not in the typed model
 * yet (Phase 7+ deferral); their adjustment lands with those phases.
 */

export type Axis = "row" | "column";
export type Operation = "insert" | "delete";

interface StructuralPayload {
  readonly sheet: string;
  readonly at: number;
  readonly count: number;
}

const MAX_ROW = 1_048_576;
const MAX_COL = 16_384;

export function applyStructuralShift(
  snapshot: XlsxSnapshot,
  payload: StructuralPayload,
  axis: Axis,
  op: Operation
): { next: XlsxSnapshot; diff: ReturnType<typeof buildDiff> } {
  const sheet = resolveSheet(snapshot.root, payload.sheet);
  validatePayload(payload, axis);
  const at0 = payload.at - 1;
  const count = payload.count;

  assertNoMergeBoundaryCross(sheet, axis, op, at0, count);

  const shiftedSheet = applySheetShift(sheet, axis, op, at0, count);

  const formulaRewrites = rewriteWorkbookFormulas(snapshot.root, shiftedSheet, axis, op, at0, count);
  const workbookAfterFormulas = applyFormulaRewrites(snapshot.root, shiftedSheet, formulaRewrites);

  const recalc = recalcWorkbook(workbookAfterFormulas);
  const finalWorkbook = applyCachedValues(workbookAfterFormulas, recalc.values);

  const dirtySheets = new Set<string>([sheet.partPath]);
  for (const [name, _] of formulaRewrites.byCell) {
    void _;
    const s = finalWorkbook.sheets.find((x) => x.name === splitCellKey(name).sheet);
    if (s) dirtySheets.add(s.partPath);
  }
  for (const key of recalc.values.keys()) {
    const parts = splitCellKey(key);
    const s = finalWorkbook.sheets.find((x) => x.name === parts.sheet);
    if (s) dirtySheets.add(s.partPath);
  }

  const next = evolveSnapshot(snapshot, finalWorkbook, { sheets: [...dirtySheets] });

  const changes = buildChanges(
    snapshot,
    finalWorkbook,
    sheet,
    axis,
    op,
    payload,
    shiftedSheet,
    formulaRewrites,
    recalc.values
  );

  return { next, diff: buildDiff(snapshot.revision, next.revision, changes) };
}

// ── Validation ────────────────────────────────────────────────────────────

function validatePayload(payload: StructuralPayload, axis: Axis): void {
  if (!Number.isInteger(payload.at) || payload.at < 1) {
    throw new CommandError("invalid-position", `\`at\` must be a 1-based integer ≥ 1; got ${payload.at}`);
  }
  if (!Number.isInteger(payload.count) || payload.count < 1) {
    throw new CommandError("invalid-count", `\`count\` must be an integer ≥ 1; got ${payload.count}`);
  }
  const max = axis === "row" ? MAX_ROW : MAX_COL;
  const label = axis === "row" ? "row" : "column";
  if (payload.at > max) {
    throw new CommandError("invalid-position", `\`at\` (${payload.at}) exceeds Excel max ${label} ${max}`);
  }
  if (payload.at + payload.count - 1 > max) {
    throw new CommandError(
      "invalid-count",
      `\`at + count - 1\` (${payload.at + payload.count - 1}) exceeds Excel max ${label} ${max}`
    );
  }
}

function assertNoMergeBoundaryCross(
  sheet: Sheet,
  axis: Axis,
  op: Operation,
  at0: number,
  count: number
): void {
  if (op === "insert") {
    for (const m of sheet.merges) {
      const lo = axis === "row" ? m.r1 : m.c1;
      const hi = axis === "row" ? m.r2 : m.c2;
      if (lo < at0 && at0 < hi) {
        throw new CommandError(
          "merge-boundary-crossed",
          `Insertion at ${axis} ${at0 + 1} would split the merge ${formatMerge(m)}; insert above/below the merge or call xlsx:unmerge-cells first`
        );
      }
    }
    return;
  }
  const end = at0 + count - 1;
  for (const m of sheet.merges) {
    const lo = axis === "row" ? m.r1 : m.c1;
    const hi = axis === "row" ? m.r2 : m.c2;
    const overlaps = lo <= end && hi >= at0;
    if (!overlaps) continue;
    const fullyInside = lo >= at0 && hi <= end;
    if (fullyInside) continue;
    throw new CommandError(
      "merge-boundary-crossed",
      `Deletion of ${axis}s ${at0 + 1}..${end + 1} would split the merge ${formatMerge(m)}; delete the entire merge band or call xlsx:unmerge-cells first`
    );
  }
}

function formatMerge(m: MergedCell): string {
  return `${formatA1({ row: m.r1, col: m.c1 })}:${formatA1({ row: m.r2, col: m.c2 })}`;
}

// ── Sheet-local cell + merge shift ────────────────────────────────────────

interface ShiftedSheet {
  readonly nextSheet: Sheet;
  readonly shiftedCells: number;
  readonly droppedCells: ReadonlyArray<{ row: number; col: number; before: Cell }>;
}

function applySheetShift(sheet: Sheet, axis: Axis, op: Operation, at0: number, count: number): ShiftedSheet {
  const newCells = new Map<string, Cell>();
  const dropped: Array<{ row: number; col: number; before: Cell }> = [];
  let shifted = 0;
  const end = at0 + count - 1;

  for (const cell of sheet.cells.values()) {
    const coord = axis === "row" ? cell.row : cell.col;
    let nextCoord = coord;

    if (op === "insert") {
      if (coord >= at0) nextCoord = coord + count;
    } else {
      if (coord >= at0 && coord <= end) {
        dropped.push({ row: cell.row, col: cell.col, before: cell });
        continue;
      }
      if (coord > end) nextCoord = coord - count;
    }

    if (nextCoord !== coord) shifted++;
    const next: Cell =
      nextCoord === coord ? cell : axis === "row" ? { ...cell, row: nextCoord } : { ...cell, col: nextCoord };
    newCells.set(cellKey(next.row, next.col), next);
  }

  const newMerges = shiftMerges(sheet.merges, axis, op, at0, count);
  const nextSheet: Sheet = { ...sheet, cells: newCells, merges: newMerges };
  return { nextSheet, shiftedCells: shifted, droppedCells: dropped };
}

function shiftMerges(
  merges: ReadonlyArray<MergedCell>,
  axis: Axis,
  op: Operation,
  at0: number,
  count: number
): MergedCell[] {
  const out: MergedCell[] = [];
  const end = at0 + count - 1;
  for (const m of merges) {
    const lo = axis === "row" ? m.r1 : m.c1;
    const hi = axis === "row" ? m.r2 : m.c2;

    if (op === "insert") {
      // Three cases per applied precheck (mid-merge already rejected):
      //   - merge fully below `at0`            → shift both endpoints
      //   - merge straddles `at0` from above   → expand `hi += count`
      //   - merge fully above (`hi < at0`)     → unchanged
      let nextLo = lo;
      let nextHi = hi;
      if (lo >= at0) {
        nextLo = lo + count;
        nextHi = hi + count;
      } else if (hi >= at0) {
        nextHi = hi + count;
      }
      out.push(rebuildMerge(m, axis, nextLo, nextHi));
      continue;
    }

    // delete: precheck guaranteed no partial overlaps; only fully-inside,
    // entirely-above, or entirely-below survive.
    const fullyInside = lo >= at0 && hi <= end;
    if (fullyInside) continue;
    if (hi < at0) {
      out.push(m);
      continue;
    }
    out.push(rebuildMerge(m, axis, lo - count, hi - count));
  }
  return out;
}

function rebuildMerge(m: MergedCell, axis: Axis, lo: number, hi: number): MergedCell {
  if (axis === "row") return { ...m, r1: lo, r2: hi };
  return { ...m, c1: lo, c2: hi };
}

// ── Workbook formula rewrite ──────────────────────────────────────────────

interface FormulaRewrite {
  readonly sheetName: string;
  readonly cell: Cell;
  readonly oldText: string;
  readonly newText: string;
  readonly hasRefError: boolean;
}

interface FormulaRewriteResult {
  /** `${sheet}!row:col` → rewrite. Includes both shifted-coord and unchanged-coord cells. */
  readonly byCell: ReadonlyMap<string, FormulaRewrite>;
}

function rewriteWorkbookFormulas(
  workbook: XlsxWorkbook,
  shifted: ShiftedSheet,
  axis: Axis,
  op: Operation,
  at0: number,
  count: number
): FormulaRewriteResult {
  const adjust = pickAdjustFn(axis, op, shifted.nextSheet.name, at0, count);
  const byCell = new Map<string, FormulaRewrite>();
  for (const sheet of workbook.sheets) {
    const isTarget = sheet.name === shifted.nextSheet.name;
    const source = isTarget ? shifted.nextSheet : sheet;
    for (const cell of source.cells.values()) {
      if (!cell.formula) continue;
      const anchor: CellRef = { sheet: source.name, row: cell.row, col: cell.col, abs: 0 };
      let result;
      try {
        result = rewriteFormulaRefs(cell.formula.text, anchor, adjust);
      } catch {
        // Per `EC-F3`: malformed formula already in the workbook — leave verbatim.
        continue;
      }
      if (!result.changed) continue;
      const key = engineCellKey(source.name, cell.row, cell.col);
      byCell.set(key, {
        sheetName: source.name,
        cell,
        oldText: cell.formula.text,
        newText: result.text,
        hasRefError: result.hasRefError,
      });
    }
  }
  return { byCell };
}

function pickAdjustFn(
  axis: Axis,
  op: Operation,
  sheetName: string,
  at0: number,
  count: number
): (ref: CellRef | RangeRef) => CellRef | RangeRef | CellError {
  if (axis === "row" && op === "insert") return (ref) => adjustForInsertRow(ref, sheetName, at0, count);
  if (axis === "row" && op === "delete") return (ref) => adjustForDeleteRow(ref, sheetName, at0, count);
  if (axis === "column" && op === "insert") return (ref) => adjustForInsertColumn(ref, sheetName, at0, count);
  return (ref) => adjustForDeleteColumn(ref, sheetName, at0, count);
}

function applyFormulaRewrites(
  workbook: XlsxWorkbook,
  shifted: ShiftedSheet,
  rewrites: FormulaRewriteResult
): XlsxWorkbook {
  // Start by replacing the target sheet with its shifted copy.
  let next = replaceSheet(workbook, shifted.nextSheet);

  // Group rewrites by sheet so we mutate each sheet's cell map once.
  const bySheet = new Map<string, FormulaRewrite[]>();
  for (const rw of rewrites.byCell.values()) {
    let bucket = bySheet.get(rw.sheetName);
    if (!bucket) {
      bucket = [];
      bySheet.set(rw.sheetName, bucket);
    }
    bucket.push(rw);
  }

  for (const [sheetName, items] of bySheet) {
    const s = next.sheets.find((x) => x.name === sheetName);
    if (!s) continue;
    const newCells = new Map(s.cells);
    for (const rw of items) {
      const k = cellKey(rw.cell.row, rw.cell.col);
      const current = newCells.get(k);
      if (!current) continue;
      const updated: Cell = { ...current, formula: { text: rw.newText } };
      newCells.set(k, updated);
    }
    next = replaceSheet(next, { ...s, cells: newCells });
  }
  return next;
}

// ── Recalc ────────────────────────────────────────────────────────────────

interface RecalcResult {
  readonly values: ReadonlyMap<string, CellValue>;
}

function recalcWorkbook(workbook: XlsxWorkbook): RecalcResult {
  const { engine, host } = bindEngineToWorkbook(workbook);
  // Seed errors are non-fatal (per `EC-F3`).
  void host.seedFormulas(engine);
  const result = engine.recalc();
  const values = new Map<string, CellValue>();
  for (const [key, value] of result.values) {
    values.set(key, fromEngineValue(value));
  }
  return { values };
}

function applyCachedValues(workbook: XlsxWorkbook, values: ReadonlyMap<string, CellValue>): XlsxWorkbook {
  const bySheet = new Map<string, Map<string, CellValue>>();
  for (const [k, v] of values) {
    const parts = splitCellKey(k);
    let bucket = bySheet.get(parts.sheet);
    if (!bucket) {
      bucket = new Map();
      bySheet.set(parts.sheet, bucket);
    }
    bucket.set(cellKey(parts.row, parts.col), v);
  }
  let next = workbook;
  for (const [sheetName, perCell] of bySheet) {
    const s = next.sheets.find((x) => x.name === sheetName);
    if (!s) continue;
    const newCells = new Map(s.cells);
    let changed = false;
    for (const [key, value] of perCell) {
      const cell = newCells.get(key);
      if (!cell || !cell.formula) continue;
      if (cellValuesEqual(cell.value, value)) continue;
      newCells.set(key, { ...cell, value });
      changed = true;
    }
    if (changed) next = replaceSheet(next, { ...s, cells: newCells });
  }
  return next;
}

// ── Diff ──────────────────────────────────────────────────────────────────

function buildChanges(
  prevSnapshot: XlsxSnapshot,
  finalWorkbook: XlsxWorkbook,
  originalSheet: Sheet,
  axis: Axis,
  op: Operation,
  payload: StructuralPayload,
  shifted: ShiftedSheet,
  rewrites: FormulaRewriteResult,
  recalcValues: ReadonlyMap<string, CellValue>
): DiffChange[] {
  const finalSheet = finalWorkbook.sheets.find((s) => s.name === originalSheet.name)!;
  const summaryField = summaryFieldFor(axis, op);
  const verb = op === "insert" ? "inserted" : "deleted";
  const noun = axis === "row" ? "row" : "column";
  const changes: DiffChange[] = [];
  changes.push({
    kind: "node-updated",
    nodeId: finalSheet.id,
    path: ["sheets", finalSheet.index, summaryField],
    field: summaryField,
    summary: `${payload.count} ${noun}${payload.count === 1 ? "" : "s"} ${verb} at ${noun} ${payload.at} on '${originalSheet.name}'`,
    meta: {
      sheet: originalSheet.name,
      at: payload.at,
      count: payload.count,
      shiftedCells: shifted.shiftedCells,
      droppedCells: shifted.droppedCells.length,
    },
  });

  if (op === "delete") {
    for (const drop of shifted.droppedCells) {
      changes.push({
        kind: "node-deleted",
        nodeId: finalSheet.id,
        path: ["sheets", finalSheet.index, "cells", formatA1({ row: drop.row, col: drop.col })],
        summary: `referenced-cell-deleted ${originalSheet.name}!${formatA1({ row: drop.row, col: drop.col })}`,
        meta: {
          field: "referenced-cell-deleted",
          before: { value: drop.before.value, formula: drop.before.formula?.text ?? null },
        },
      });
    }
  }

  for (const rw of rewrites.byCell.values()) {
    const s = finalWorkbook.sheets.find((x) => x.name === rw.sheetName);
    if (!s) continue;
    changes.push({
      kind: "node-updated",
      nodeId: s.id,
      path: [
        "sheets",
        s.index,
        "cells",
        `${rw.sheetName}!${formatA1({ row: rw.cell.row, col: rw.cell.col })}`,
        "formula",
      ],
      field: "formula",
      summary: `${rw.sheetName}!${formatA1({ row: rw.cell.row, col: rw.cell.col })}: =${rw.oldText} → =${rw.newText}`,
      meta: {
        before: rw.oldText,
        after: rw.newText,
        hasRefError: rw.hasRefError,
      },
    });
    if (rw.hasRefError) {
      changes.push({
        kind: "node-updated",
        nodeId: s.id,
        path: [
          "sheets",
          s.index,
          "cells",
          `${rw.sheetName}!${formatA1({ row: rw.cell.row, col: rw.cell.col })}`,
        ],
        field: "referenced-cell-deleted",
        summary: `referenced-cell-deleted in ${rw.sheetName}!${formatA1({ row: rw.cell.row, col: rw.cell.col })} (formula ⇒ #REF!)`,
      });
    }
  }

  // Per-cell `cell-updated` for any dependent (non-rewritten formula) whose
  // cached value changed across the recalc.
  for (const [key, value] of recalcValues) {
    if (rewrites.byCell.has(key)) continue;
    const parts = splitCellKey(key);
    const finalSheetForCell = finalWorkbook.sheets.find((x) => x.name === parts.sheet);
    const previousSheetForCell = prevSnapshot.root.sheets.find((x) => x.name === parts.sheet);
    if (!finalSheetForCell || !previousSheetForCell) continue;
    const before = previousSheetForCell.cells.get(cellKey(parts.row, parts.col));
    const after = finalSheetForCell.cells.get(cellKey(parts.row, parts.col));
    if (!after || !after.formula) continue;
    const beforeValue = before?.value ?? null;
    if (cellValuesEqual(beforeValue, value)) continue;
    changes.push({
      kind: "node-updated",
      nodeId: finalSheetForCell.id,
      path: [
        "sheets",
        finalSheetForCell.index,
        "cells",
        `${parts.sheet}!${formatA1({ row: parts.row, col: parts.col })}`,
      ],
      field: "cachedValue",
      summary: `${parts.sheet}!${formatA1({ row: parts.row, col: parts.col })}: ${formatVal(beforeValue)} → ${formatVal(value)}`,
    });
  }

  return changes;
}

function summaryFieldFor(axis: Axis, op: Operation): string {
  if (axis === "row" && op === "insert") return "rows-inserted";
  if (axis === "row" && op === "delete") return "rows-deleted";
  if (axis === "column" && op === "insert") return "columns-inserted";
  return "columns-deleted";
}

// ── Misc ──────────────────────────────────────────────────────────────────

function engineCellKey(sheet: string, row: number, col: number): string {
  return `${sheet}!${row}:${col}`;
}

function splitCellKey(key: string): { sheet: string; row: number; col: number } {
  const idx = key.lastIndexOf("!");
  const sheet = key.slice(0, idx);
  const rest = key.slice(idx + 1);
  const colon = rest.indexOf(":");
  return { sheet, row: Number(rest.slice(0, colon)), col: Number(rest.slice(colon + 1)) };
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
