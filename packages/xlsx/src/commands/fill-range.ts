import { CommandError, type CommandHandler, type DiffChange } from "@officeai/core";
import { cellsToSamples, pickSeries, type SeriesSample } from "../fill/series.js";
import { bindEngineToWorkbook, fromEngineValue } from "../formula/workbook-host.js";
import { rewriteFormulaRefs, type AdjustFn } from "../formula/rewrite-refs.js";
import type { CellRef, RangeRef } from "../formula/references.js";
import { cellKey, formatA1 } from "../model/refs.js";
import type { Cell, Sheet, XlsxSnapshot } from "../model/types.js";
import { buildDiff, evolveSnapshot, replaceSheet } from "./helpers.js";
import type { FillRangePayload } from "./payloads.js";
import { parseRangeRef, resolveSheet } from "./validation.js";

/**
 * `xlsx:fill-range` — Excel's fill handle (drag the bottom-right
 * corner of the selection to extend a series).
 *
 * Spec: `spec/xlsx/agent-commands.md` §15.
 *
 * Inputs:
 *   - `source`  : the originally selected rectangle (1-D in the
 *     fill direction, but allowed to be ≥ 1 in the orthogonal
 *     direction — we run an independent series per orthogonal lane).
 *   - `target`  : the FULL extended rectangle (must enclose `source`).
 *   - `direction` : "down" | "right" | "up" | "left".
 *
 * Algorithm:
 *   1. Validate direction against the rectangle deltas.
 *   2. For each "lane" (row when filling left/right, column when
 *      filling up/down), pull cells from the source, run the
 *      detector pipeline, and emit values into the new cells.
 *   3. Cells with formulas always go through the formula path
 *      (re-anchored via `rewriteFormulaRefs` against destination).
 *   4. After writing, recalc once so any new formulas pick up
 *      cached values.
 */
export const fillRangeHandler: CommandHandler<FillRangePayload, XlsxSnapshot> = {
  type: "xlsx:fill-range",
  apply(snapshot, payload) {
    const sheet = resolveSheet(snapshot.root, payload.sheet);
    const source = parseRangeRef(payload.source);
    const target = parseRangeRef(payload.target);
    const direction = payload.direction;

    if (
      target.start.row > source.start.row ||
      target.start.col > source.start.col ||
      target.end.row < source.end.row ||
      target.end.col < source.end.col
    ) {
      throw new CommandError(
        "invalid-range",
        `target ${payload.target} must fully contain source ${payload.source}`
      );
    }

    const dTop = source.start.row - target.start.row;
    const dBottom = target.end.row - source.end.row;
    const dLeft = source.start.col - target.start.col;
    const dRight = target.end.col - source.end.col;

    if (
      (direction === "down" && (dBottom <= 0 || dTop !== 0 || dLeft !== 0 || dRight !== 0)) ||
      (direction === "up" && (dTop <= 0 || dBottom !== 0 || dLeft !== 0 || dRight !== 0)) ||
      (direction === "right" && (dRight <= 0 || dTop !== 0 || dBottom !== 0 || dLeft !== 0)) ||
      (direction === "left" && (dLeft <= 0 || dTop !== 0 || dBottom !== 0 || dRight !== 0))
    ) {
      throw new CommandError(
        "invalid-range",
        `direction "${direction}" inconsistent with source/target geometry`
      );
    }

    const cells = new Map(sheet.cells);
    const changes: DiffChange[] = [];
    const formulaTargets: Array<{ row: number; col: number; text: string }> = [];

    if (direction === "down" || direction === "up") {
      // One lane per source column.
      for (let col = source.start.col; col <= source.end.col; col++) {
        const isDown = direction === "down";
        const sourceCells: (Cell | undefined)[] = [];
        const rowsAsc: number[] = [];
        for (let r = source.start.row; r <= source.end.row; r++) {
          rowsAsc.push(r);
          sourceCells.push(cells.get(cellKey(r, col)));
        }
        const samples = isDown
          ? cellsToSamples(sourceCells)
          : cellsToSamples([...sourceCells].reverse());
        runLane({
          sheet,
          cells,
          samples,
          sourceCells: isDown ? sourceCells : [...sourceCells].reverse(),
          sourceRowsOrCols: isDown ? rowsAsc : [...rowsAsc].reverse(),
          newCoords: isDown
            ? rangeAsc(source.end.row + 1, target.end.row).map((r) => ({ row: r, col }))
            : rangeAsc(target.start.row, source.start.row - 1)
                .map((r) => ({ row: r, col }))
                .reverse(),
          axis: "row",
          changes,
          formulaTargets,
        });
      }
    } else {
      // One lane per source row.
      for (let row = source.start.row; row <= source.end.row; row++) {
        const isRight = direction === "right";
        const sourceCells: (Cell | undefined)[] = [];
        const colsAsc: number[] = [];
        for (let c = source.start.col; c <= source.end.col; c++) {
          colsAsc.push(c);
          sourceCells.push(cells.get(cellKey(row, c)));
        }
        const samples = isRight
          ? cellsToSamples(sourceCells)
          : cellsToSamples([...sourceCells].reverse());
        runLane({
          sheet,
          cells,
          samples,
          sourceCells: isRight ? sourceCells : [...sourceCells].reverse(),
          sourceRowsOrCols: isRight ? colsAsc : [...colsAsc].reverse(),
          newCoords: isRight
            ? rangeAsc(source.end.col + 1, target.end.col).map((c) => ({ row, col: c }))
            : rangeAsc(target.start.col, source.start.col - 1)
                .map((c) => ({ row, col: c }))
                .reverse(),
          axis: "col",
          changes,
          formulaTargets,
        });
      }
    }

    let nextSheet: Sheet = { ...sheet, cells };
    let nextWorkbook = replaceSheet(snapshot.root, nextSheet);

    if (formulaTargets.length > 0) {
      const { engine, host } = bindEngineToWorkbook(nextWorkbook);
      void host.seedFormulas(engine);
      const result = engine.recalc();
      const updates = new Map(cells);
      for (const [key, value] of result.values) {
        const sep = key.indexOf("!");
        if (sep === -1) continue;
        const sName = key.slice(0, sep);
        if (sName !== sheet.name) continue;
        const coord = key.slice(sep + 1);
        const colon = coord.indexOf(":");
        if (colon === -1) continue;
        const r = Number.parseInt(coord.slice(0, colon), 10);
        const c = Number.parseInt(coord.slice(colon + 1), 10);
        if (!Number.isFinite(r) || !Number.isFinite(c)) continue;
        const existing = updates.get(cellKey(r, c));
        if (!existing || !existing.formula) continue;
        const next = fromEngineValue(value);
        if (cellValuesEqual(existing.value, next)) continue;
        updates.set(cellKey(r, c), { ...existing, value: next });
      }
      nextSheet = { ...nextSheet, cells: updates };
      nextWorkbook = replaceSheet(nextWorkbook, nextSheet);
    }

    if (changes.length === 0) {
      const next = evolveSnapshot(snapshot, snapshot.root, {});
      return {
        next,
        diff: buildDiff(snapshot.revision, next.revision, [
          {
            kind: "node-updated",
            nodeId: sheet.id,
            path: ["sheets", sheet.index],
            field: "noop",
            summary: `fill ${direction} from ${payload.source} to ${payload.target}: no changes`,
          },
        ]),
      };
    }

    const next = evolveSnapshot(snapshot, nextWorkbook, { sheets: [sheet.partPath] });
    return { next, diff: buildDiff(snapshot.revision, next.revision, changes) };
  },
};

interface LaneCtx {
  readonly sheet: Sheet;
  readonly cells: Map<string, Cell>;
  readonly samples: SeriesSample[];
  readonly sourceCells: (Cell | undefined)[];
  readonly sourceRowsOrCols: number[];
  readonly newCoords: ReadonlyArray<{ row: number; col: number }>;
  readonly axis: "row" | "col";
  readonly changes: DiffChange[];
  readonly formulaTargets: Array<{ row: number; col: number; text: string }>;
}

function runLane(ctx: LaneCtx): void {
  if (ctx.newCoords.length === 0) return;
  const allFormulas = ctx.sourceCells.every((c) => !!c?.formula);
  const generator = pickSeries(ctx.samples);

  for (let i = 0; i < ctx.newCoords.length; i++) {
    const { row, col } = ctx.newCoords[i]!;
    const offset = i + 1;
    const before = ctx.cells.get(cellKey(row, col));
    let formulaText: string | null = null;
    let value: Cell["value"] = null;

    if (allFormulas) {
      // Cycle through the source formulas and re-anchor.
      const srcIdx = (offset - 1) % ctx.sourceCells.length;
      const src = ctx.sourceCells[srcIdx]!;
      const sr = ctx.axis === "row" ? ctx.sourceRowsOrCols[srcIdx]! : row;
      const sc = ctx.axis === "col" ? ctx.sourceRowsOrCols[srcIdx]! : col;
      formulaText = shiftFormula(
        src.formula!.text,
        ctx.sheet.name,
        sr,
        sc,
        row,
        col
      );
      ctx.formulaTargets.push({ row, col, text: formulaText });
    } else {
      value = generator.next(offset);
    }

    const next: Cell = {
      row,
      col,
      value,
      ...(formulaText !== null ? { formula: { text: formulaText } } : {}),
      ...(before?.styleId !== undefined ? { styleId: before.styleId } : {}),
    };
    if (cellsEqual(before, next)) continue;
    ctx.cells.set(cellKey(row, col), next);
    ctx.changes.push(diffChange(ctx.sheet, row, col, before?.value ?? null, next.value, formulaText));
  }
}

function shiftFormula(
  text: string,
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
  const anchor: CellRef = { sheet: destSheet, row: dr, col: dc, abs: 0 };
  return rewriteFormulaRefs(text, anchor, adjust).text;
}

function rangeAsc(start: number, end: number): number[] {
  const out: number[] = [];
  for (let i = start; i <= end; i++) out.push(i);
  return out;
}

function cellsEqual(a: Cell | undefined, b: Cell): boolean {
  if (!a) return b.value === null && !b.formula;
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

function diffChange(
  sheet: Sheet,
  row: number,
  col: number,
  beforeVal: Cell["value"],
  afterVal: Cell["value"],
  afterFormula: string | null
): DiffChange {
  return {
    kind: "node-updated",
    nodeId: sheet.id,
    path: ["sheets", sheet.index, "cells", `${sheet.name}!${formatA1({ row, col })}`],
    field: afterFormula !== null ? "formula" : "value",
    summary: `${formatA1({ row, col })}: ${formatVal(beforeVal)} → ${formatVal(afterVal)}${
      afterFormula !== null ? ` (=${afterFormula})` : ""
    }`,
    meta: {
      before: { value: beforeVal },
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
