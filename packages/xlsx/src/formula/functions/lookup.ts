import type { AstNode } from "../ast.js";
import { ErrorKinds } from "../errors.js";
import {
  arity,
  type EvalContext,
  type LazyEvalAccess,
  type MutableFunctionRegistry,
} from "../function-registry.js";
import { AbsRef, parseA1, parseA1Range, type CellRef, type RangeRef } from "../references.js";
import {
  compare,
  err,
  isError,
  num,
  rangeVal,
  toBoolean,
  toNumber,
  toString,
  type Range2D,
  type Value,
} from "../values.js";

/**
 * Lookup category — `VLOOKUP`, `HLOOKUP`, `INDEX`, `MATCH`, `XLOOKUP`,
 * `CHOOSE`, `OFFSET`, `INDIRECT`, `ROW`, `ROWS`, `COLUMN`, `COLUMNS`.
 *
 * Spec: `spec/xlsx/formula-engine.md` §16.3 (function inventory) and
 * §13.4 (volatility / dynamic-ref handling for `INDIRECT` / `OFFSET`).
 *
 * Volatility:
 *
 * - `INDIRECT` and `OFFSET` are registered with `volatile: true`. The
 *   parser also surfaces this via `containsVolatile` (see
 *   `ast.VOLATILE_FUNCTIONS`) so the recalc loop force-dirties their
 *   cells every recalc — Excel parity, per §13.4.
 *
 * Lazy args:
 *
 * - `OFFSET`, `ROW`, `ROWS`, `COLUMN`, `COLUMNS` declare
 *   `lazyArgs: true`. They peek at the source AST to recover the
 *   underlying `RefNode` / `RangeRefNode` *before* eager evaluation
 *   collapses it to a `Value`. This is essential for `OFFSET` to
 *   compute a new ref relative to the original and for
 *   `ROW`/`COLUMN` to report the operand's grid coordinates.
 */

const MAX_ROW = 1048575;
const MAX_COL = 16383;

interface BaseRect {
  readonly sheet: string;
  readonly r0: number;
  readonly c0: number;
  readonly r1: number;
  readonly c1: number;
}

function rectFromRefNode(node: AstNode): BaseRect | undefined {
  switch (node.kind) {
    case "ref":
      return {
        sheet: node.ref.sheet,
        r0: node.ref.row,
        c0: node.ref.col,
        r1: node.ref.row,
        c1: node.ref.col,
      };
    case "range":
      return {
        sheet: node.ref.sheet,
        r0: node.ref.r0,
        c0: node.ref.c0,
        r1: node.ref.r1,
        c1: node.ref.c1,
      };
    default:
      return undefined;
  }
}

function asGrid(v: Value): Range2D {
  return v.kind === "r" ? v.v : [[v]];
}

function gridDims(g: Range2D): { rows: number; cols: number } {
  return { rows: g.length, cols: g.length === 0 ? 0 : g[0].length };
}

function valuesEqual(a: Value, b: Value): boolean {
  const c = compare(a, b);
  return typeof c === "number" && c === 0;
}

// ── VLOOKUP / HLOOKUP ─────────────────────────────────────────────────────

function lookupExact(table: Range2D, key: Value, axis: "v" | "h", outIdx: number): Value {
  const major = axis === "v" ? table.length : (table[0]?.length ?? 0);
  for (let i = 0; i < major; i++) {
    const probe = axis === "v" ? table[i][0] : table[0][i];
    if (probe.kind === "e") return probe;
    if (valuesEqual(probe, key)) {
      return axis === "v" ? table[i][outIdx] : table[outIdx][i];
    }
  }
  return err(ErrorKinds.NA);
}

function lookupApprox(table: Range2D, key: Value, axis: "v" | "h", outIdx: number): Value {
  const major = axis === "v" ? table.length : (table[0]?.length ?? 0);
  let lo = 0;
  let hi = major - 1;
  let best = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const probe = axis === "v" ? table[mid][0] : table[0][mid];
    if (probe.kind === "e") return probe;
    const c = compare(probe, key);
    if (typeof c !== "number") return c;
    if (c <= 0) {
      best = mid;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  if (best < 0) return err(ErrorKinds.NA);
  return axis === "v" ? table[best][outIdx] : table[outIdx][best];
}

function vlookupLike(args: ReadonlyArray<Value>, axis: "v" | "h"): Value {
  const [key, table, idxArg, modeArg] = args;
  if (key.kind === "e") return key;
  if (table.kind === "e") return table;
  if (idxArg.kind === "e") return idxArg;
  if (modeArg !== undefined && modeArg.kind === "e") return modeArg;
  if (table.kind !== "r") return err(ErrorKinds.VALUE);
  const grid = table.v;
  if (grid.length === 0 || grid[0].length === 0) return err(ErrorKinds.NA);

  const idxN = toNumber(idxArg);
  if (isError(idxN)) return idxN;
  const idx = Math.trunc(idxN.v);
  if (idx < 1) return err(ErrorKinds.VALUE);
  const limit = axis === "v" ? grid[0].length : grid.length;
  if (idx > limit) return err(ErrorKinds.REF);

  let approximate = true;
  if (modeArg !== undefined) {
    const b = toBoolean(modeArg);
    if (isError(b)) return b;
    approximate = b.v;
  }

  return approximate ? lookupApprox(grid, key, axis, idx - 1) : lookupExact(grid, key, axis, idx - 1);
}

// ── INDEX ─────────────────────────────────────────────────────────────────

function indexFn(args: ReadonlyArray<Value>): Value {
  const [range, rowArg, colArg] = args;
  if (range.kind === "e") return range;
  if (rowArg.kind === "e") return rowArg;
  if (colArg !== undefined && colArg.kind === "e") return colArg;

  const grid = asGrid(range);
  const { rows, cols } = gridDims(grid);
  if (rows === 0 || cols === 0) return err(ErrorKinds.REF);

  const rowN = toNumber(rowArg);
  if (isError(rowN)) return rowN;
  const r = Math.trunc(rowN.v);

  let c: number;
  if (colArg !== undefined) {
    const colN = toNumber(colArg);
    if (isError(colN)) return colN;
    c = Math.trunc(colN.v);
  } else if (rows === 1) {
    // Single-row range: interpret the lone index as a column position.
    c = r;
    return indexLookup(grid, 1, c, rows, cols);
  } else if (cols === 1) {
    c = 1;
    return indexLookup(grid, r, c, rows, cols);
  } else {
    // 2-D range, single index → return the entire row at index `r`
    // (Excel's documented behaviour).
    c = 0;
  }

  return indexLookup(grid, r, c, rows, cols);
}

function indexLookup(grid: Range2D, r: number, c: number, rows: number, cols: number): Value {
  if (r < 0 || c < 0) return err(ErrorKinds.VALUE);
  if (r > rows || c > cols) return err(ErrorKinds.REF);

  if (r === 0 && c === 0) return rangeVal(grid);
  if (r === 0) {
    const out: Value[][] = [];
    for (const row of grid) out.push([row[c - 1]]);
    return rangeVal(out);
  }
  if (c === 0) {
    return rangeVal([[...grid[r - 1]]]);
  }
  return grid[r - 1][c - 1];
}

// ── MATCH ─────────────────────────────────────────────────────────────────

function flatten1D(grid: Range2D): Value[] {
  const { rows, cols } = gridDims(grid);
  const out: Value[] = [];
  if (cols === 1) {
    for (let i = 0; i < rows; i++) out.push(grid[i][0]);
  } else {
    for (let i = 0; i < cols; i++) out.push(grid[0][i]);
  }
  return out;
}

function matchFn(args: ReadonlyArray<Value>): Value {
  const [lookup, range, typeArg] = args;
  if (lookup.kind === "e") return lookup;
  if (range.kind === "e") return range;
  if (typeArg !== undefined && typeArg.kind === "e") return typeArg;

  let matchType: -1 | 0 | 1 = 1;
  if (typeArg !== undefined) {
    const n = toNumber(typeArg);
    if (isError(n)) return n;
    matchType = n.v > 0 ? 1 : n.v < 0 ? -1 : 0;
  }

  const arr = flatten1D(asGrid(range));
  if (arr.length === 0) return err(ErrorKinds.NA);

  switch (matchType) {
    case 0: {
      for (let i = 0; i < arr.length; i++) {
        const cell = arr[i];
        if (cell.kind === "e") continue;
        if (valuesEqual(cell, lookup)) return num(i + 1);
      }
      return err(ErrorKinds.NA);
    }
    case 1: {
      let best = -1;
      for (let i = 0; i < arr.length; i++) {
        const cell = arr[i];
        if (cell.kind === "e") return cell;
        const c = compare(cell, lookup);
        if (typeof c !== "number") return c;
        if (c <= 0) best = i;
        else break;
      }
      if (best < 0) return err(ErrorKinds.NA);
      return num(best + 1);
    }
    case -1: {
      let best = -1;
      for (let i = 0; i < arr.length; i++) {
        const cell = arr[i];
        if (cell.kind === "e") return cell;
        const c = compare(cell, lookup);
        if (typeof c !== "number") return c;
        if (c >= 0) best = i;
        else break;
      }
      if (best < 0) return err(ErrorKinds.NA);
      return num(best + 1);
    }
  }
}

// ── XLOOKUP ───────────────────────────────────────────────────────────────

function xlookupFn(args: ReadonlyArray<Value>): Value {
  const [lookup, lookupArr, returnArr, ifNotFound, matchModeArg, searchModeArg] = args;
  if (lookup.kind === "e") return lookup;
  if (lookupArr.kind === "e") return lookupArr;
  if (returnArr.kind === "e") return returnArr;
  if (ifNotFound !== undefined && ifNotFound.kind === "e") return ifNotFound;
  if (matchModeArg !== undefined && matchModeArg.kind === "e") return matchModeArg;
  if (searchModeArg !== undefined && searchModeArg.kind === "e") return searchModeArg;

  const lookupGrid = asGrid(lookupArr);
  const returnGrid = asGrid(returnArr);
  const lookupDims = gridDims(lookupGrid);
  if (lookupDims.rows === 0 || lookupDims.cols === 0) return err(ErrorKinds.NA);
  const isRow = lookupDims.rows === 1;
  const lookupSeries = flatten1D(lookupGrid);

  let matchMode = 0;
  if (matchModeArg !== undefined) {
    const n = toNumber(matchModeArg);
    if (isError(n)) return n;
    matchMode = Math.trunc(n.v);
  }
  let searchMode = 1;
  if (searchModeArg !== undefined) {
    const n = toNumber(searchModeArg);
    if (isError(n)) return n;
    searchMode = Math.trunc(n.v);
  }

  const order: number[] = [];
  if (searchMode === -1) {
    for (let i = lookupSeries.length - 1; i >= 0; i--) order.push(i);
  } else {
    for (let i = 0; i < lookupSeries.length; i++) order.push(i);
  }

  let found = -1;
  let bestIdx = -1;
  for (const i of order) {
    const cell = lookupSeries[i];
    if (cell.kind === "e") continue;
    const c = compare(cell, lookup);
    if (typeof c !== "number") continue;
    if (c === 0) {
      found = i;
      break;
    }
    if (matchMode === -1 && c < 0) {
      // exact or next smaller (largest probe ≤ lookup)
      if (bestIdx < 0) {
        bestIdx = i;
      } else {
        const bcmp = compare(cell, lookupSeries[bestIdx]);
        if (typeof bcmp === "number" && bcmp > 0) bestIdx = i;
      }
    } else if (matchMode === 1 && c > 0) {
      // exact or next larger (smallest probe ≥ lookup)
      if (bestIdx < 0) {
        bestIdx = i;
      } else {
        const bcmp = compare(cell, lookupSeries[bestIdx]);
        if (typeof bcmp === "number" && bcmp < 0) bestIdx = i;
      }
    }
  }

  const idx = found >= 0 ? found : bestIdx;
  if (idx < 0) {
    if (ifNotFound !== undefined) return ifNotFound;
    return err(ErrorKinds.NA);
  }

  const returnDims = gridDims(returnGrid);
  if (isRow) {
    if (idx >= returnDims.cols) return err(ErrorKinds.REF);
    if (returnDims.rows === 1) return returnGrid[0][idx];
    const col: Value[][] = [];
    for (const row of returnGrid) col.push([row[idx]]);
    return rangeVal(col);
  }
  if (idx >= returnDims.rows) return err(ErrorKinds.REF);
  if (returnDims.cols === 1) return returnGrid[idx][0];
  return rangeVal([[...returnGrid[idx]]]);
}

// ── CHOOSE ────────────────────────────────────────────────────────────────

function chooseFn(args: ReadonlyArray<Value>): Value {
  const idxArg = args[0];
  if (idxArg.kind === "e") return idxArg;
  const n = toNumber(idxArg);
  if (isError(n)) return n;
  const idx = Math.trunc(n.v);
  if (idx < 1 || idx > args.length - 1) return err(ErrorKinds.VALUE);
  return args[idx];
}

// ── OFFSET (lazy) ─────────────────────────────────────────────────────────

function offsetFn(args: ReadonlyArray<AstNode>, lazy: LazyEvalAccess): Value {
  const baseRect = rectFromRefNode(args[0]);
  if (!baseRect) {
    // P0 limitation: OFFSET requires its base to be a literal cell or
    // range reference (the only forms that survive eager-eval intact).
    // Computed bases (e.g. `OFFSET(INDIRECT("A1"), …)`) are deferred.
    return err(ErrorKinds.VALUE);
  }

  const rowOff = lazy.evaluate(args[1]);
  if (rowOff.kind === "e") return rowOff;
  const rowN = toNumber(rowOff);
  if (isError(rowN)) return rowN;

  const colOff = lazy.evaluate(args[2]);
  if (colOff.kind === "e") return colOff;
  const colN = toNumber(colOff);
  if (isError(colN)) return colN;

  let height = baseRect.r1 - baseRect.r0 + 1;
  let width = baseRect.c1 - baseRect.c0 + 1;

  if (args.length >= 4) {
    const h = lazy.evaluate(args[3]);
    if (h.kind === "e") return h;
    const hN = toNumber(h);
    if (isError(hN)) return hN;
    height = Math.trunc(hN.v);
  }
  if (args.length >= 5) {
    const w = lazy.evaluate(args[4]);
    if (w.kind === "e") return w;
    const wN = toNumber(w);
    if (isError(wN)) return wN;
    width = Math.trunc(wN.v);
  }

  if (height <= 0 || width <= 0) return err(ErrorKinds.REF);

  const r0 = baseRect.r0 + Math.trunc(rowN.v);
  const c0 = baseRect.c0 + Math.trunc(colN.v);
  const r1 = r0 + height - 1;
  const c1 = c0 + width - 1;
  if (r0 < 0 || c0 < 0 || r1 > MAX_ROW || c1 > MAX_COL) return err(ErrorKinds.REF);

  if (height === 1 && width === 1) {
    const ref: CellRef = { sheet: baseRect.sheet, row: r0, col: c0, abs: AbsRef.NONE };
    return lazy.ctx.getCell(ref);
  }
  const ref: RangeRef = { sheet: baseRect.sheet, r0, c0, r1, c1, abs0: AbsRef.NONE, abs1: AbsRef.NONE };
  return rangeVal(lazy.ctx.getRange(ref));
}

// ── INDIRECT ──────────────────────────────────────────────────────────────

function indirectFn(args: ReadonlyArray<Value>, ctx: EvalContext): Value {
  const [textArg, a1Arg] = args;
  if (textArg.kind === "e") return textArg;
  if (a1Arg !== undefined && a1Arg.kind === "e") return a1Arg;

  const s = toString(textArg);
  if (isError(s)) return s;

  let a1Style = true;
  if (a1Arg !== undefined) {
    const b = toBoolean(a1Arg);
    if (isError(b)) return b;
    a1Style = b.v;
  }
  // R1C1 text is not supported in P0; doc'd deferral.
  if (!a1Style) return err(ErrorKinds.REF);

  const sheet = ctx.anchor.sheet;
  const range = parseA1Range(s.v, sheet);
  if (range) return rangeVal(ctx.getRange(range));
  const cell = parseA1(s.v, sheet);
  if (cell) return ctx.getCell(cell);
  return err(ErrorKinds.REF);
}

// ── ROW / COLUMN / ROWS / COLUMNS (lazy) ──────────────────────────────────

function rowFn(args: ReadonlyArray<AstNode>, lazy: LazyEvalAccess): Value {
  if (args.length === 0) return num(lazy.ctx.anchor.row + 1);
  const a = args[0];
  if (a.kind === "ref") return num(a.ref.row + 1);
  if (a.kind === "range") return num(a.ref.r0 + 1);
  return err(ErrorKinds.VALUE);
}

function columnFn(args: ReadonlyArray<AstNode>, lazy: LazyEvalAccess): Value {
  if (args.length === 0) return num(lazy.ctx.anchor.col + 1);
  const a = args[0];
  if (a.kind === "ref") return num(a.ref.col + 1);
  if (a.kind === "range") return num(a.ref.c0 + 1);
  return err(ErrorKinds.VALUE);
}

function rowsFn(args: ReadonlyArray<AstNode>, lazy: LazyEvalAccess): Value {
  const a = args[0];
  if (a.kind === "ref") return num(1);
  if (a.kind === "range") return num(a.ref.r1 - a.ref.r0 + 1);
  const v = lazy.evaluate(a);
  if (v.kind === "e") return v;
  if (v.kind === "r") return num(v.v.length);
  return num(1);
}

function columnsFn(args: ReadonlyArray<AstNode>, lazy: LazyEvalAccess): Value {
  const a = args[0];
  if (a.kind === "ref") return num(1);
  if (a.kind === "range") return num(a.ref.c1 - a.ref.c0 + 1);
  const v = lazy.evaluate(a);
  if (v.kind === "e") return v;
  if (v.kind === "r") return num(v.v[0]?.length ?? 0);
  return num(1);
}

// ── Registration ──────────────────────────────────────────────────────────

export function registerLookup(reg: MutableFunctionRegistry): void {
  reg.register({
    name: "VLOOKUP",
    arity: arity(3, 4),
    fn: (args) => vlookupLike(args, "v"),
  });
  reg.register({
    name: "HLOOKUP",
    arity: arity(3, 4),
    fn: (args) => vlookupLike(args, "h"),
  });
  reg.register({
    name: "INDEX",
    arity: arity(2, 3),
    fn: (args) => indexFn(args),
  });
  reg.register({
    name: "MATCH",
    arity: arity(2, 3),
    fn: (args) => matchFn(args),
  });
  reg.register({
    name: "XLOOKUP",
    arity: arity(3, 6),
    fn: (args) => xlookupFn(args),
  });
  reg.register({
    name: "CHOOSE",
    arity: arity(2, 254),
    fn: (args) => chooseFn(args),
  });
  reg.register({
    name: "OFFSET",
    arity: arity(3, 5),
    volatile: true,
    lazyArgs: true,
    fn: (args, lazy) => offsetFn(args, lazy),
  });
  reg.register({
    name: "INDIRECT",
    arity: arity(1, 2),
    volatile: true,
    fn: (args, ctx) => indirectFn(args, ctx),
  });
  reg.register({
    name: "ROW",
    arity: arity(0, 1),
    lazyArgs: true,
    fn: (args, lazy) => rowFn(args, lazy),
  });
  reg.register({
    name: "ROWS",
    arity: arity(1, 1),
    lazyArgs: true,
    fn: (args, lazy) => rowsFn(args, lazy),
  });
  reg.register({
    name: "COLUMN",
    arity: arity(0, 1),
    lazyArgs: true,
    fn: (args, lazy) => columnFn(args, lazy),
  });
  reg.register({
    name: "COLUMNS",
    arity: arity(1, 1),
    lazyArgs: true,
    fn: (args, lazy) => columnsFn(args, lazy),
  });
}
