import { describe, expect, it } from "vitest";
import { evaluate } from "../../evaluator.js";
import { createRegistry, type EvalContext, type FunctionRegistry } from "../../function-registry.js";
import { registerLookup } from "../../functions/lookup.js";
import { parse } from "../../parser.js";
import { AbsRef, type CellRef, type RangeRef } from "../../references.js";
import { isError, num, str, type Range2D, type Value } from "../../values.js";

const SHEET = "Sheet1";
const ANCHOR: CellRef = { sheet: SHEET, row: 0, col: 0, abs: AbsRef.NONE };

interface CtxOptions {
  cells?: Record<string, Value>;
  ranges?: Record<string, Value[][]>;
  anchor?: CellRef;
}

function rangeKey(ref: RangeRef): string {
  return `${ref.sheet}!R${ref.r0}C${ref.c0}:R${ref.r1}C${ref.c1}`;
}

function cellKey(ref: CellRef): string {
  return `${ref.sheet}!R${ref.row}C${ref.col}`;
}

function makeRegistry(): FunctionRegistry {
  const reg = createRegistry();
  registerLookup(reg);
  return reg;
}

function makeContext(opts: CtxOptions = {}): EvalContext {
  const cells = opts.cells ?? {};
  const ranges = opts.ranges ?? {};
  const reg = makeRegistry();
  return {
    getCell(ref) {
      // If the explicit cell is registered, use it; otherwise see if it
      // falls inside a registered range so OFFSET / INDIRECT can pluck a
      // single cell out of the same fixture used for range tests.
      const direct = cells[cellKey(ref)];
      if (direct !== undefined) return direct;
      for (const [k, grid] of Object.entries(ranges)) {
        const m = /^([^!]+)!R(\d+)C(\d+):R(\d+)C(\d+)$/.exec(k);
        if (!m) continue;
        if (m[1] !== ref.sheet) continue;
        const r0 = Number(m[2]);
        const c0 = Number(m[3]);
        const r1 = Number(m[4]);
        const c1 = Number(m[5]);
        if (ref.row >= r0 && ref.row <= r1 && ref.col >= c0 && ref.col <= c1) {
          return grid[ref.row - r0][ref.col - c0];
        }
      }
      return num(0);
    },
    getRange(ref) {
      const direct = ranges[rangeKey(ref)];
      if (direct) return direct;
      // Synthesise from cells fixture for arbitrary OFFSET requests.
      const out: Value[][] = [];
      for (let r = ref.r0; r <= ref.r1; r++) {
        const row: Value[] = [];
        for (let c = ref.c0; c <= ref.c1; c++) {
          const k = cellKey({ sheet: ref.sheet, row: r, col: c, abs: AbsRef.NONE });
          row.push(cells[k] ?? num(0));
        }
        out.push(row);
      }
      return out;
    },
    resolveName: () => undefined,
    now: () => 0,
    random: () => 0,
    anchor: opts.anchor ?? ANCHOR,
    registry: reg,
  };
}

function evalText(text: string, ctx: EvalContext): Value {
  const ast = parse(text, { anchor: ctx.anchor }).ast;
  return evaluate(ast, ctx);
}

const TABLE_RANGE = "Sheet1!R0C0:R3C1";
const TABLE: Range2D = [
  [str("A"), num(10)],
  [str("B"), num(20)],
  [str("C"), num(30)],
  [str("D"), num(40)],
];

const HTABLE_RANGE = "Sheet1!R0C0:R1C3";
const HTABLE: Range2D = [
  [str("A"), str("B"), str("C"), str("D")],
  [num(10), num(20), num(30), num(40)],
];

const ASC_RANGE = "Sheet1!R0C0:R3C0";
const ASC: Range2D = [[num(10)], [num(20)], [num(30)], [num(40)]];

const DESC_RANGE = "Sheet1!R0C0:R3C0";
const DESC: Range2D = [[num(40)], [num(30)], [num(20)], [num(10)]];

const GRID_RANGE = "Sheet1!R0C0:R2C2";
const GRID: Range2D = [
  [num(1), num(2), num(3)],
  [num(4), num(5), num(6)],
  [num(7), num(8), num(9)],
];

// ── VLOOKUP ───────────────────────────────────────────────────────────────

describe("VLOOKUP", () => {
  const ctx = makeContext({ ranges: { [TABLE_RANGE]: TABLE.map((r) => [...r]) } });

  it("exact match returns the value in the target column", () => {
    expect(evalText('=VLOOKUP("B", A1:B4, 2, FALSE)', ctx)).toEqual(num(20));
  });

  it("exact-match miss returns #N/A", () => {
    const out = evalText('=VLOOKUP("Z", A1:B4, 2, FALSE)', ctx);
    expect(isError(out) && out.v.kind).toBe("#N/A");
  });

  it("approximate match (default) returns largest key ≤ lookup", () => {
    const numericRange = "Sheet1!R0C0:R3C1";
    const numericCtx = makeContext({
      ranges: {
        [numericRange]: [
          [num(10), str("low")],
          [num(20), str("mid")],
          [num(30), str("high")],
          [num(40), str("max")],
        ],
      },
    });
    expect(evalText("=VLOOKUP(25, A1:B4, 2)", numericCtx)).toEqual(str("mid"));
    expect(evalText("=VLOOKUP(40, A1:B4, 2)", numericCtx)).toEqual(str("max"));
  });

  it("col index out of range returns #REF!", () => {
    const out = evalText('=VLOOKUP("A", A1:B4, 5, FALSE)', ctx);
    expect(isError(out) && out.v.kind).toBe("#REF!");
  });
});

// ── HLOOKUP ───────────────────────────────────────────────────────────────

describe("HLOOKUP", () => {
  const ctx = makeContext({ ranges: { [HTABLE_RANGE]: HTABLE.map((r) => [...r]) } });

  it("exact match returns value from the target row", () => {
    expect(evalText('=HLOOKUP("C", A1:D2, 2, FALSE)', ctx)).toEqual(num(30));
  });

  it("miss returns #N/A on exact match", () => {
    const out = evalText('=HLOOKUP("Z", A1:D2, 2, FALSE)', ctx);
    expect(isError(out) && out.v.kind).toBe("#N/A");
  });

  it("approximate match works on numeric headers", () => {
    const ctx2 = makeContext({
      ranges: {
        "Sheet1!R0C0:R1C3": [
          [num(10), num(20), num(30), num(40)],
          [str("a"), str("b"), str("c"), str("d")],
        ],
      },
    });
    expect(evalText("=HLOOKUP(25, A1:D2, 2)", ctx2)).toEqual(str("b"));
  });
});

// ── INDEX ─────────────────────────────────────────────────────────────────

describe("INDEX", () => {
  const ctx = makeContext({ ranges: { [GRID_RANGE]: GRID.map((r) => [...r]) } });

  it("returns the cell at (row, col)", () => {
    expect(evalText("=INDEX(A1:C3, 2, 2)", ctx)).toEqual(num(5));
    expect(evalText("=INDEX(A1:C3, 1, 3)", ctx)).toEqual(num(3));
  });

  it("row=0 returns the entire nth column", () => {
    const out = evalText("=INDEX(A1:C3, 0, 2)", ctx);
    expect(out.kind).toBe("r");
    if (out.kind === "r") {
      expect(out.v).toEqual([[num(2)], [num(5)], [num(8)]]);
    }
  });

  it("col=0 returns the entire nth row", () => {
    const out = evalText("=INDEX(A1:C3, 2, 0)", ctx);
    expect(out.kind).toBe("r");
    if (out.kind === "r") {
      expect(out.v).toEqual([[num(4), num(5), num(6)]]);
    }
  });

  it("out-of-bounds index returns #REF!", () => {
    const out = evalText("=INDEX(A1:C3, 5, 1)", ctx);
    expect(isError(out) && out.v.kind).toBe("#REF!");
  });
});

// ── MATCH ─────────────────────────────────────────────────────────────────

describe("MATCH", () => {
  it("type 0 (exact) returns 1-based position", () => {
    const ctx = makeContext({ ranges: { [ASC_RANGE]: ASC.map((r) => [...r]) } });
    expect(evalText("=MATCH(30, A1:A4, 0)", ctx)).toEqual(num(3));
  });

  it("type 0 miss returns #N/A", () => {
    const ctx = makeContext({ ranges: { [ASC_RANGE]: ASC.map((r) => [...r]) } });
    const out = evalText("=MATCH(99, A1:A4, 0)", ctx);
    expect(isError(out) && out.v.kind).toBe("#N/A");
  });

  it("type 1 (default) returns largest ≤ lookup on sorted-asc data", () => {
    const ctx = makeContext({ ranges: { [ASC_RANGE]: ASC.map((r) => [...r]) } });
    expect(evalText("=MATCH(25, A1:A4)", ctx)).toEqual(num(2));
    expect(evalText("=MATCH(40, A1:A4)", ctx)).toEqual(num(4));
  });

  it("type -1 returns smallest ≥ lookup on sorted-desc data", () => {
    const ctx = makeContext({ ranges: { [DESC_RANGE]: DESC.map((r) => [...r]) } });
    expect(evalText("=MATCH(25, A1:A4, -1)", ctx)).toEqual(num(2));
  });
});

// ── XLOOKUP ───────────────────────────────────────────────────────────────

describe("XLOOKUP", () => {
  const ctx = makeContext({
    ranges: {
      "Sheet1!R0C0:R3C0": [[str("A")], [str("B")], [str("C")], [str("D")]],
      "Sheet1!R0C1:R3C1": [[num(10)], [num(20)], [num(30)], [num(40)]],
    },
  });

  it("returns the parallel value on exact match", () => {
    expect(evalText('=XLOOKUP("C", A1:A4, B1:B4)', ctx)).toEqual(num(30));
  });

  it("returns supplied if_not_found on miss", () => {
    expect(evalText('=XLOOKUP("Z", A1:A4, B1:B4, "missing")', ctx)).toEqual(str("missing"));
  });

  it("defaults if_not_found to #N/A", () => {
    const out = evalText('=XLOOKUP("Z", A1:A4, B1:B4)', ctx);
    expect(isError(out) && out.v.kind).toBe("#N/A");
  });

  it("match_mode=-1 finds next smaller item", () => {
    const numCtx = makeContext({
      ranges: {
        "Sheet1!R0C0:R3C0": [[num(10)], [num(20)], [num(30)], [num(40)]],
        "Sheet1!R0C1:R3C1": [[str("a")], [str("b")], [str("c")], [str("d")]],
      },
    });
    expect(evalText('=XLOOKUP(25, A1:A4, B1:B4, "x", -1)', numCtx)).toEqual(str("b"));
  });
});

// ── CHOOSE ────────────────────────────────────────────────────────────────

describe("CHOOSE", () => {
  const ctx = makeContext();

  it("returns the value at the 1-based index", () => {
    expect(evalText('=CHOOSE(1, "a", "b", "c")', ctx)).toEqual(str("a"));
    expect(evalText('=CHOOSE(3, "a", "b", "c")', ctx)).toEqual(str("c"));
  });

  it("index out of range returns #VALUE!", () => {
    const out = evalText('=CHOOSE(4, "a", "b", "c")', ctx);
    expect(isError(out) && out.v.kind).toBe("#VALUE!");
  });

  it("index < 1 returns #VALUE!", () => {
    const out = evalText('=CHOOSE(0, "a")', ctx);
    expect(isError(out) && out.v.kind).toBe("#VALUE!");
  });
});

// ── OFFSET ────────────────────────────────────────────────────────────────

describe("OFFSET", () => {
  const ctx = makeContext({
    cells: {
      "Sheet1!R0C0": num(1),
      "Sheet1!R1C0": num(2),
      "Sheet1!R2C0": num(3),
      "Sheet1!R0C1": num(10),
      "Sheet1!R1C1": num(20),
      "Sheet1!R2C1": num(30),
    },
  });

  it("offsets a single cell by (row, col)", () => {
    expect(evalText("=OFFSET(A1, 1, 1)", ctx)).toEqual(num(20));
    expect(evalText("=OFFSET(A1, 2, 0)", ctx)).toEqual(num(3));
  });

  it("returns a range when height/width > 1", () => {
    const out = evalText("=OFFSET(A1, 0, 0, 2, 2)", ctx);
    expect(out.kind).toBe("r");
    if (out.kind === "r") {
      expect(out.v).toEqual([
        [num(1), num(10)],
        [num(2), num(20)],
      ]);
    }
  });

  it("negative offset out of bounds returns #REF!", () => {
    const out = evalText("=OFFSET(A1, -1, 0)", ctx);
    expect(isError(out) && out.v.kind).toBe("#REF!");
  });

  it("zero or negative height returns #REF!", () => {
    const out = evalText("=OFFSET(A1, 0, 0, 0, 1)", ctx);
    expect(isError(out) && out.v.kind).toBe("#REF!");
  });
});

// ── INDIRECT ──────────────────────────────────────────────────────────────

describe("INDIRECT", () => {
  const ctx = makeContext({
    cells: { "Sheet1!R0C0": num(7), "Sheet1!R4C0": num(99) },
  });

  it("resolves an A1-style cell reference at eval time", () => {
    expect(evalText('=INDIRECT("A1")', ctx)).toEqual(num(7));
    expect(evalText('=INDIRECT("A5")', ctx)).toEqual(num(99));
  });

  it("supports range references", () => {
    const rangeCtx = makeContext({
      ranges: { "Sheet1!R0C0:R0C1": [[num(1), num(2)]] },
    });
    const out = evalText('=INDIRECT("A1:B1")', rangeCtx);
    expect(out.kind).toBe("r");
  });

  it("invalid ref string returns #REF!", () => {
    const out = evalText('=INDIRECT("not a ref")', ctx);
    expect(isError(out) && out.v.kind).toBe("#REF!");
  });

  it("a1=FALSE (R1C1) is unsupported in P0 → #REF!", () => {
    const out = evalText('=INDIRECT("R1C1", FALSE)', ctx);
    expect(isError(out) && out.v.kind).toBe("#REF!");
  });
});

// ── ROW / COLUMN ──────────────────────────────────────────────────────────

describe("ROW / COLUMN", () => {
  it("ROW() with no args returns caller's 1-based row", () => {
    const ctx = makeContext({ anchor: { sheet: SHEET, row: 4, col: 2, abs: AbsRef.NONE } });
    expect(evalText("=ROW()", ctx)).toEqual(num(5));
  });

  it("ROW(ref) returns ref's 1-based row", () => {
    const ctx = makeContext();
    expect(evalText("=ROW(C7)", ctx)).toEqual(num(7));
    expect(evalText("=ROW(B2:D9)", ctx)).toEqual(num(2));
  });

  it("COLUMN() with no args returns caller's 1-based column", () => {
    const ctx = makeContext({ anchor: { sheet: SHEET, row: 0, col: 5, abs: AbsRef.NONE } });
    expect(evalText("=COLUMN()", ctx)).toEqual(num(6));
  });

  it("COLUMN(ref) returns ref's 1-based column", () => {
    const ctx = makeContext();
    expect(evalText("=COLUMN(C7)", ctx)).toEqual(num(3));
  });
});

// ── ROWS / COLUMNS ────────────────────────────────────────────────────────

describe("ROWS / COLUMNS", () => {
  const ctx = makeContext();

  it("ROWS counts rows in a range", () => {
    expect(evalText("=ROWS(A1:A10)", ctx)).toEqual(num(10));
    expect(evalText("=ROWS(A1:C5)", ctx)).toEqual(num(5));
  });

  it("ROWS of a single cell is 1", () => {
    expect(evalText("=ROWS(B2)", ctx)).toEqual(num(1));
  });

  it("COLUMNS counts columns in a range", () => {
    expect(evalText("=COLUMNS(A1:E1)", ctx)).toEqual(num(5));
    expect(evalText("=COLUMNS(A1:C5)", ctx)).toEqual(num(3));
  });

  it("COLUMNS of a single cell is 1", () => {
    expect(evalText("=COLUMNS(B2)", ctx)).toEqual(num(1));
  });
});

// ── Volatility metadata ───────────────────────────────────────────────────

describe("registerLookup — registry metadata", () => {
  const reg = makeRegistry();

  it("declares OFFSET and INDIRECT as volatile", () => {
    const v = reg.volatileNames();
    expect(v.has("OFFSET")).toBe(true);
    expect(v.has("INDIRECT")).toBe(true);
  });

  it("registers all 12 lookup functions", () => {
    const names = [
      "VLOOKUP",
      "HLOOKUP",
      "INDEX",
      "MATCH",
      "XLOOKUP",
      "CHOOSE",
      "OFFSET",
      "INDIRECT",
      "ROW",
      "ROWS",
      "COLUMN",
      "COLUMNS",
    ];
    for (const n of names) expect(reg.has(n)).toBe(true);
  });

  it("OFFSET / ROW / etc declare lazyArgs", () => {
    expect(reg.get("OFFSET")?.lazyArgs).toBe(true);
    expect(reg.get("ROW")?.lazyArgs).toBe(true);
    expect(reg.get("COLUMN")?.lazyArgs).toBe(true);
    expect(reg.get("ROWS")?.lazyArgs).toBe(true);
    expect(reg.get("COLUMNS")?.lazyArgs).toBe(true);
  });
});
