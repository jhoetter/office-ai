import { describe, expect, it } from "vitest";
import {
  arity,
  createRegistry,
  type FunctionRegistry,
  type MutableFunctionRegistry,
} from "../function-registry.js";
import { createFormulaEngine, type EngineHost, type FormulaEngine } from "../recalc.js";
import { AbsRef, cellRefKey, makeCellKey, parseA1, type CellRef, type RangeRef } from "../references.js";
import { isError, num, type Range2D, type Value } from "../values.js";

const SHEET = "Sheet1";

function ref(a1: string): CellRef {
  const r = parseA1(a1, SHEET);
  if (!r) throw new Error(`bad a1 ${a1}`);
  return r;
}

function _rangeRef(sheet: string, r0: number, c0: number, r1: number, c1: number): RangeRef {
  return { sheet, r0, c0, r1, c1, abs0: AbsRef.NONE, abs1: AbsRef.NONE };
}

function makeRegistry(): MutableFunctionRegistry {
  const reg = createRegistry();
  reg.register({
    name: "SUM",
    arity: arity(1, 255),
    fn: (args) => {
      let total = 0;
      for (const a of args) {
        if (a.kind === "e") return a;
        if (a.kind === "r") {
          for (const row of a.v) {
            for (const cell of row) {
              if (cell.kind === "n") total += cell.v;
              else if (cell.kind === "b") total += cell.v ? 1 : 0;
              else if (cell.kind === "e") return cell;
            }
          }
          continue;
        }
        if (a.kind === "n") total += a.v;
        else if (a.kind === "b") total += a.v ? 1 : 0;
      }
      return num(total);
    },
  });
  reg.register({
    name: "RAND",
    arity: arity(0, 0),
    volatile: true,
    fn: (_args, ctx) => num(ctx.random()),
  });
  return reg;
}

class HostWorld implements EngineHost {
  // Raw values the user typed (separate from engine's formula cache).
  private raw = new Map<string, Value>();
  constructor(public engine: FormulaEngine | null = null) {}
  setRawCell(r: CellRef, v: Value): void {
    this.raw.set(cellRefKey(r), v);
  }
  readCell(ref: CellRef): Value {
    return this.raw.get(cellRefKey(ref)) ?? num(0);
  }
  readRange(ref: RangeRef): Range2D {
    const out: Value[][] = [];
    for (let r = ref.r0; r <= ref.r1; r++) {
      const row: Value[] = [];
      for (let c = ref.c0; c <= ref.c1; c++) {
        const cellRef: CellRef = { sheet: ref.sheet, row: r, col: c, abs: AbsRef.NONE };
        // Prefer engine cache (formula cells) over raw value (typed cells).
        const cached = this.engine?.getCachedValue(cellRef);
        row.push(cached ?? this.raw.get(cellRefKey(cellRef)) ?? num(0));
      }
      out.push(row);
    }
    return out;
  }
}

function buildEngine(reg: FunctionRegistry, host: HostWorld): FormulaEngine {
  const engine = createFormulaEngine({ registry: reg, host });
  host.engine = engine;
  return engine;
}

function setValue(engine: FormulaEngine, host: HostWorld, r: CellRef, value: Value): void {
  host.setRawCell(r, value);
  engine.addCell(r, null, value);
}

function setFormula(engine: FormulaEngine, r: CellRef, text: string): void {
  const formula = engine.parse(text, r);
  engine.addCell(r, formula, null);
}

describe("formula/dependency-graph + recalc — basics", () => {
  it("evaluates a simple chain A1 → B1 → C1 in topological order", () => {
    const reg = makeRegistry();
    const host = new HostWorld();
    const engine = buildEngine(reg, host);
    setValue(engine, host, ref("A1"), num(10));
    setFormula(engine, ref("B1"), "=A1+1");
    setFormula(engine, ref("C1"), "=B1*2");
    const result = engine.recalc();
    expect(result.values.get(makeCellKey(SHEET, 0, 1))).toEqual(num(11));
    expect(result.values.get(makeCellKey(SHEET, 0, 2))).toEqual(num(22));
  });

  it("after editing a leaf cell only the affected dependents recalc", () => {
    const reg = makeRegistry();
    const host = new HostWorld();
    const engine = buildEngine(reg, host);
    setValue(engine, host, ref("A1"), num(1));
    setValue(engine, host, ref("Z1"), num(99));
    setFormula(engine, ref("B1"), "=A1+1");
    setFormula(engine, ref("C1"), "=Z1+1"); // independent branch
    engine.recalc();

    // Now mutate A1 and recalc: only B1 should appear in `values`.
    setValue(engine, host, ref("A1"), num(5));
    engine.onCellChanged(ref("A1"));
    const second = engine.recalc();
    expect(second.values.get(makeCellKey(SHEET, 0, 1))).toEqual(num(6));
    expect(second.values.has(makeCellKey(SHEET, 0, 2))).toBe(false);
    expect(second.cycles.length).toBe(0);
  });

  it("removeCell drops edges and dirties downstream", () => {
    const reg = makeRegistry();
    const host = new HostWorld();
    const engine = buildEngine(reg, host);
    setValue(engine, host, ref("A1"), num(7));
    setFormula(engine, ref("B1"), "=A1+10");
    engine.recalc();
    expect(engine.getCachedValue(ref("B1"))).toEqual(num(17));

    engine.removeCell(ref("A1"));
    host.setRawCell(ref("A1"), num(0));
    const result = engine.recalc();
    expect(result.values.get(makeCellKey(SHEET, 0, 1))).toEqual(num(10));
  });
});

describe("formula/dependency-graph + recalc — range deps", () => {
  it("dirty inside a watched range bubbles to the SUM cell", () => {
    const reg = makeRegistry();
    const host = new HostWorld();
    const engine = buildEngine(reg, host);
    for (let i = 0; i < 3; i++) {
      setValue(engine, host, ref(`A${i + 1}`), num(i + 1));
    }
    setFormula(engine, ref("B1"), "=SUM(A1:A3)");
    const first = engine.recalc();
    expect(first.values.get(makeCellKey(SHEET, 0, 1))).toEqual(num(6));

    setValue(engine, host, ref("A2"), num(10));
    engine.onCellChanged(ref("A2"));
    const second = engine.recalc();
    expect(second.values.get(makeCellKey(SHEET, 0, 1))).toEqual(num(14));
  });

  it("a cell outside the range does not dirty the SUM cell", () => {
    const reg = makeRegistry();
    const host = new HostWorld();
    const engine = buildEngine(reg, host);
    for (let i = 0; i < 3; i++) {
      setValue(engine, host, ref(`A${i + 1}`), num(i + 1));
    }
    setValue(engine, host, ref("D1"), num(999));
    setFormula(engine, ref("B1"), "=SUM(A1:A3)");
    engine.recalc();

    setValue(engine, host, ref("D1"), num(0));
    engine.onCellChanged(ref("D1"));
    const result = engine.recalc();
    expect(result.values.has(makeCellKey(SHEET, 0, 1))).toBe(false);
  });
});

describe("formula/dependency-graph — cycle detection", () => {
  it("detects a 2-cell cycle and returns #REF! with cycle metadata", () => {
    const reg = makeRegistry();
    const host = new HostWorld();
    const engine = buildEngine(reg, host);
    setFormula(engine, ref("A1"), "=B1+1");
    setFormula(engine, ref("B1"), "=A1+1");
    const result = engine.recalc();
    expect(result.cycles.length).toBe(1);
    const a = result.values.get(makeCellKey(SHEET, 0, 0));
    const b = result.values.get(makeCellKey(SHEET, 0, 1));
    expect(a && isError(a)).toBe(true);
    expect(b && isError(b)).toBe(true);
    if (a && a.kind === "e") expect(a.v.kind).toBe("#REF!");
  });

  it("detects a 3-cell cycle", () => {
    const reg = makeRegistry();
    const host = new HostWorld();
    const engine = buildEngine(reg, host);
    setFormula(engine, ref("A1"), "=B1+1");
    setFormula(engine, ref("B1"), "=C1+1");
    setFormula(engine, ref("C1"), "=A1+1");
    const result = engine.recalc();
    expect(result.cycles.length).toBe(1);
    expect(result.cycles[0].length).toBe(3);
  });

  it("a self-referential formula is reported as a cycle", () => {
    const reg = makeRegistry();
    const host = new HostWorld();
    const engine = buildEngine(reg, host);
    setFormula(engine, ref("A1"), "=A1+1");
    const result = engine.recalc();
    expect(result.cycles.length).toBe(1);
  });
});

describe("formula/dependency-graph — volatile", () => {
  it("volatile cells force-dirty every recalc", () => {
    const reg = makeRegistry();
    const host = new HostWorld();
    const engine = buildEngine(reg, host);
    setFormula(engine, ref("A1"), "=RAND()");
    const first = engine.recalc();
    expect(first.values.has(makeCellKey(SHEET, 0, 0))).toBe(true);

    // Drain again with no edits — RAND should re-fire.
    const second = engine.recalc();
    expect(second.values.has(makeCellKey(SHEET, 0, 0))).toBe(true);
  });
});

describe("formula/dependency-graph — recalcAll", () => {
  it("re-evaluates every known cell", () => {
    const reg = makeRegistry();
    const host = new HostWorld();
    const engine = buildEngine(reg, host);
    setValue(engine, host, ref("A1"), num(2));
    setFormula(engine, ref("B1"), "=A1*A1");
    setFormula(engine, ref("C1"), "=B1+1");
    engine.recalc();

    // Reach into the host: pretend A1 magically changed but no
    // onCellChanged was called. recalcAll() should still pick it up.
    host.setRawCell(ref("A1"), num(10));
    engine.addCell(ref("A1"), null, num(10)); // reflect into graph
    const result = engine.recalcAll();
    expect(result.values.get(makeCellKey(SHEET, 0, 1))).toEqual(num(100));
    expect(result.values.get(makeCellKey(SHEET, 0, 2))).toEqual(num(101));
  });
});

describe("formula/dependency-graph — perf budget (smoke)", () => {
  it("recalc on a 1k-cell linear chain after a single edit completes well under the §17 budget", () => {
    const reg = makeRegistry();
    const host = new HostWorld();
    const engine = buildEngine(reg, host);
    const N = 1000;
    setValue(engine, host, ref("A1"), num(1));
    for (let i = 2; i <= N; i++) {
      setFormula(engine, ref(`A${i}`), `=A${i - 1}+1`);
    }
    engine.recalc(); // initial calc
    setValue(engine, host, ref("A1"), num(100));
    engine.onCellChanged(ref("A1"));
    const result = engine.recalc();
    // A1 → A1000 cascade, single dirty edit dirties N-1 dependents.
    expect(result.values.size).toBeGreaterThanOrEqual(N - 1);
    expect(result.cycles.length).toBe(0);
    // The §17 budget is 100ms for 10k formulas; 1k should fit well
    // under the same ceiling. We use 100ms here (rather than a tight
    // 50ms) so the smoke is robust under noisy CI runners while
    // still catching any real order-of-magnitude regression.
    expect(result.elapsedMs).toBeLessThan(100);
    // Last cell value: 100 + (N-1)
    expect(result.values.get(makeCellKey(SHEET, N - 1, 0))).toEqual(num(100 + N - 1));
  });
});
