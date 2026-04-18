import { describe, expect, it } from "vitest";
import { evaluate } from "../evaluator.js";
import {
  arity,
  createRegistry,
  type EvalContext,
  type FunctionRegistry,
  type MutableFunctionRegistry,
} from "../function-registry.js";
import { parse } from "../parser.js";
import { AbsRef, type CellRef, type RangeRef } from "../references.js";
import { bool, err, isError, num, rangeVal, str, toNumber, type Value } from "../values.js";

const ANCHOR: CellRef = { sheet: "Sheet1", row: 0, col: 0, abs: AbsRef.NONE };

function makeContext(opts: {
  cells?: Record<string, Value>;
  ranges?: Record<string, Value[][]>;
  names?: Record<string, RangeRef | CellRef>;
  registry?: FunctionRegistry;
  now?: number;
  random?: number;
}): EvalContext {
  const cells = opts.cells ?? {};
  const ranges = opts.ranges ?? {};
  const names = opts.names ?? {};
  const registry = opts.registry ?? createRegistry();
  const now = opts.now ?? 0;
  const random = opts.random ?? 0;
  return {
    getCell(ref: CellRef): Value {
      const k = `${ref.sheet}!R${ref.row}C${ref.col}`;
      return cells[k] ?? num(0);
    },
    getRange(ref: RangeRef): readonly (readonly Value[])[] {
      const k = `${ref.sheet}!R${ref.r0}C${ref.c0}:R${ref.r1}C${ref.c1}`;
      return ranges[k] ?? [[num(0)]];
    },
    resolveName(name: string) {
      return names[name] ?? names[name.toUpperCase()];
    },
    now: () => now,
    random: () => random,
    anchor: ANCHOR,
    registry,
  };
}

function evalText(text: string, ctx?: EvalContext): Value {
  const ast = parse(text, { anchor: ANCHOR }).ast;
  return evaluate(ast, ctx ?? makeContext({}));
}

describe("formula/evaluator — literals and operators", () => {
  it("evaluates numeric literals", () => {
    expect(evalText("=42")).toEqual(num(42));
  });

  it("evaluates basic arithmetic", () => {
    expect(evalText("=1+2*3")).toEqual(num(7));
    expect(evalText("=(1+2)*3")).toEqual(num(9));
    expect(evalText("=10/4")).toEqual(num(2.5));
    expect(evalText("=2^10")).toEqual(num(1024));
  });

  it("encodes the Excel unary-minus quirk: -2^2 == 4", () => {
    expect(evalText("=-2^2")).toEqual(num(4));
  });

  it("evaluates string concatenation with coercion", () => {
    expect(evalText('="hi"&" "&"there"')).toEqual(str("hi there"));
    expect(evalText('=1&"x"')).toEqual(str("1x"));
  });

  it("evaluates comparisons returning booleans", () => {
    expect(evalText("=1<2")).toEqual(bool(true));
    expect(evalText("=1=1")).toEqual(bool(true));
    expect(evalText("=1<>2")).toEqual(bool(true));
    expect(evalText('="a"="A"')).toEqual(bool(true)); // case-insensitive
  });

  it("propagates errors through arithmetic", () => {
    const out = evalText("=1/0+5");
    expect(isError(out)).toBe(true);
  });

  it("evaluates percent as a postfix /100", () => {
    expect(evalText("=50%")).toEqual(num(0.5));
    expect(evalText("=200%+0.5")).toEqual(num(2.5));
  });
});

describe("formula/evaluator — references", () => {
  it("reads a cell value via getCell", () => {
    const ctx = makeContext({
      cells: { "Sheet1!R0C0": num(7), "Sheet1!R1C0": num(35) },
    });
    expect(evalText("=A1+A2", ctx)).toEqual(num(42));
  });

  it("reads a range via getRange and surfaces it as a RangeValue", () => {
    const ctx = makeContext({
      ranges: { "Sheet1!R0C0:R2C0": [[num(1)], [num(2)], [num(3)]] },
    });
    const out = evalText("=A1:A3", ctx);
    expect(out.kind).toBe("r");
  });

  it("returns #NAME? for unresolved defined names", () => {
    const out = evalText("=Foo+1", makeContext({}));
    expect(isError(out)).toBe(true);
  });

  it("resolves defined names via context", () => {
    const ctx = makeContext({
      cells: { "Sheet1!R8C25": num(99) },
      names: { Anchor: { sheet: "Sheet1", row: 8, col: 25, abs: AbsRef.NONE } },
    });
    expect(evalText("=Anchor+1", ctx)).toEqual(num(100));
  });
});

describe("formula/evaluator — function dispatch", () => {
  function withSum(): MutableFunctionRegistry {
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
          const n = toNumber(a);
          if (isError(n)) return n;
          total += n.v;
        }
        return num(total);
      },
    });
    return reg;
  }

  it("dispatches to a registered eager function", () => {
    const reg = withSum();
    const ctx = makeContext({ registry: reg });
    expect(evalText("=SUM(1,2,3,4)", ctx)).toEqual(num(10));
  });

  it("upper-cases function names at lookup time", () => {
    const reg = withSum();
    const ctx = makeContext({ registry: reg });
    expect(evalText("=sum(1,2,3)", ctx)).toEqual(num(6));
  });

  it("returns #NAME? for unknown functions", () => {
    const ctx = makeContext({});
    const out = evalText("=NOPE(1,2)", ctx);
    expect(isError(out)).toBe(true);
  });

  it("returns #N/A when arity does not match", () => {
    const reg = createRegistry();
    reg.register({
      name: "PAIR",
      arity: arity(2, 2),
      fn: (args) => num((args[0] as { v: number }).v + (args[1] as { v: number }).v),
    });
    const ctx = makeContext({ registry: reg });
    const out = evalText("=PAIR(1)", ctx);
    expect(isError(out)).toBe(true);
  });

  it("flattens range arguments inside SUM", () => {
    const reg = withSum();
    const ctx = makeContext({
      registry: reg,
      ranges: {
        "Sheet1!R0C0:R2C0": [[num(10)], [num(20)], [num(30)]],
      },
    });
    expect(evalText("=SUM(A1:A3)", ctx)).toEqual(num(60));
  });
});

describe("formula/evaluator — lazy-arg dispatch", () => {
  function withIf(): MutableFunctionRegistry {
    const reg = createRegistry();
    reg.register({
      name: "IF",
      arity: arity(2, 3),
      lazyArgs: true,
      fn: (args, lazy) => {
        const cond = lazy.evaluate(args[0]);
        if (cond.kind === "e") return cond;
        // Excel coerces to boolean: 0 → FALSE, anything else → TRUE.
        const truthy = cond.kind === "b" ? cond.v : cond.kind === "n" ? cond.v !== 0 : true;
        if (truthy) return lazy.evaluate(args[1]);
        if (args.length === 3) return lazy.evaluate(args[2]);
        return bool(false);
      },
    });
    return reg;
  }

  it("evaluates the chosen branch only", () => {
    const reg = withIf();
    let leftCalls = 0;
    let rightCalls = 0;
    reg.register({
      name: "LCOUNT",
      arity: arity(1, 1),
      fn: (args) => {
        leftCalls++;
        return args[0];
      },
    });
    reg.register({
      name: "RCOUNT",
      arity: arity(1, 1),
      fn: (args) => {
        rightCalls++;
        return args[0];
      },
    });
    const ctx = makeContext({ registry: reg });
    expect(evalText("=IF(TRUE, LCOUNT(1), RCOUNT(2))", ctx)).toEqual(num(1));
    expect(leftCalls).toBe(1);
    expect(rightCalls).toBe(0);
  });

  it("omitting third arg defaults to FALSE", () => {
    const reg = withIf();
    const ctx = makeContext({ registry: reg });
    expect(evalText("=IF(FALSE, 1)", ctx)).toEqual(bool(false));
  });

  it("propagates an error condition", () => {
    const reg = withIf();
    const ctx = makeContext({ registry: reg });
    const out = evalText("=IF(1/0, 1, 2)", ctx);
    expect(isError(out)).toBe(true);
  });
});

describe("formula/evaluator — array literals", () => {
  it("evaluates {1,2;3,4} as a 2×2 range", () => {
    const out = evalText("={1,2;3,4}");
    expect(out.kind).toBe("r");
    if (out.kind === "r") {
      expect(out.v).toEqual([
        [num(1), num(2)],
        [num(3), num(4)],
      ]);
    }
  });
});

describe("formula/evaluator — context plumbing", () => {
  it("exposes ctx.now / ctx.random / ctx.anchor to functions", () => {
    const reg = createRegistry();
    reg.register({
      name: "PROBE",
      arity: arity(0, 0),
      volatile: true,
      fn: (_args, ctx) =>
        rangeVal([[num(ctx.now()), num(ctx.random()), num(ctx.anchor.row), num(ctx.anchor.col)]]),
    });
    const ctx = makeContext({ registry: reg, now: 1234, random: 0.5 });
    const out = evalText("=PROBE()", ctx);
    expect(out.kind).toBe("r");
    if (out.kind === "r") {
      expect(out.v[0].map((v) => (v as { v: number }).v)).toEqual([1234, 0.5, 0, 0]);
    }
  });

  it("never throws on unrecognised input — returns an error value instead", () => {
    const out = evalText("=UNKNOWNFN(1,2,3)");
    expect(isError(out)).toBe(true);
  });

  it("registry tracks volatile function names", () => {
    const reg = createRegistry();
    reg.register({ name: "RAND", arity: arity(0, 0), volatile: true, fn: () => num(0.5) });
    reg.register({ name: "SUM", arity: arity(1, 255), fn: () => num(0) });
    expect(reg.volatileNames().has("RAND")).toBe(true);
    expect(reg.volatileNames().has("SUM")).toBe(false);
  });
});

describe("formula/evaluator — wires through the parser", () => {
  it("end-to-end: parse + evaluate produces the same result as direct construction", () => {
    const reg = createRegistry();
    reg.register({
      name: "DOUBLE",
      arity: arity(1, 1),
      fn: (args) => {
        const n = toNumber(args[0]);
        if (isError(n)) return n;
        return num(n.v * 2);
      },
    });
    const ctx = makeContext({ registry: reg });
    expect(evalText("=DOUBLE(3)+1", ctx)).toEqual(num(7));
  });

  it("propagates an explicit error literal through arithmetic", () => {
    const out = evalText("=#REF!+1");
    expect(out.kind).toBe("e");
    if (out.kind === "e") {
      expect(out.v.kind).toBe("#REF!");
    }
  });

  it("eq() catches the err helper too", () => {
    const out = evalText("=#N/A=#N/A");
    // Comparing two errors short-circuits with the first error
    // (matches the §7.4 propagation rule).
    expect(out.kind).toBe("e");
  });

  it("err() helper round-trips a kind into a runtime ErrorValue", () => {
    expect(err("#NAME?")).toMatchObject({ kind: "e", v: { kind: "#NAME?" } });
  });
});
