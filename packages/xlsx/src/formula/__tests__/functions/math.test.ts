import { describe, expect, it } from "vitest";
import { evaluate } from "../../evaluator.js";
import {
  createRegistry,
  type EvalContext,
  type FunctionRegistry,
  type MutableFunctionRegistry,
} from "../../function-registry.js";
import { registerMath } from "../../functions/math.js";
import { parse } from "../../parser.js";
import { AbsRef, type CellRef, type RangeRef } from "../../references.js";
import { bool, err, isError, num, rangeVal, str, type Range2D, type Value } from "../../values.js";

const ANCHOR: CellRef = { sheet: "S", row: 0, col: 0, abs: AbsRef.NONE };

function makeRegistry(): MutableFunctionRegistry {
  const reg = createRegistry();
  registerMath(reg);
  return reg;
}

interface CtxOpts {
  cells?: Record<string, Value>;
  ranges?: Record<string, Range2D>;
  registry?: FunctionRegistry;
  random?: number;
}

function makeCtx(opts: CtxOpts = {}): EvalContext {
  const reg = opts.registry ?? makeRegistry();
  const cells = opts.cells ?? {};
  const ranges = opts.ranges ?? {};
  const random = opts.random ?? 0;
  return {
    getCell(ref: CellRef): Value {
      const k = `${ref.sheet}!R${ref.row}C${ref.col}`;
      return cells[k] ?? num(0);
    },
    getRange(ref: RangeRef): Range2D {
      const k = `${ref.sheet}!R${ref.r0}C${ref.c0}:R${ref.r1}C${ref.c1}`;
      return ranges[k] ?? [[num(0)]];
    },
    resolveName: () => undefined,
    now: () => 0,
    random: () => random,
    anchor: ANCHOR,
    registry: reg,
  };
}

function ev(text: string, ctx?: EvalContext): Value {
  const c = ctx ?? makeCtx();
  return evaluate(parse(text, { anchor: ANCHOR }).ast, c);
}

function call(name: string, args: ReadonlyArray<Value>, ctx?: EvalContext): Value {
  const c = ctx ?? makeCtx();
  const entry = c.registry.get(name);
  if (!entry || entry.lazyArgs) throw new Error(`no eager fn ${name}`);
  return entry.fn(args, c);
}

function rng(rows: Value[][]): Value {
  return rangeVal(rows);
}

function expectClose(actual: Value, expected: number, tol = 1e-9): void {
  expect(actual.kind).toBe("n");
  if (actual.kind === "n") expect(Math.abs(actual.v - expected)).toBeLessThan(tol);
}

// ── Aggregations ──────────────────────────────────────────────────────────

describe("MATH/SUM", () => {
  it("sums scalar args", () => {
    expect(ev("=SUM(1,2,3,4)")).toEqual(num(10));
  });
  it("sums a range and silently skips text cells", () => {
    expect(call("SUM", [rng([[num(1), str("oops"), num(2)]])])).toEqual(num(3));
  });
  it("coerces a literal text scalar to number", () => {
    expect(call("SUM", [str("5"), num(10)])).toEqual(num(15));
  });
  it("returns #VALUE! on uncoercible scalar text", () => {
    const out = call("SUM", [str("abc"), num(1)]);
    expect(isError(out)).toBe(true);
  });
  it("propagates errors inside a range", () => {
    const out = call("SUM", [rng([[num(1), err("#REF!")]])]);
    expect(isError(out)).toBe(true);
  });
});

describe("MATH/AVERAGE", () => {
  it("averages scalar args", () => {
    expect(ev("=AVERAGE(1,2,3,4)")).toEqual(num(2.5));
  });
  it("averages a range, skipping text", () => {
    expect(call("AVERAGE", [rng([[num(2), str("x"), num(4)]])])).toEqual(num(3));
  });
  it("returns #DIV/0! when no numeric values", () => {
    const out = call("AVERAGE", [rng([[str("a"), str("b")]])]);
    expect(out.kind).toBe("e");
    if (out.kind === "e") expect(out.v.kind).toBe("#DIV/0!");
  });
});

describe("MATH/COUNT", () => {
  it("counts numeric scalar args", () => {
    expect(ev("=COUNT(1,2,3)")).toEqual(num(3));
  });
  it("ignores text cells inside a range but counts numbers", () => {
    expect(call("COUNT", [rng([[num(1), str("x"), num(2), str("y")]])])).toEqual(num(2));
  });
  it("counts a coercible scalar string as numeric", () => {
    expect(call("COUNT", [str("5"), str("nope")])).toEqual(num(1));
  });
});

describe("MATH/COUNTA", () => {
  it("counts every non-empty cell in a range", () => {
    expect(call("COUNTA", [rng([[num(1), str("x"), bool(true), str("")]])])).toEqual(num(3));
  });
  it("counts each scalar arg", () => {
    expect(call("COUNTA", [num(1), str("a"), bool(false)])).toEqual(num(3));
  });
  it("propagates a range error", () => {
    const out = call("COUNTA", [err("#NAME?")]);
    expect(isError(out)).toBe(true);
  });
});

describe("MATH/COUNTBLANK", () => {
  it("counts empty-string cells", () => {
    expect(call("COUNTBLANK", [rng([[str(""), num(1), str(""), str("a")]])])).toEqual(num(2));
  });
  it("returns 0 for a non-blank scalar", () => {
    expect(call("COUNTBLANK", [num(0)])).toEqual(num(0));
  });
  it("propagates errors", () => {
    const out = call("COUNTBLANK", [err("#REF!")]);
    expect(isError(out)).toBe(true);
  });
});

describe("MATH/MIN", () => {
  it("returns the smallest scalar", () => {
    expect(ev("=MIN(3,1,2)")).toEqual(num(1));
  });
  it("returns 0 when no numeric values found", () => {
    expect(call("MIN", [rng([[str("a"), str("b")]])])).toEqual(num(0));
  });
  it("propagates errors", () => {
    const out = call("MIN", [rng([[err("#NUM!"), num(1)]])]);
    expect(isError(out)).toBe(true);
  });
});

describe("MATH/MAX", () => {
  it("returns the largest scalar", () => {
    expect(ev("=MAX(3,1,2)")).toEqual(num(3));
  });
  it("returns 0 when no numeric values found", () => {
    expect(call("MAX", [rng([[str("a")]])])).toEqual(num(0));
  });
  it("propagates errors", () => {
    const out = call("MAX", [err("#VALUE!")]);
    expect(isError(out)).toBe(true);
  });
});

// ── *IF / *IFS family ────────────────────────────────────────────────────

describe("MATH/SUMIF", () => {
  it("sums values matching a numeric criterion", () => {
    const r = rng([[num(1), num(5), num(10), num(20)]]);
    expect(call("SUMIF", [r, str(">5")])).toEqual(num(30));
  });
  it("uses a separate sum_range when supplied", () => {
    const range = rng([[str("a"), str("b"), str("a")]]);
    const sums = rng([[num(1), num(2), num(3)]]);
    expect(call("SUMIF", [range, str("a"), sums])).toEqual(num(4));
  });
  it("supports wildcards", () => {
    const range = rng([[str("apple"), str("apricot"), str("banana")]]);
    const sums = rng([[num(1), num(2), num(4)]]);
    expect(call("SUMIF", [range, str("ap*"), sums])).toEqual(num(3));
  });
});

describe("MATH/SUMIFS", () => {
  it("sums when all criteria match", () => {
    const sums = rng([[num(10), num(20), num(30), num(40)]]);
    const r1 = rng([[str("a"), str("a"), str("b"), str("a")]]);
    const r2 = rng([[num(1), num(2), num(2), num(3)]]);
    expect(call("SUMIFS", [sums, r1, str("a"), r2, str(">1")])).toEqual(num(60));
  });
  it("returns 0 when no rows match", () => {
    const sums = rng([[num(1), num(2)]]);
    const r1 = rng([[str("x"), str("y")]]);
    expect(call("SUMIFS", [sums, r1, str("z")])).toEqual(num(0));
  });
  it("returns #VALUE! on shape mismatch", () => {
    const sums = rng([[num(1), num(2)]]);
    const r1 = rng([[str("a"), str("b"), str("c")]]);
    const out = call("SUMIFS", [sums, r1, str("a")]);
    expect(out.kind).toBe("e");
    if (out.kind === "e") expect(out.v.kind).toBe("#VALUE!");
  });
});

describe("MATH/COUNTIF", () => {
  it("counts cells matching a numeric criterion", () => {
    expect(call("COUNTIF", [rng([[num(1), num(5), num(10)]]), str(">=5")])).toEqual(num(2));
  });
  it("treats a number criterion as equality", () => {
    expect(call("COUNTIF", [rng([[num(1), num(2), num(2)]]), num(2)])).toEqual(num(2));
  });
  it("supports the ? wildcard", () => {
    const r = rng([[str("ab"), str("ac"), str("abc")]]);
    expect(call("COUNTIF", [r, str("a?")])).toEqual(num(2));
  });
});

describe("MATH/COUNTIFS", () => {
  it("counts only rows matching every criterion", () => {
    const r1 = rng([[str("a"), str("a"), str("b")]]);
    const r2 = rng([[num(1), num(2), num(2)]]);
    expect(call("COUNTIFS", [r1, str("a"), r2, str(">1")])).toEqual(num(1));
  });
  it("returns #VALUE! on odd argument count", () => {
    const out = call("COUNTIFS", [rng([[num(1)]]), str(">0"), rng([[num(2)]])]);
    expect(out.kind).toBe("e");
  });
});

describe("MATH/AVERAGEIF", () => {
  it("averages matching values", () => {
    expect(call("AVERAGEIF", [rng([[num(1), num(2), num(3), num(4)]]), str(">2")])).toEqual(num(3.5));
  });
  it("uses a separate average_range", () => {
    const range = rng([[str("a"), str("b"), str("a")]]);
    const avgs = rng([[num(2), num(99), num(4)]]);
    expect(call("AVERAGEIF", [range, str("a"), avgs])).toEqual(num(3));
  });
  it("returns #DIV/0! when no match", () => {
    const out = call("AVERAGEIF", [rng([[num(1), num(2)]]), str(">100")]);
    expect(out.kind).toBe("e");
    if (out.kind === "e") expect(out.v.kind).toBe("#DIV/0!");
  });
});

describe("MATH/AVERAGEIFS", () => {
  it("averages rows matching every criterion", () => {
    const avgs = rng([[num(10), num(20), num(30), num(40)]]);
    const r1 = rng([[str("a"), str("a"), str("b"), str("a")]]);
    const r2 = rng([[num(1), num(2), num(2), num(3)]]);
    expect(call("AVERAGEIFS", [avgs, r1, str("a"), r2, str(">1")])).toEqual(num(30));
  });
  it("returns #DIV/0! when no rows match", () => {
    const avgs = rng([[num(1), num(2)]]);
    const r1 = rng([[str("x"), str("y")]]);
    const out = call("AVERAGEIFS", [avgs, r1, str("z")]);
    expect(out.kind).toBe("e");
    if (out.kind === "e") expect(out.v.kind).toBe("#DIV/0!");
  });
  it("propagates an error inside the average range", () => {
    const avgs = rng([[err("#REF!"), num(2)]]);
    const r1 = rng([[str("a"), str("a")]]);
    const out = call("AVERAGEIFS", [avgs, r1, str("a")]);
    expect(isError(out)).toBe(true);
  });
});

// ── Rounding ──────────────────────────────────────────────────────────────

describe("MATH/ROUND", () => {
  it("rounds to the requested number of decimals", () => {
    expect(ev("=ROUND(1.234, 2)")).toEqual(num(1.23));
  });
  it("rounds half away from zero (negative)", () => {
    expect(ev("=ROUND(-2.5, 0)")).toEqual(num(-3));
  });
  it("returns #VALUE! on uncoercible text", () => {
    const out = call("ROUND", [str("hi"), num(0)]);
    expect(isError(out)).toBe(true);
  });
});

describe("MATH/ROUNDUP", () => {
  it("rounds away from zero up", () => {
    expect(ev("=ROUNDUP(1.001, 2)")).toEqual(num(1.01));
  });
  it("rounds negatives away from zero", () => {
    expect(ev("=ROUNDUP(-1.001, 2)")).toEqual(num(-1.01));
  });
  it("rounds to integer when digits=0", () => {
    expect(ev("=ROUNDUP(1.1, 0)")).toEqual(num(2));
  });
});

describe("MATH/ROUNDDOWN", () => {
  it("rounds toward zero", () => {
    expect(ev("=ROUNDDOWN(1.999, 2)")).toEqual(num(1.99));
  });
  it("rounds negatives toward zero", () => {
    expect(ev("=ROUNDDOWN(-1.999, 2)")).toEqual(num(-1.99));
  });
  it("returns 0 for tiny positives at digits=0", () => {
    expect(ev("=ROUNDDOWN(0.99, 0)")).toEqual(num(0));
  });
});

describe("MATH/INT", () => {
  it("floors a positive number", () => {
    expect(ev("=INT(3.7)")).toEqual(num(3));
  });
  it("floors a negative number toward -inf", () => {
    expect(ev("=INT(-3.2)")).toEqual(num(-4));
  });
  it("returns #VALUE! on uncoercible text", () => {
    const out = call("INT", [str("hi")]);
    expect(isError(out)).toBe(true);
  });
});

describe("MATH/ABS", () => {
  it("returns absolute value", () => {
    expect(ev("=ABS(-7)")).toEqual(num(7));
  });
  it("coerces a numeric string", () => {
    expect(call("ABS", [str("-3.5")])).toEqual(num(3.5));
  });
  it("propagates errors", () => {
    expect(call("ABS", [err("#REF!")])).toEqual(err("#REF!"));
  });
});

describe("MATH/MOD", () => {
  it("returns the remainder", () => {
    expect(ev("=MOD(10, 3)")).toEqual(num(1));
  });
  it("result has same sign as the divisor", () => {
    expect(ev("=MOD(-3, 2)")).toEqual(num(1));
  });
  it("returns #DIV/0! when divisor is zero", () => {
    const out = ev("=MOD(5, 0)");
    expect(out.kind).toBe("e");
    if (out.kind === "e") expect(out.v.kind).toBe("#DIV/0!");
  });
});

describe("MATH/POWER", () => {
  it("computes base^exp", () => {
    expect(ev("=POWER(2, 10)")).toEqual(num(1024));
  });
  it("supports fractional exponents", () => {
    expectClose(ev("=POWER(9, 0.5)"), 3);
  });
  it("returns #NUM! on undefined results", () => {
    const out = ev("=POWER(-1, 0.5)");
    expect(out.kind).toBe("e");
    if (out.kind === "e") expect(out.v.kind).toBe("#NUM!");
  });
});

describe("MATH/SQRT", () => {
  it("computes sqrt", () => {
    expect(ev("=SQRT(16)")).toEqual(num(4));
  });
  it("returns #NUM! on negative input", () => {
    const out = ev("=SQRT(-1)");
    expect(out.kind).toBe("e");
    if (out.kind === "e") expect(out.v.kind).toBe("#NUM!");
  });
  it("coerces a numeric string", () => {
    expect(call("SQRT", [str("25")])).toEqual(num(5));
  });
});

describe("MATH/CEILING", () => {
  it("rounds up to the nearest integer by default", () => {
    expect(ev("=CEILING(2.1)")).toEqual(num(3));
  });
  it("rounds away from zero with significance", () => {
    expect(ev("=CEILING(-2.1, 1)")).toEqual(num(-3));
  });
  it("rounds to a non-1 significance", () => {
    expect(ev("=CEILING(7, 5)")).toEqual(num(10));
  });
});

describe("MATH/FLOOR", () => {
  it("floors to the nearest integer by default", () => {
    expect(ev("=FLOOR(2.9)")).toEqual(num(2));
  });
  it("floors toward zero with significance", () => {
    expect(ev("=FLOOR(-2.9, 1)")).toEqual(num(-2));
  });
  it("returns #DIV/0! when significance is zero", () => {
    const out = ev("=FLOOR(5, 0)");
    expect(out.kind).toBe("e");
    if (out.kind === "e") expect(out.v.kind).toBe("#DIV/0!");
  });
});

// ── Random ────────────────────────────────────────────────────────────────

describe("MATH/RAND", () => {
  it("returns ctx.random()", () => {
    const ctx = makeCtx({ random: 0.42 });
    expect(ev("=RAND()", ctx)).toEqual(num(0.42));
  });
  it("is registered as volatile", () => {
    const reg = makeRegistry();
    expect(reg.volatileNames().has("RAND")).toBe(true);
  });
  it("ignores arguments (arity 0)", () => {
    const out = ev("=RAND(1)");
    expect(isError(out)).toBe(true);
  });
});

describe("MATH/RANDBETWEEN", () => {
  it("returns an integer in the inclusive range", () => {
    const ctx = makeCtx({ random: 0 });
    expect(ev("=RANDBETWEEN(1, 10)", ctx)).toEqual(num(1));
  });
  it("returns the upper bound when random is near 1", () => {
    const ctx = makeCtx({ random: 0.9999 });
    expect(ev("=RANDBETWEEN(1, 10)", ctx)).toEqual(num(10));
  });
  it("returns #NUM! when low > high", () => {
    const out = ev("=RANDBETWEEN(10, 1)");
    expect(out.kind).toBe("e");
    if (out.kind === "e") expect(out.v.kind).toBe("#NUM!");
  });
});

// ── Order statistics ─────────────────────────────────────────────────────

describe("MATH/LARGE", () => {
  it("returns the kth largest value", () => {
    expect(call("LARGE", [rng([[num(3), num(1), num(4), num(1), num(5)]]), num(2)])).toEqual(num(4));
  });
  it("returns #NUM! when k is out of bounds", () => {
    const out = call("LARGE", [rng([[num(1), num(2)]]), num(5)]);
    expect(out.kind).toBe("e");
    if (out.kind === "e") expect(out.v.kind).toBe("#NUM!");
  });
  it("returns #NUM! when k < 1", () => {
    const out = call("LARGE", [rng([[num(1), num(2)]]), num(0)]);
    expect(out.kind).toBe("e");
    if (out.kind === "e") expect(out.v.kind).toBe("#NUM!");
  });
});

describe("MATH/SMALL", () => {
  it("returns the kth smallest value", () => {
    expect(call("SMALL", [rng([[num(3), num(1), num(4), num(1), num(5)]]), num(2)])).toEqual(num(1));
  });
  it("returns #NUM! when k is out of bounds", () => {
    const out = call("SMALL", [rng([[num(1)]]), num(2)]);
    expect(isError(out)).toBe(true);
  });
  it("propagates errors in the source range", () => {
    const out = call("SMALL", [rng([[num(1), err("#REF!")]]), num(1)]);
    expect(isError(out)).toBe(true);
  });
});

describe("MATH/RANK", () => {
  it("ranks descending by default", () => {
    expect(call("RANK", [num(5), rng([[num(1), num(3), num(5), num(7)]])])).toEqual(num(2));
  });
  it("ranks ascending when order is non-zero", () => {
    expect(call("RANK", [num(5), rng([[num(1), num(3), num(5), num(7)]]), num(1)])).toEqual(num(3));
  });
  it("returns #N/A when value is not in range", () => {
    const out = call("RANK", [num(99), rng([[num(1), num(2)]])]);
    expect(out.kind).toBe("e");
    if (out.kind === "e") expect(out.v.kind).toBe("#N/A");
  });
});

describe("MATH/MEDIAN", () => {
  it("returns the middle value with odd count", () => {
    expect(ev("=MEDIAN(1,2,3,4,5)")).toEqual(num(3));
  });
  it("averages the two middles with even count", () => {
    expect(ev("=MEDIAN(1,2,3,4)")).toEqual(num(2.5));
  });
  it("returns #NUM! when no numeric values", () => {
    const out = call("MEDIAN", [rng([[str("a")]])]);
    expect(out.kind).toBe("e");
    if (out.kind === "e") expect(out.v.kind).toBe("#NUM!");
  });
});

describe("MATH/STDEV", () => {
  it("computes sample standard deviation", () => {
    expectClose(ev("=STDEV(2,4,4,4,5,5,7,9)"), Math.sqrt(32 / 7));
  });
  it("returns #DIV/0! with fewer than 2 values", () => {
    const out = ev("=STDEV(5)");
    expect(out.kind).toBe("e");
    if (out.kind === "e") expect(out.v.kind).toBe("#DIV/0!");
  });
  it("propagates errors in the input", () => {
    const out = call("STDEV", [rng([[num(1), err("#NUM!"), num(2)]])]);
    expect(isError(out)).toBe(true);
  });
});

describe("MATH/VAR", () => {
  it("computes sample variance", () => {
    expectClose(ev("=VAR(2,4,4,4,5,5,7,9)"), 32 / 7);
  });
  it("returns #DIV/0! with fewer than 2 values", () => {
    const out = ev("=VAR(1)");
    expect(out.kind).toBe("e");
    if (out.kind === "e") expect(out.v.kind).toBe("#DIV/0!");
  });
  it("ignores text inside ranges", () => {
    expectClose(call("VAR", [rng([[num(2), str("x"), num(4), str("y"), num(6)]])]), 4);
  });
});

describe("MATH/PRODUCT", () => {
  it("multiplies scalar args", () => {
    expect(ev("=PRODUCT(2,3,4)")).toEqual(num(24));
  });
  it("ignores text in ranges", () => {
    expect(call("PRODUCT", [rng([[num(2), str("x"), num(5)]])])).toEqual(num(10));
  });
  it("returns 0 when no numeric values", () => {
    expect(call("PRODUCT", [rng([[str("a")]])])).toEqual(num(0));
  });
});

describe("MATH/SUMPRODUCT", () => {
  it("multiplies element-wise and sums", () => {
    const a = rng([[num(1), num(2), num(3)]]);
    const b = rng([[num(4), num(5), num(6)]]);
    expect(call("SUMPRODUCT", [a, b])).toEqual(num(32));
  });
  it("returns #VALUE! on shape mismatch", () => {
    const a = rng([[num(1), num(2)]]);
    const b = rng([[num(1), num(2), num(3)]]);
    const out = call("SUMPRODUCT", [a, b]);
    expect(out.kind).toBe("e");
    if (out.kind === "e") expect(out.v.kind).toBe("#VALUE!");
  });
  it("treats text cells as 0", () => {
    const a = rng([[num(1), num(2), num(3)]]);
    const b = rng([[num(4), str("x"), num(6)]]);
    expect(call("SUMPRODUCT", [a, b])).toEqual(num(22));
  });
});
