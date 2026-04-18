import { describe, expect, it } from "vitest";
import { evaluate } from "../../evaluator.js";
import {
  arity,
  createRegistry,
  type EvalContext,
  type FunctionRegistry,
  type MutableFunctionRegistry,
} from "../../function-registry.js";
import { registerLogic } from "../../functions/logic.js";
import { parse } from "../../parser.js";
import { AbsRef, type CellRef, type RangeRef } from "../../references.js";
import { bool, err, isError, num, str, type Value } from "../../values.js";

const ANCHOR: CellRef = { sheet: "Sheet1", row: 0, col: 0, abs: AbsRef.NONE };

function makeContext(opts: {
  cells?: Record<string, Value>;
  ranges?: Record<string, Value[][]>;
  registry?: FunctionRegistry;
}): EvalContext {
  const cells = opts.cells ?? {};
  const ranges = opts.ranges ?? {};
  const registry = opts.registry ?? createRegistry();
  return {
    getCell(ref: CellRef): Value {
      const k = `${ref.sheet}!R${ref.row}C${ref.col}`;
      return cells[k] ?? num(0);
    },
    getRange(ref: RangeRef) {
      const k = `${ref.sheet}!R${ref.r0}C${ref.c0}:R${ref.r1}C${ref.c1}`;
      return ranges[k] ?? [[num(0)]];
    },
    resolveName: () => undefined,
    now: () => 0,
    random: () => 0,
    anchor: ANCHOR,
    registry,
  };
}

function buildRegistry(): MutableFunctionRegistry {
  const reg = createRegistry();
  registerLogic(reg);
  return reg;
}

function evalText(
  text: string,
  reg: FunctionRegistry,
  extra?: Partial<Parameters<typeof makeContext>[0]>
): Value {
  const ast = parse(text, { anchor: ANCHOR }).ast;
  return evaluate(ast, makeContext({ registry: reg, ...extra }));
}

describe("logic — IF", () => {
  it("returns the then-branch on truthy condition", () => {
    const reg = buildRegistry();
    expect(evalText('=IF(TRUE, "yes", "no")', reg)).toEqual(str("yes"));
    expect(evalText("=IF(1, 10, 20)", reg)).toEqual(num(10));
  });

  it("returns the else-branch on falsy condition; defaults to FALSE when omitted", () => {
    const reg = buildRegistry();
    expect(evalText("=IF(FALSE, 10, 20)", reg)).toEqual(num(20));
    expect(evalText("=IF(0, 10, 20)", reg)).toEqual(num(20));
    expect(evalText('=IF(TRUE, "yes")', reg)).toEqual(str("yes"));
    expect(evalText("=IF(FALSE, 10)", reg)).toEqual(bool(false));
  });

  it("propagates errors and rejects non-coercible string conditions", () => {
    const reg = buildRegistry();
    const div0 = evalText("=IF(1/0, 1, 2)", reg);
    expect(isError(div0)).toBe(true);
    const badStr = evalText('=IF("maybe", 1, 2)', reg);
    expect(isError(badStr)).toBe(true);
    if (badStr.kind === "e") expect(badStr.v.kind).toBe("#VALUE!");
  });

  it("evaluates only the chosen branch (lazy)", () => {
    const reg = buildRegistry();
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
    expect(evalText("=IF(TRUE, LCOUNT(1), RCOUNT(2))", reg)).toEqual(num(1));
    expect(leftCalls).toBe(1);
    expect(rightCalls).toBe(0);
    expect(evalText("=IF(FALSE, LCOUNT(1), RCOUNT(2))", reg)).toEqual(num(2));
    expect(leftCalls).toBe(1);
    expect(rightCalls).toBe(1);
  });
});

describe("logic — IFS", () => {
  it("returns the first matching branch's value", () => {
    const reg = buildRegistry();
    expect(evalText('=IFS(FALSE, "a", TRUE, "b", TRUE, "c")', reg)).toEqual(str("b"));
  });

  it("returns #N/A when no condition matches", () => {
    const reg = buildRegistry();
    const out = evalText('=IFS(FALSE, "a", FALSE, "b")', reg);
    expect(isError(out)).toBe(true);
    if (out.kind === "e") expect(out.v.kind).toBe("#N/A");
  });

  it("evaluates only matching branch (lazy)", () => {
    const reg = buildRegistry();
    let calls = 0;
    reg.register({
      name: "TRACE",
      arity: arity(1, 1),
      fn: (args) => {
        calls++;
        return args[0];
      },
    });
    expect(evalText("=IFS(FALSE, TRACE(1), TRUE, TRACE(2), TRUE, TRACE(3))", reg)).toEqual(num(2));
    expect(calls).toBe(1);
  });

  it("propagates errors from the condition under test", () => {
    const reg = buildRegistry();
    const out = evalText('=IFS(1/0, "a", TRUE, "b")', reg);
    expect(isError(out)).toBe(true);
  });
});

describe("logic — AND", () => {
  it("returns TRUE only when every arg is truthy", () => {
    const reg = buildRegistry();
    expect(evalText("=AND(TRUE, TRUE, 1)", reg)).toEqual(bool(true));
    expect(evalText("=AND(TRUE, FALSE, TRUE)", reg)).toEqual(bool(false));
    expect(evalText("=AND(1, 2, 3)", reg)).toEqual(bool(true));
  });

  it("flattens range args; errors propagate; strings inside a range are ignored", () => {
    const reg = buildRegistry();
    const ctx = makeContext({
      registry: reg,
      ranges: {
        "Sheet1!R0C0:R0C2": [[bool(true), str("ignored"), num(7)]],
      },
    });
    const ast = parse("=AND(A1:C1)", { anchor: ANCHOR }).ast;
    expect(evaluate(ast, ctx)).toEqual(bool(true));
  });

  it("propagates errors", () => {
    const reg = buildRegistry();
    const out = evalText("=AND(TRUE, 1/0)", reg);
    expect(isError(out)).toBe(true);
  });
});

describe("logic — OR", () => {
  it("returns TRUE if any arg is truthy", () => {
    const reg = buildRegistry();
    expect(evalText("=OR(FALSE, FALSE, TRUE)", reg)).toEqual(bool(true));
    expect(evalText("=OR(0, 0, 0)", reg)).toEqual(bool(false));
    expect(evalText("=OR(0, 5)", reg)).toEqual(bool(true));
  });

  it("propagates errors from later args (parity: all eagerly evaluated)", () => {
    const reg = buildRegistry();
    const out = evalText("=OR(TRUE, 1/0)", reg);
    expect(isError(out)).toBe(true);
  });

  it("rejects non-coercible string args with #VALUE!", () => {
    const reg = buildRegistry();
    const out = evalText('=OR("maybe")', reg);
    expect(isError(out)).toBe(true);
    if (out.kind === "e") expect(out.v.kind).toBe("#VALUE!");
  });
});

describe("logic — NOT", () => {
  it("inverts the boolean coercion of its arg", () => {
    const reg = buildRegistry();
    expect(evalText("=NOT(TRUE)", reg)).toEqual(bool(false));
    expect(evalText("=NOT(FALSE)", reg)).toEqual(bool(true));
    expect(evalText("=NOT(0)", reg)).toEqual(bool(true));
    expect(evalText("=NOT(7)", reg)).toEqual(bool(false));
  });

  it("propagates errors", () => {
    const reg = buildRegistry();
    const out = evalText("=NOT(1/0)", reg);
    expect(isError(out)).toBe(true);
  });

  it("returns #VALUE! on non-coercible string", () => {
    const reg = buildRegistry();
    const out = evalText('=NOT("maybe")', reg);
    expect(isError(out)).toBe(true);
    if (out.kind === "e") expect(out.v.kind).toBe("#VALUE!");
  });
});

describe("logic — XOR", () => {
  it("returns TRUE iff an odd number of args are truthy", () => {
    const reg = buildRegistry();
    expect(evalText("=XOR(TRUE)", reg)).toEqual(bool(true));
    expect(evalText("=XOR(TRUE, TRUE)", reg)).toEqual(bool(false));
    expect(evalText("=XOR(TRUE, TRUE, TRUE)", reg)).toEqual(bool(true));
    expect(evalText("=XOR(FALSE, FALSE)", reg)).toEqual(bool(false));
  });

  it("counts truthy values across a range arg", () => {
    const reg = buildRegistry();
    const ctx = makeContext({
      registry: reg,
      ranges: {
        "Sheet1!R0C0:R0C2": [[num(1), num(0), num(1)]],
      },
    });
    const ast = parse("=XOR(A1:C1)", { anchor: ANCHOR }).ast;
    expect(evaluate(ast, ctx)).toEqual(bool(false));
  });

  it("propagates errors", () => {
    const reg = buildRegistry();
    const out = evalText("=XOR(TRUE, 1/0)", reg);
    expect(isError(out)).toBe(true);
  });
});

describe("logic — IFERROR", () => {
  it("returns the value when not an error", () => {
    const reg = buildRegistry();
    expect(evalText('=IFERROR(42, "fallback")', reg)).toEqual(num(42));
  });

  it("catches every error including #N/A", () => {
    const reg = buildRegistry();
    expect(evalText('=IFERROR(1/0, "fallback")', reg)).toEqual(str("fallback"));
    expect(evalText("=IFERROR(#N/A, 0)", reg)).toEqual(num(0));
    expect(evalText("=IFERROR(#REF!, 0)", reg)).toEqual(num(0));
  });

  it("does not evaluate the fallback when the value is fine (lazy)", () => {
    const reg = buildRegistry();
    let fallbackCalls = 0;
    reg.register({
      name: "FB",
      arity: arity(0, 0),
      fn: () => {
        fallbackCalls++;
        return num(99);
      },
    });
    expect(evalText("=IFERROR(7, FB())", reg)).toEqual(num(7));
    expect(fallbackCalls).toBe(0);
  });
});

describe("logic — IFNA", () => {
  it("catches #N/A only", () => {
    const reg = buildRegistry();
    expect(evalText('=IFNA(#N/A, "missing")', reg)).toEqual(str("missing"));
  });

  it("does not catch other errors", () => {
    const reg = buildRegistry();
    const out = evalText('=IFNA(#DIV/0!, "missing")', reg);
    expect(isError(out)).toBe(true);
    if (out.kind === "e") expect(out.v.kind).toBe("#DIV/0!");
  });

  it("returns the value when not #N/A and skips fallback (lazy)", () => {
    const reg = buildRegistry();
    let calls = 0;
    reg.register({
      name: "FB",
      arity: arity(0, 0),
      fn: () => {
        calls++;
        return num(0);
      },
    });
    expect(evalText("=IFNA(5, FB())", reg)).toEqual(num(5));
    expect(calls).toBe(0);
  });
});

describe("logic — SWITCH", () => {
  it("returns the value paired with the first matching case", () => {
    const reg = buildRegistry();
    expect(evalText('=SWITCH(2, 1, "a", 2, "b", 3, "c")', reg)).toEqual(str("b"));
    expect(evalText('=SWITCH("X", "x", "lower", "X", "upper")', reg)).toEqual(str("lower"));
  });

  it("uses trailing odd arg as default; #N/A if no match and no default", () => {
    const reg = buildRegistry();
    expect(evalText('=SWITCH(9, 1, "a", 2, "b", "default")', reg)).toEqual(str("default"));
    const noMatch = evalText('=SWITCH(9, 1, "a", 2, "b")', reg);
    expect(isError(noMatch)).toBe(true);
    if (noMatch.kind === "e") expect(noMatch.v.kind).toBe("#N/A");
  });

  it("evaluates only the chosen branch (lazy)", () => {
    const reg = buildRegistry();
    let calls = 0;
    reg.register({
      name: "TRACE",
      arity: arity(1, 1),
      fn: (args) => {
        calls++;
        return args[0];
      },
    });
    expect(evalText("=SWITCH(2, 1, TRACE(10), 2, TRACE(20), 3, TRACE(30))", reg)).toEqual(num(20));
    expect(calls).toBe(1);
  });

  it("propagates errors from the expression", () => {
    const reg = buildRegistry();
    const out = evalText('=SWITCH(1/0, 1, "a")', reg);
    expect(isError(out)).toBe(true);
  });
});

describe("logic — TRUE / FALSE", () => {
  it("TRUE() returns BoolValue(true)", () => {
    const reg = buildRegistry();
    expect(evalText("=TRUE()", reg)).toEqual(bool(true));
  });

  it("FALSE() returns BoolValue(false)", () => {
    const reg = buildRegistry();
    expect(evalText("=FALSE()", reg)).toEqual(bool(false));
  });

  it("compose with operators as plain booleans", () => {
    const reg = buildRegistry();
    expect(evalText("=IF(TRUE(), 1, 2)", reg)).toEqual(num(1));
    expect(evalText("=IF(FALSE(), 1, 2)", reg)).toEqual(num(2));
  });
});

describe("logic — registry surface", () => {
  it("registers all eleven Logic P0 functions", () => {
    const reg = buildRegistry();
    for (const name of [
      "IF",
      "IFS",
      "AND",
      "OR",
      "NOT",
      "XOR",
      "IFERROR",
      "IFNA",
      "SWITCH",
      "TRUE",
      "FALSE",
    ]) {
      expect(reg.has(name)).toBe(true);
    }
  });

  it("rejects bad arities at the dispatcher with #N/A", () => {
    const reg = buildRegistry();
    const out = evalText("=NOT(1, 2)", reg);
    expect(isError(out)).toBe(true);
    if (out.kind === "e") expect(out.v.kind).toBe("#N/A");
  });

  it("IFERROR catches an error read from a cell", () => {
    const reg = buildRegistry();
    const ctx = makeContext({
      registry: reg,
      cells: { "Sheet1!R0C0": err("#REF!") },
    });
    const ast = parse('=IFERROR(A1, "ok")', { anchor: ANCHOR }).ast;
    expect(evaluate(ast, ctx)).toEqual(str("ok"));
  });
});
