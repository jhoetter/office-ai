import { describe, expect, it } from "vitest";
import { evaluate } from "../../evaluator.js";
import { registerInfo } from "../../functions/info.js";
import { createRegistry, type EvalContext, type MutableFunctionRegistry } from "../../function-registry.js";
import { parse } from "../../parser.js";
import { AbsRef, type CellRef, type RangeRef } from "../../references.js";
import { Blank, bool, err, num, str, type Value } from "../../values.js";

const ANCHOR: CellRef = { sheet: "Sheet1", row: 0, col: 0, abs: AbsRef.NONE };

function makeRegistry(): MutableFunctionRegistry {
  const reg = createRegistry();
  registerInfo(reg);
  return reg;
}

function makeContext(
  opts: {
    cells?: Record<string, Value>;
    ranges?: Record<string, Value[][]>;
    registry?: MutableFunctionRegistry;
  } = {}
): EvalContext {
  const cells = opts.cells ?? {};
  const ranges = opts.ranges ?? {};
  const registry = opts.registry ?? makeRegistry();
  return {
    getCell(ref: CellRef): Value {
      const k = `${ref.sheet}!R${ref.row}C${ref.col}`;
      return cells[k] ?? Blank;
    },
    getRange(ref: RangeRef): readonly (readonly Value[])[] {
      const k = `${ref.sheet}!R${ref.r0}C${ref.c0}:R${ref.r1}C${ref.c1}`;
      return ranges[k] ?? [[Blank]];
    },
    resolveName: () => undefined,
    now: () => 0,
    random: () => 0,
    anchor: ANCHOR,
    registry,
  };
}

function evalText(text: string, ctx?: EvalContext): Value {
  const ast = parse(text, { anchor: ANCHOR }).ast;
  return evaluate(ast, ctx ?? makeContext());
}

describe("formula/functions/info — ISBLANK", () => {
  it("returns TRUE for the Blank sentinel via cell reference", () => {
    const ctx = makeContext({ cells: { "Sheet1!R0C0": Blank } });
    expect(evalText("=ISBLANK(A1)", ctx)).toEqual(bool(true));
  });

  it("returns FALSE for an explicit zero", () => {
    expect(evalText("=ISBLANK(0)")).toEqual(bool(false));
  });

  it("returns FALSE for an empty string (Excel parity)", () => {
    expect(evalText('=ISBLANK("")')).toEqual(bool(false));
  });

  it("returns FALSE for an error value", () => {
    expect(evalText("=ISBLANK(#N/A)")).toEqual(bool(false));
  });
});

describe("formula/functions/info — ISNUMBER", () => {
  it("returns TRUE for a numeric literal", () => {
    expect(evalText("=ISNUMBER(42)")).toEqual(bool(true));
  });

  it("returns FALSE for a text literal", () => {
    expect(evalText('=ISNUMBER("42")')).toEqual(bool(false));
  });

  it("returns FALSE — never propagates — when given an error", () => {
    expect(evalText("=ISNUMBER(1/0)")).toEqual(bool(false));
  });

  it("returns FALSE for a boolean", () => {
    expect(evalText("=ISNUMBER(TRUE)")).toEqual(bool(false));
  });
});

describe("formula/functions/info — ISTEXT", () => {
  it("returns TRUE for a text literal", () => {
    expect(evalText('=ISTEXT("hello")')).toEqual(bool(true));
  });

  it("returns FALSE for a number", () => {
    expect(evalText("=ISTEXT(1)")).toEqual(bool(false));
  });

  it("returns FALSE for an error (no propagation)", () => {
    expect(evalText("=ISTEXT(#N/A)")).toEqual(bool(false));
  });

  it("returns TRUE for an empty string", () => {
    expect(evalText('=ISTEXT("")')).toEqual(bool(true));
  });
});

describe("formula/functions/info — ISERROR", () => {
  it("returns TRUE for #DIV/0!", () => {
    expect(evalText("=ISERROR(1/0)")).toEqual(bool(true));
  });

  it("returns TRUE for #N/A as well", () => {
    expect(evalText("=ISERROR(#N/A)")).toEqual(bool(true));
  });

  it("returns FALSE for a plain string", () => {
    expect(evalText('=ISERROR("hi")')).toEqual(bool(false));
  });

  it("returns FALSE for a number", () => {
    expect(evalText("=ISERROR(0)")).toEqual(bool(false));
  });
});

describe("formula/functions/info — ISNA", () => {
  it("returns TRUE for #N/A", () => {
    expect(evalText("=ISNA(#N/A)")).toEqual(bool(true));
  });

  it("returns FALSE for #DIV/0!", () => {
    expect(evalText("=ISNA(1/0)")).toEqual(bool(false));
  });

  it("returns FALSE for a non-error value", () => {
    expect(evalText("=ISNA(1)")).toEqual(bool(false));
  });

  it("returns FALSE for a string", () => {
    expect(evalText('=ISNA("#N/A")')).toEqual(bool(false));
  });
});

describe("formula/functions/info — ISODD", () => {
  it("returns TRUE for an odd integer", () => {
    expect(evalText("=ISODD(3)")).toEqual(bool(true));
  });

  it("returns FALSE for an even integer", () => {
    expect(evalText("=ISODD(4)")).toEqual(bool(false));
  });

  it("truncates fractional values toward zero before parity", () => {
    expect(evalText("=ISODD(3.9)")).toEqual(bool(true));
    expect(evalText("=ISODD(2.5)")).toEqual(bool(false));
  });

  it("propagates errors", () => {
    const out = evalText("=ISODD(1/0)");
    expect(out.kind).toBe("e");
  });

  it("handles negative integers", () => {
    expect(evalText("=ISODD(-7)")).toEqual(bool(true));
    expect(evalText("=ISODD(-2)")).toEqual(bool(false));
  });
});

describe("formula/functions/info — ISEVEN", () => {
  it("returns TRUE for an even integer", () => {
    expect(evalText("=ISEVEN(4)")).toEqual(bool(true));
  });

  it("returns FALSE for an odd integer", () => {
    expect(evalText("=ISEVEN(3)")).toEqual(bool(false));
  });

  it("truncates fractional values toward zero before parity", () => {
    expect(evalText("=ISEVEN(2.9)")).toEqual(bool(true));
    expect(evalText("=ISEVEN(3.5)")).toEqual(bool(false));
  });

  it("propagates errors", () => {
    const out = evalText("=ISEVEN(1/0)");
    expect(out.kind).toBe("e");
  });

  it("returns TRUE for zero", () => {
    expect(evalText("=ISEVEN(0)")).toEqual(bool(true));
  });
});

describe("formula/functions/info — TYPE", () => {
  it("returns 1 for a number", () => {
    expect(evalText("=TYPE(42)")).toEqual(num(1));
  });

  it("returns 2 for text", () => {
    expect(evalText('=TYPE("hi")')).toEqual(num(2));
  });

  it("returns 4 for a logical", () => {
    expect(evalText("=TYPE(TRUE)")).toEqual(num(4));
  });

  it("returns 16 for an error", () => {
    expect(evalText("=TYPE(#N/A)")).toEqual(num(16));
  });

  it("returns 64 for an array literal", () => {
    expect(evalText("=TYPE({1,2;3,4})")).toEqual(num(64));
  });

  it("returns 1 for the Blank sentinel (consistent with N(blank)=0)", () => {
    const ctx = makeContext({ cells: { "Sheet1!R0C0": Blank } });
    expect(evalText("=TYPE(A1)", ctx)).toEqual(num(1));
  });
});

describe("formula/functions/info — N", () => {
  it("returns the number unchanged", () => {
    expect(evalText("=N(42)")).toEqual(num(42));
  });

  it("coerces TRUE → 1 and FALSE → 0", () => {
    expect(evalText("=N(TRUE)")).toEqual(num(1));
    expect(evalText("=N(FALSE)")).toEqual(num(0));
  });

  it("returns 0 for any text (even text that looks numeric)", () => {
    expect(evalText('=N("hello")')).toEqual(num(0));
    expect(evalText('=N("42")')).toEqual(num(0));
  });

  it("propagates errors", () => {
    const out = evalText("=N(#N/A)");
    expect(out).toEqual(err("#N/A"));
  });

  it("returns 0 for the Blank sentinel", () => {
    const ctx = makeContext({ cells: { "Sheet1!R0C0": Blank } });
    expect(evalText("=N(A1)", ctx)).toEqual(num(0));
  });
});

describe("formula/functions/info — NA", () => {
  it("returns #N/A", () => {
    expect(evalText("=NA()")).toEqual(err("#N/A"));
  });

  it("propagates through arithmetic", () => {
    const out = evalText("=NA()+1");
    expect(out).toEqual(err("#N/A"));
  });

  it("is detected by ISNA", () => {
    expect(evalText("=ISNA(NA())")).toEqual(bool(true));
  });

  it("rejects any argument with #N/A (arity)", () => {
    const out = evalText("=NA(1)");
    expect(out.kind).toBe("e");
  });
});

describe("formula/functions/info — registry surface", () => {
  it("registers all ten P0 info functions", () => {
    const reg = makeRegistry();
    for (const name of [
      "ISBLANK",
      "ISNUMBER",
      "ISTEXT",
      "ISERROR",
      "ISNA",
      "ISODD",
      "ISEVEN",
      "TYPE",
      "N",
      "NA",
    ]) {
      expect(reg.has(name)).toBe(true);
    }
  });

  it("does not mark any info function as volatile", () => {
    const reg = makeRegistry();
    expect(reg.volatileNames().size).toBe(0);
  });

  it("ISBLANK identity check uses the Blank singleton (not value-equality)", () => {
    const reg = makeRegistry();
    const impl = reg.get("ISBLANK");
    expect(impl).toBeDefined();
    if (impl && !impl.lazyArgs) {
      // Direct invocation: handing the Blank singleton returns TRUE,
      // a freshly-constructed `num(0)` returns FALSE — even though
      // they are deeply equal. This is the documented limitation.
      expect(impl.fn([Blank], makeContext())).toEqual(bool(true));
      expect(impl.fn([num(0)], makeContext())).toEqual(bool(false));
      expect(impl.fn([str("")], makeContext())).toEqual(bool(false));
    }
  });
});
