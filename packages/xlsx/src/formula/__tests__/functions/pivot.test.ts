import { describe, expect, it } from "vitest";
import { evaluate } from "../../evaluator.js";
import {
  createRegistry,
  type EvalContext,
  type FunctionRegistry,
  type MutableFunctionRegistry,
} from "../../function-registry.js";
import { registerPivotStubs } from "../../functions/pivot.js";
import { parse } from "../../parser.js";
import { AbsRef, type CellRef, type RangeRef } from "../../references.js";
import { num, type Range2D, type Value } from "../../values.js";

const ANCHOR: CellRef = { sheet: "S", row: 0, col: 0, abs: AbsRef.NONE };

function makeRegistry(): MutableFunctionRegistry {
  const reg = createRegistry();
  registerPivotStubs(reg);
  return reg;
}

function makeCtx(): EvalContext {
  const reg: FunctionRegistry = makeRegistry();
  return {
    getCell: (_ref: CellRef): Value => num(0),
    getRange: (_ref: RangeRef): Range2D => [[num(0)]],
    resolveName: () => undefined,
    now: () => 0,
    random: () => 0,
    anchor: ANCHOR,
    registry: reg,
  };
}

function ev(text: string): Value {
  return evaluate(parse(text, { anchor: ANCHOR }).ast, makeCtx());
}

/**
 * Stub functions for pivot/CUBE families. These return `#NAME?` until
 * the pivot evaluator lands, but they MUST be registered so the
 * formula parser doesn't reject the names outright (which would
 * surface as parse errors instead of evaluable cells in fixtures).
 */
describe("pivot stubs", () => {
  it("registers GETPIVOTDATA and returns #NAME?", () => {
    const out = ev('=GETPIVOTDATA("Sum of Sales", $A$3, "Region", "East")');
    expect(out.kind).toBe("e");
    if (out.kind === "e") expect(out.v.kind).toBe("#NAME?");
  });

  it("registers CUBEVALUE and returns #NAME?", () => {
    const out = ev('=CUBEVALUE("ThisWorkbookDataModel","[Measures].[Sum of Sales]")');
    expect(out.kind).toBe("e");
    if (out.kind === "e") expect(out.v.kind).toBe("#NAME?");
  });

  it("registers CUBEMEMBER and returns #NAME?", () => {
    const out = ev('=CUBEMEMBER("conn","[Region].[East]")');
    expect(out.kind).toBe("e");
    if (out.kind === "e") expect(out.v.kind).toBe("#NAME?");
  });

  it("registers CUBESET and returns #NAME?", () => {
    const out = ev('=CUBESET("conn","{[Region].Members}","Regions")');
    expect(out.kind).toBe("e");
    if (out.kind === "e") expect(out.v.kind).toBe("#NAME?");
  });
});
