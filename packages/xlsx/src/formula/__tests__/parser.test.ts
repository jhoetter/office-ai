import { describe, expect, it } from "vitest";
import type { AstNode, BinaryNode, CallNode, RefNode } from "../ast.js";
import { parse } from "../parser.js";
import { AbsRef, parseA1, type CellRef } from "../references.js";

const ANCHOR: CellRef = { sheet: "Sheet1", row: 0, col: 0, abs: AbsRef.NONE };

function parseAst(text: string): AstNode {
  return parse(text, { anchor: ANCHOR }).ast;
}

describe("formula/parser — primaries", () => {
  it("parses literals", () => {
    expect(parseAst("=42")).toMatchObject({ kind: "lit", value: { kind: "n", v: 42 } });
    expect(parseAst('="hi"')).toMatchObject({ kind: "lit", value: { kind: "s", v: "hi" } });
    expect(parseAst("=TRUE")).toMatchObject({ kind: "lit", value: { kind: "b", v: true } });
  });

  it("parses cell and range refs", () => {
    expect(parseAst("=A1")).toMatchObject({ kind: "ref" });
    expect(parseAst("=A1:B5")).toMatchObject({ kind: "range" });
    expect(parseAst("=Sheet2!A1")).toMatchObject({
      kind: "ref",
      ref: { sheet: "Sheet2" },
    });
  });

  it("parses function calls with multiple args", () => {
    const ast = parseAst("=SUM(A1, A2, A3)") as CallNode;
    expect(ast.kind).toBe("call");
    expect(ast.name).toBe("SUM");
    expect(ast.args).toHaveLength(3);
  });

  it("upper-cases function names", () => {
    const ast = parseAst("=sum(1,2)") as CallNode;
    expect(ast.name).toBe("SUM");
  });

  it("parses parenthesised expressions", () => {
    const ast = parseAst("=(1+2)*3") as BinaryNode;
    expect(ast.op).toBe("*");
  });

  it("parses array literals", () => {
    const ast = parseAst("={1,2;3,4}");
    expect(ast.kind).toBe("array");
    if (ast.kind === "array") {
      expect(ast.rows.length).toBe(2);
      expect(ast.rows[0].length).toBe(2);
    }
  });
});

describe("formula/parser — operator precedence", () => {
  it("multiplication binds tighter than addition", () => {
    const ast = parseAst("=1+2*3") as BinaryNode;
    expect(ast.op).toBe("+");
    expect((ast.right as BinaryNode).op).toBe("*");
  });

  it("exponent is right-associative", () => {
    const ast = parseAst("=2^3^2") as BinaryNode;
    expect(ast.op).toBe("^");
    expect((ast.right as BinaryNode).op).toBe("^");
  });

  it("unary minus binds tighter than ^ (Excel quirk)", () => {
    // -2^2 == 4 in Excel: parses as (-(2))^2.
    const ast = parseAst("=-2^2") as BinaryNode;
    expect(ast.op).toBe("^");
    expect(ast.left.kind).toBe("unary");
  });

  it("comparison binds loosest", () => {
    const ast = parseAst("=1+2<3*4") as BinaryNode;
    expect(ast.op).toBe("<");
    expect((ast.left as BinaryNode).op).toBe("+");
    expect((ast.right as BinaryNode).op).toBe("*");
  });

  it("string concat binds between additive and comparison", () => {
    const ast = parseAst('="x"&"y"="xy"') as BinaryNode;
    expect(ast.op).toBe("=");
    expect((ast.left as BinaryNode).op).toBe("&");
  });

  it("percent binds as a postfix unary", () => {
    const ast = parseAst("=50%");
    expect(ast.kind).toBe("pct");
  });
});

describe("formula/parser — dependencies and volatility", () => {
  it("collects unique cell references across a formula", () => {
    const f = parse("=A1+B2+A1", { anchor: ANCHOR });
    expect(f.dependencies).toHaveLength(2);
  });

  it("flags RAND/NOW/INDIRECT as volatile", () => {
    expect(parse("=RAND()", { anchor: ANCHOR }).volatile).toBe(true);
    expect(parse('=INDIRECT("A1")', { anchor: ANCHOR }).volatile).toBe(true);
    expect(parse("=A1+1", { anchor: ANCHOR }).volatile).toBe(false);
  });
});

describe("formula/parser — error paths", () => {
  it("rejects empty formula", () => {
    expect(() => parse("=", { anchor: ANCHOR })).toThrow(/empty/i);
  });

  it("rejects intersection operator", () => {
    expect(() => parse("=A1;B1", { anchor: ANCHOR })).toThrow(/intersection/);
  });

  it("rejects unbalanced parens", () => {
    expect(() => parse("=(1+2", { anchor: ANCHOR })).toThrow();
  });

  it("rejects stray operators", () => {
    // Excel parses `A1B` as a name (resolved at eval time as #NAME?), so
    // the parse-time error path here uses an unmistakable syntax error.
    expect(() => parse("=*1", { anchor: ANCHOR })).toThrow();
  });
});

describe("formula/parser — defined names", () => {
  it("resolves defined names to refs at parse time", () => {
    const named = parseA1("Sheet1!Z9", "Sheet1")!;
    const f = parse("=MyName+1", {
      anchor: ANCHOR,
      definedNames: new Map([["MyName", named]]),
    });
    const left = (f.ast as BinaryNode).left as RefNode;
    expect(left.kind).toBe("ref");
    expect(left.ref).toEqual(named);
  });
});
