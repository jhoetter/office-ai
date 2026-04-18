import { describe, expect, it } from "vitest";
import type { AstNode } from "../ast.js";
import { parse } from "../parser.js";
import { AbsRef, type CellRef } from "../references.js";
import { serializeAst } from "../serialize-ast.js";

const ANCHOR: CellRef = { sheet: "Sheet1", row: 0, col: 0, abs: AbsRef.NONE };

function parseAst(text: string): AstNode {
  return parse(text, { anchor: ANCHOR }).ast;
}

function roundtrip(text: string, anchorSheet = ANCHOR.sheet): string {
  return serializeAst(parseAst(text), anchorSheet);
}

describe("formula/serialize-ast — literals", () => {
  it("serializes integer numbers", () => {
    expect(serializeAst(parseAst("=42"))).toBe("42");
  });

  it("serializes fractional numbers", () => {
    expect(serializeAst(parseAst("=3.14"))).toBe("3.14");
  });

  it("serializes string literals with embedded quotes doubled", () => {
    expect(serializeAst(parseAst('="he said ""hi"""'))).toBe('"he said ""hi"""');
  });

  it("serializes booleans as upper-case TRUE/FALSE", () => {
    expect(serializeAst(parseAst("=TRUE"))).toBe("TRUE");
    expect(serializeAst(parseAst("=FALSE"))).toBe("FALSE");
  });

  it("serializes error literals", () => {
    expect(serializeAst(parseAst("=#REF!"))).toBe("#REF!");
    expect(serializeAst(parseAst("=#DIV/0!"))).toBe("#DIV/0!");
  });
});

describe("formula/serialize-ast — references", () => {
  it("emits a bare cell ref when sheet matches anchor", () => {
    expect(serializeAst(parseAst("=A1"), "Sheet1")).toBe("A1");
  });

  it("preserves absolute markers on cell refs", () => {
    expect(serializeAst(parseAst("=$A$1"), "Sheet1")).toBe("$A$1");
    expect(serializeAst(parseAst("=A$1"), "Sheet1")).toBe("A$1");
    expect(serializeAst(parseAst("=$A1"), "Sheet1")).toBe("$A1");
  });

  it("prefixes refs from a different sheet", () => {
    expect(serializeAst(parseAst("=Sheet2!A1"), "Sheet1")).toBe("Sheet2!A1");
  });

  it("emits a range ref bare when sheet matches anchor", () => {
    expect(serializeAst(parseAst("=A1:B5"), "Sheet1")).toBe("A1:B5");
  });

  it("preserves absolute markers on range endpoints", () => {
    expect(serializeAst(parseAst("=$A$1:$B$5"), "Sheet1")).toBe("$A$1:$B$5");
  });
});

describe("formula/serialize-ast — operators + calls", () => {
  it("serializes a function call with multiple args", () => {
    expect(serializeAst(parseAst("=SUM(A1, A2, A3)"), "Sheet1")).toBe("SUM(A1,A2,A3)");
  });

  it("upper-cases function names", () => {
    expect(serializeAst(parseAst("=sum(1,2)"))).toBe("SUM(1,2)");
  });

  it("serializes a binary op", () => {
    expect(serializeAst(parseAst("=1+2"))).toBe("1+2");
  });

  it("parenthesises binary children of binary ops", () => {
    expect(serializeAst(parseAst("=(1+2)*3"))).toBe("(1+2)*3");
  });

  it("serializes unary minus with operand parens when nested", () => {
    expect(serializeAst(parseAst("=-A1"), "Sheet1")).toBe("-A1");
    expect(serializeAst(parseAst("=-(A1+1)"), "Sheet1")).toBe("-(A1+1)");
  });

  it("serializes percent suffix", () => {
    expect(serializeAst(parseAst("=A1%"), "Sheet1")).toBe("A1%");
  });

  it("serializes array literals row-major with `;` between rows", () => {
    expect(serializeAst(parseAst("={1,2;3,4}"))).toBe("{1,2;3,4}");
  });

  it("serializes a defined-name token", () => {
    expect(serializeAst(parseAst("=foo"))).toBe("foo");
  });

  it("round-trips a SUM expression with a cross-sheet range", () => {
    expect(serializeAst(parseAst("=SUM(Sheet2!A1:B5)+1"), "Sheet1")).toBe("SUM(Sheet2!A1:B5)+1");
  });
});

describe("formula/serialize-ast — semantic round-trip", () => {
  it("re-parses to an equivalent AST", () => {
    const cases = ["=1+2*3", "=A1+SUM(B1:B3)", "=IF(A1>0,1,-1)", "=Sheet2!A1*2"];
    for (const text of cases) {
      const out = roundtrip(text);
      const reparsed = parse(out, { anchor: ANCHOR });
      expect(reparsed.ast.kind).toBe(parse(text, { anchor: ANCHOR }).ast.kind);
    }
  });
});
