import { describe, expect, it } from "vitest";
import { ErrorKinds } from "../errors.js";
import { lex } from "../lexer.js";

function types(src: string): string[] {
  return lex(src)
    .filter((t) => t.type !== "EOF")
    .map((t) => t.type);
}

describe("formula/lexer", () => {
  it("strips a leading =", () => {
    expect(types("=1+2")).toEqual(["NUMBER", "OPERATOR", "NUMBER"]);
  });

  it("tokenizes numbers in all forms", () => {
    expect(lex("=42")[0].value).toBe(42);
    expect(lex("=3.14")[0].value).toBe(3.14);
    expect(lex("=.5")[0].value).toBe(0.5);
    expect(lex("=1.5e-3")[0].value).toBeCloseTo(0.0015);
    expect(lex("=1E5")[0].value).toBe(100000);
  });

  it("tokenizes string literals with doubled quotes", () => {
    const t = lex('="he said ""hi"""')[0];
    expect(t.type).toBe("STRING");
    expect(t.value).toBe('he said "hi"');
  });

  it("recognises booleans (case-insensitive) and not as functions", () => {
    expect(lex("=TRUE")[0].type).toBe("BOOL");
    expect(lex("=false")[0].value).toBe(false);
    expect(lex("=TRUE()")[0].type).toBe("FUNCTION");
  });

  it("recognises error literals", () => {
    expect(lex("=#DIV/0!")[0].type).toBe("ERROR");
    expect(lex("=#REF!")[0].value).toEqual({ kind: ErrorKinds.REF });
  });

  it("emits FUNCTION for ident-followed-by-(", () => {
    const ts = lex("=SUM(A1:A3)");
    expect(ts[0]).toMatchObject({ type: "FUNCTION", value: "SUM" });
    expect(ts[1]).toMatchObject({ type: "LPAREN" });
    expect(ts[2]).toMatchObject({ type: "RANGE_REF" });
  });

  it("recognises absolute and sheet-qualified refs", () => {
    expect(lex("=$A$1")[0].type).toBe("REF");
    expect(lex("=Sheet1!A1")[0]).toMatchObject({ type: "REF", text: "Sheet1!A1" });
    expect(lex("='My Sheet'!A1")[0]).toMatchObject({ type: "REF", text: "'My Sheet'!A1" });
  });

  it("recognises whole-row and whole-column ranges", () => {
    expect(lex("=A:A")[0].type).toBe("RANGE_REF");
    expect(lex("=3:5")[0].type).toBe("RANGE_REF");
  });

  it("emits two-char comparison operators", () => {
    expect(lex("=A1<>B1").map((t) => t.text)).toContain("<>");
    expect(lex("=A1<=B1").map((t) => t.text)).toContain("<=");
    expect(lex("=A1>=B1").map((t) => t.text)).toContain(">=");
  });

  it("treats whitespace as separator only", () => {
    expect(types("= 1 +  2 ")).toEqual(["NUMBER", "OPERATOR", "NUMBER"]);
  });

  it("throws on unterminated string", () => {
    expect(() => lex('="open')).toThrow(/Unterminated string/);
  });

  it("throws on unknown character", () => {
    expect(() => lex("=`bad`")).toThrow(/Unexpected character/);
  });

  it("emits SEMICOLON so parser can short-circuit deferred ops", () => {
    expect(types("=A1;B1")).toEqual(["REF", "SEMICOLON", "REF"]);
  });

  it("treats percent as its own token", () => {
    expect(types("=50%")).toEqual(["NUMBER", "PERCENT"]);
  });

  it("supports negative-exponent numbers", () => {
    expect(lex("=1e-2")[0].value).toBe(0.01);
  });
});
