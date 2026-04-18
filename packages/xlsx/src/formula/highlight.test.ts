import { describe, expect, it } from "vitest";

import { assignRefColors, DEFAULT_REF_COLORS, tokenizeForDisplay, type DisplayToken } from "./highlight.js";

const concat = (tokens: ReadonlyArray<DisplayToken>): string => tokens.map((t) => t.text).join("");

describe("tokenizeForDisplay", () => {
  it("returns an empty array for empty input", () => {
    expect(tokenizeForDisplay("")).toEqual([]);
  });

  it("emits the leading `=` as its own text token", () => {
    const out = tokenizeForDisplay("=1");
    expect(out[0]).toMatchObject({ start: 0, end: 1, text: "=", kind: "text" });
  });

  it("classifies refs, ranges, functions, numbers, ops and punct", () => {
    const out = tokenizeForDisplay("=SUM(A1:B2, 3) + C4");
    const kinds = out.map((t) => t.kind);
    expect(kinds).toContain("function");
    expect(kinds).toContain("range");
    expect(kinds).toContain("ref");
    expect(kinds).toContain("number");
    expect(kinds).toContain("operator");
    expect(kinds).toContain("punct");
  });

  it("normalises absolute / mixed refs to the same refKey", () => {
    const out = tokenizeForDisplay("=A1+$A$1+a1+$A1");
    const refKeys = out.filter((t) => t.kind === "ref").map((t) => t.refKey);
    expect(refKeys).toEqual(["A1", "A1", "A1", "A1"]);
  });

  it("includes the sheet name in the refKey when present", () => {
    const out = tokenizeForDisplay("=Sheet2!A1");
    const ref = out.find((t) => t.kind === "ref");
    expect(ref?.refKey).toBe("SHEET2!A1");
    expect(ref?.target).toMatchObject({ kind: "ref", sheet: "Sheet2", row: 0, col: 0 });
  });

  it("handles quoted sheet names with embedded apostrophes", () => {
    const out = tokenizeForDisplay("='Sheet 1'!B2");
    const ref = out.find((t) => t.kind === "ref");
    expect(ref?.refKey).toBe("SHEET 1!B2");
    expect(ref?.target).toMatchObject({ kind: "ref", sheet: "Sheet 1", row: 1, col: 1 });
  });

  it("resolves range targets to inclusive normalised rectangles", () => {
    const out = tokenizeForDisplay("=B2:A1");
    const range = out.find((t) => t.kind === "range");
    expect(range?.refKey).toBe("B2:A1");
    expect(range?.target).toMatchObject({ kind: "range", r1: 0, c1: 0, r2: 1, c2: 1 });
  });

  it("never throws on partial / malformed input mid-typing", () => {
    expect(() => tokenizeForDisplay("=A1+")).not.toThrow();
    expect(() => tokenizeForDisplay('=SUM("hello')).not.toThrow();
    expect(() => tokenizeForDisplay("=??")).not.toThrow();
  });

  it("covers the input contiguously (no gaps, no overlaps)", () => {
    const samples = ["=SUM(A1:B2,3)+C4", '=IF(A1>0,"hi","bye")', "=A1+", "=??"];
    for (const s of samples) {
      const tokens = tokenizeForDisplay(s);
      expect(concat(tokens)).toBe(s);
      for (let i = 1; i < tokens.length; i++) {
        expect(tokens[i]!.start).toBe(tokens[i - 1]!.end);
      }
    }
  });

  it("does not classify a function name without an opening paren", () => {
    const out = tokenizeForDisplay("=SUMTHING");
    expect(out.some((t) => t.kind === "function")).toBe(false);
  });
});

describe("assignRefColors", () => {
  it("returns a map keyed by refKey, palette-cycled", () => {
    const tokens = tokenizeForDisplay("=A1+B2+A1+C3");
    const colors = assignRefColors(tokens);
    expect(colors.get("A1")).toBe(DEFAULT_REF_COLORS[0]);
    expect(colors.get("B2")).toBe(DEFAULT_REF_COLORS[1]);
    expect(colors.get("C3")).toBe(DEFAULT_REF_COLORS[2]);
    expect(colors.size).toBe(3);
  });

  it("ignores non-ref tokens", () => {
    const tokens = tokenizeForDisplay('=SUM(1,2,"x")');
    expect(assignRefColors(tokens).size).toBe(0);
  });

  it("respects a custom palette", () => {
    const tokens = tokenizeForDisplay("=A1+B2");
    const colors = assignRefColors(tokens, ["#aaa", "#bbb"]);
    expect(colors.get("A1")).toBe("#aaa");
    expect(colors.get("B2")).toBe("#bbb");
  });

  it("wraps around the palette when more refs than colours", () => {
    const tokens = tokenizeForDisplay("=A1+B2+C3");
    const colors = assignRefColors(tokens, ["#x", "#y"]);
    expect(colors.get("A1")).toBe("#x");
    expect(colors.get("B2")).toBe("#y");
    expect(colors.get("C3")).toBe("#x");
  });
});
