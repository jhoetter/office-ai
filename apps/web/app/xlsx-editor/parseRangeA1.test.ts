import { describe, expect, it } from "vitest";
import { letterToColIndex, parseRangeA1 } from "./parseRangeA1";

describe("parseRangeA1", () => {
  it("parses a single-cell A1 ref", () => {
    expect(parseRangeA1("A1")).toEqual({ r1: 0, c1: 0, r2: 0, c2: 0 });
    expect(parseRangeA1("B4")).toEqual({ r1: 3, c1: 1, r2: 3, c2: 1 });
  });

  it("parses a rectangular range", () => {
    expect(parseRangeA1("B4:D7")).toEqual({ r1: 3, c1: 1, r2: 6, c2: 3 });
  });

  it("normalises a backwards range", () => {
    expect(parseRangeA1("D7:B4")).toEqual({ r1: 3, c1: 1, r2: 6, c2: 3 });
  });

  it("handles multi-letter columns", () => {
    expect(parseRangeA1("AA1")).toEqual({ r1: 0, c1: 26, r2: 0, c2: 26 });
    expect(parseRangeA1("AB10:AC11")).toEqual({ r1: 9, c1: 27, r2: 10, c2: 28 });
  });

  it("rejects malformed input", () => {
    expect(parseRangeA1("")).toBeNull();
    expect(parseRangeA1("Sheet1!A1")).toBeNull();
    expect(parseRangeA1("1A")).toBeNull();
    expect(parseRangeA1("A:B")).toBeNull();
  });
});

describe("letterToColIndex", () => {
  it("matches Excel column conventions", () => {
    expect(letterToColIndex("A")).toBe(0);
    expect(letterToColIndex("Z")).toBe(25);
    expect(letterToColIndex("AA")).toBe(26);
    expect(letterToColIndex("AZ")).toBe(51);
    expect(letterToColIndex("BA")).toBe(52);
  });
});
