import { describe, expect, it } from "vitest";
import {
  normalizeRangeForStorage,
  parseChartRangeShape,
  pickToggleDefaults,
  validateChartShape,
} from "./chartShape";

describe("parseChartRangeShape", () => {
  it("parses a multi-cell range", () => {
    expect(parseChartRangeShape("A1:B5")).toEqual({ r1: 0, c1: 0, r2: 4, c2: 1 });
  });

  it("parses a single-column range", () => {
    expect(parseChartRangeShape("A1:A4")).toEqual({ r1: 0, c1: 0, r2: 3, c2: 0 });
  });

  it("parses $-anchored references the same as bare ones", () => {
    expect(parseChartRangeShape("$A$1:$B$5")).toEqual(parseChartRangeShape("A1:B5"));
  });

  it("strips an optional sheet prefix", () => {
    expect(parseChartRangeShape("Sheet1!A1:B5")).toEqual({ r1: 0, c1: 0, r2: 4, c2: 1 });
  });

  it("returns null for empty / invalid input", () => {
    expect(parseChartRangeShape("")).toBeNull();
    expect(parseChartRangeShape("   ")).toBeNull();
    expect(parseChartRangeShape("garbage")).toBeNull();
  });
});

describe("normalizeRangeForStorage", () => {
  it("uppercases and strips $ anchors", () => {
    expect(normalizeRangeForStorage("$a$1:$b$5")).toBe("A1:B5");
  });

  it("preserves a sheet prefix", () => {
    expect(normalizeRangeForStorage("Sheet1!a1:b5")).toBe("SHEET1!A1:B5");
  });
});

describe("pickToggleDefaults", () => {
  it("turns both toggles off for a single-column range like A1:A4", () => {
    expect(pickToggleDefaults(parseChartRangeShape("A1:A4"))).toEqual({
      hasHeaderRow: false,
      hasCategoryColumn: false,
    });
  });

  it("turns both toggles off for a single-row range like A1:D1", () => {
    expect(pickToggleDefaults(parseChartRangeShape("A1:D1"))).toEqual({
      hasHeaderRow: false,
      hasCategoryColumn: false,
    });
  });

  it("keeps Excel-parity defaults on for a multi-row + multi-col range", () => {
    expect(pickToggleDefaults(parseChartRangeShape("A1:B5"))).toEqual({
      hasHeaderRow: true,
      hasCategoryColumn: true,
    });
  });

  it("falls back to both-on when the range is unparseable", () => {
    expect(pickToggleDefaults(null)).toEqual({ hasHeaderRow: true, hasCategoryColumn: true });
  });
});

describe("validateChartShape — Bug A regression", () => {
  it("approves A1:A4 with the new shape-aware defaults (no 'No data' state)", () => {
    const range = "A1:A4";
    const shape = parseChartRangeShape(range);
    const defaults = pickToggleDefaults(shape);
    expect(validateChartShape(range, shape, defaults.hasHeaderRow, defaults.hasCategoryColumn)).toEqual({
      kind: "ok",
    });
  });

  it("flags A1:A4 with the old hard-coded both-on defaults as 'no-values' (the actual bug)", () => {
    const range = "A1:A4";
    const shape = parseChartRangeShape(range);
    expect(validateChartShape(range, shape, true, true)).toEqual({
      kind: "no-values",
      axis: "column",
    });
  });

  it("approves a normal multi-cell range with both toggles on", () => {
    const range = "A1:B5";
    const shape = parseChartRangeShape(range);
    expect(validateChartShape(range, shape, true, true)).toEqual({ kind: "ok" });
  });

  it("flags a single-cell selection", () => {
    const range = "C3:C3";
    const shape = parseChartRangeShape(range);
    expect(validateChartShape(range, shape, false, false)).toEqual({ kind: "single-cell" });
  });

  it("flags an empty range as 'empty', not 'invalid'", () => {
    expect(validateChartShape("", null, true, true)).toEqual({ kind: "empty" });
  });

  it("flags a syntactically broken range as 'invalid'", () => {
    expect(validateChartShape("garbage", null, true, true)).toEqual({ kind: "invalid" });
  });
});
