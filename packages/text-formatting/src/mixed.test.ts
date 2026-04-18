import { describe, expect, it } from "vitest";
import { collapse, isMixed, valueOr } from "./mixed";
import { MIXED } from "./types";

describe("collapse", () => {
  it("returns the value when all agree", () => {
    expect(collapse([true, true, true])).toBe(true);
    expect(collapse(["Arial", "Arial"])).toBe("Arial");
  });

  it("returns undefined when all undefined", () => {
    expect(collapse([undefined, undefined])).toBeUndefined();
  });

  it("returns MIXED when values disagree", () => {
    expect(collapse([true, false])).toBe(MIXED);
    expect(collapse(["Arial", "Calibri"])).toBe(MIXED);
  });

  it("returns MIXED when mixing defined and undefined", () => {
    expect(collapse([true, undefined])).toBe(MIXED);
    expect(collapse([undefined, "Arial"])).toBe(MIXED);
  });

  it("supports a custom equality function", () => {
    expect(
      collapse<boolean | string>([true, "single"], (a, b) => {
        const norm = (x: boolean | string) => (x === true ? "single" : x);
        return norm(a) === norm(b);
      })
    ).toBe(true);
  });

  it("returns undefined for empty iterable", () => {
    expect(collapse([])).toBeUndefined();
  });
});

describe("valueOr / isMixed", () => {
  it("valueOr returns value when concrete", () => {
    expect(valueOr<number>(11, 12)).toBe(11);
  });
  it("valueOr returns fallback when MIXED or undefined", () => {
    expect(valueOr<number>(MIXED, 12)).toBe(12);
    expect(valueOr<number>(undefined, 12)).toBe(12);
  });
  it("isMixed", () => {
    expect(isMixed(MIXED)).toBe(true);
    expect(isMixed(true)).toBe(false);
    expect(isMixed(undefined)).toBe(false);
  });
});
