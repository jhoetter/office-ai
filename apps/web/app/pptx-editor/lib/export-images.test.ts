import { describe, expect, it } from "vitest";
import { parseSlideRange } from "./export-images";

/**
 * The slide-range mini-DSL is the only piece of `export-images`
 * that's reasonable to unit-test in Node — `snapshotToPngZip` needs
 * a DOM canvas and `snapshotToSvgZip` needs an in-memory snapshot
 * built from a real fixture (covered by the e2e specs).
 */
describe("parseSlideRange", () => {
  it("returns every slide for empty input", () => {
    expect(parseSlideRange("", 5)).toEqual([0, 1, 2, 3, 4]);
    expect(parseSlideRange("   ", 3)).toEqual([0, 1, 2]);
  });

  it("parses a single 1-based index", () => {
    expect(parseSlideRange("3", 5)).toEqual([2]);
  });

  it("parses comma-separated singletons", () => {
    expect(parseSlideRange("1,3,5", 5)).toEqual([0, 2, 4]);
  });

  it("parses inclusive ranges", () => {
    expect(parseSlideRange("2-4", 5)).toEqual([1, 2, 3]);
  });

  it("mixes singletons and ranges", () => {
    expect(parseSlideRange("1,3-5,7", 8)).toEqual([0, 2, 3, 4, 6]);
  });

  it("dedupes overlapping ranges and singletons", () => {
    expect(parseSlideRange("1-3,2,3-4", 5)).toEqual([0, 1, 2, 3]);
  });

  it("normalizes reversed ranges", () => {
    expect(parseSlideRange("4-2", 5)).toEqual([1, 2, 3]);
  });

  it("drops out-of-range entries silently", () => {
    expect(parseSlideRange("0,1,99", 3)).toEqual([0]);
    expect(parseSlideRange("8-10", 5)).toEqual([]);
  });

  it("ignores garbage tokens", () => {
    expect(parseSlideRange("abc,2,xx-yy", 5)).toEqual([1]);
  });

  it("returns sorted indices regardless of input order", () => {
    expect(parseSlideRange("5,1,3", 5)).toEqual([0, 2, 4]);
  });

  it("tolerates extra whitespace", () => {
    expect(parseSlideRange(" 1 , 3 - 5 ", 5)).toEqual([0, 2, 3, 4]);
  });
});
