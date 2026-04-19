import { describe, expect, it } from "vitest";
import type { PptxSnapshot } from "@officeai/pptx";
import { parseSlideRange, snapshotToSlideSvg } from "./export-images";

/**
 * Pure helpers from `export-images` get unit coverage here.
 * `snapshotToPng/Jpeg` and the SVG-rendering happy path need a real
 * snapshot from a fixture and a DOM canvas, respectively — those are
 * covered by the Playwright specs.
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

describe("snapshotToSlideSvg", () => {
  it("throws a clear error when the slide index is out of range", () => {
    const empty = { root: { slides: [] } } as unknown as PptxSnapshot;
    expect(() => snapshotToSlideSvg(empty, 0)).toThrow(/out of range/i);
    expect(() => snapshotToSlideSvg(empty, 7)).toThrow(/out of range/i);
  });
});
