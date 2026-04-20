import { describe, expect, it } from "vitest";
import type { PdfEngineTextItem } from "../types.js";
import { buildGlyphRuns } from "./build-glyph-runs.js";

describe("buildGlyphRuns", () => {
  it("produces one glyph per character with even subdivision of the run width", () => {
    const item: PdfEngineTextItem = {
      str: "Hello",
      // Identity-scale text matrix at (100, 200) with font size 12.
      transform: [12, 0, 0, 12, 100, 200],
      width: 60, // 5 chars × 12pt advance ≈ 60pt
      height: 12,
      fontName: "Helv",
    };
    const runs = buildGlyphRuns([item]);
    expect(runs.length).toBe(1);
    const run = runs[0];
    expect(run.chars).toBe("Hello");
    expect(run.glyphs.length).toBe(5);
    // Glyph 0 starts at the text origin; glyph 4 ends at origin + width.
    expect(run.glyphs[0][0]).toBeCloseTo(100);
    expect(run.glyphs[4][2]).toBeCloseTo(160);
    // Every glyph spans the full font height in y.
    for (const g of run.glyphs) {
      expect(g[3] - g[1]).toBeCloseTo(12);
    }
    expect(run.baselineY).toBe(200);
    expect(run.fontKey).toBe("Helv");
    expect(run.dir).toBe("ltr");
  });

  it("flags rtl when the matrix x-basis points left", () => {
    const item: PdfEngineTextItem = {
      str: "אבג",
      transform: [-12, 0, 0, 12, 200, 200],
      width: 36,
      height: 12,
    };
    const runs = buildGlyphRuns([item]);
    expect(runs[0].dir).toBe("rtl");
    // Glyph rects remain valid bounding boxes regardless of direction.
    for (const g of runs[0].glyphs) {
      expect(g[2]).toBeGreaterThan(g[0]);
      expect(g[3]).toBeGreaterThan(g[1]);
    }
  });

  it("skips empty items", () => {
    expect(
      buildGlyphRuns([
        { str: "", transform: [1, 0, 0, 1, 0, 0], width: 0, height: 0 },
      ]),
    ).toEqual([]);
  });
});
