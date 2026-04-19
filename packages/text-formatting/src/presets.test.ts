import { describe, expect, it } from "vitest";
import { COLOR_PALETTE, HIGHLIGHT_PALETTE, highlightByDocxName, nearestHighlight } from "./presets";

describe("presets", () => {
  it("color palette swatches are normalised lowercase RRGGBB", () => {
    for (const swatch of COLOR_PALETTE) {
      expect(swatch.hex).toMatch(/^[0-9a-f]{6}$/);
    }
  });

  it("highlight palette swatches are normalised", () => {
    for (const swatch of HIGHLIGHT_PALETTE) {
      expect(swatch.hex).toMatch(/^[0-9a-f]{6}$/);
      expect(swatch.docxName.length).toBeGreaterThan(0);
    }
  });

  it("nearestHighlight maps yellow to yellow", () => {
    expect(nearestHighlight("ffff00").docxName).toBe("yellow");
  });

  it("nearestHighlight maps near-yellow to yellow", () => {
    expect(nearestHighlight("ffee00").docxName).toBe("yellow");
  });

  it("nearestHighlight maps gray to lightGray", () => {
    expect(nearestHighlight("c0c0c0").docxName).toBe("lightGray");
  });

  it("highlightByDocxName round-trips", () => {
    const swatch = highlightByDocxName("yellow");
    expect(swatch?.hex).toBe("ffff00");
    expect(highlightByDocxName("nonsense")).toBeUndefined();
  });
});
