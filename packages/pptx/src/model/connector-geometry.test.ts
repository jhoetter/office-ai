import { describe, expect, it } from "vitest";
import { anchorPoint, type Box } from "./connector-geometry.js";

const BOX: Box = { x: 0, y: 0, cx: 1_000_000, cy: 600_000 };

describe("anchorPoint quarter-point interpolation", () => {
  it("returns the cardinal midpoint when t is undefined", () => {
    expect(anchorPoint(BOX, "n")).toEqual({ x: 500_000, y: 0 });
    expect(anchorPoint(BOX, "s")).toEqual({ x: 500_000, y: 600_000 });
    expect(anchorPoint(BOX, "w")).toEqual({ x: 0, y: 300_000 });
    expect(anchorPoint(BOX, "e")).toEqual({ x: 1_000_000, y: 300_000 });
  });

  it("interpolates along the picked edge for valid t in [0, 1]", () => {
    expect(anchorPoint(BOX, "n", 0.25)).toEqual({ x: 250_000, y: 0 });
    expect(anchorPoint(BOX, "n", 0.75)).toEqual({ x: 750_000, y: 0 });
    expect(anchorPoint(BOX, "e", 0.25)).toEqual({ x: 1_000_000, y: 150_000 });
    expect(anchorPoint(BOX, "w", 0.75)).toEqual({ x: 0, y: 450_000 });
  });

  it("clamps t outside [0, 1] to the nearest corner", () => {
    expect(anchorPoint(BOX, "n", -1)).toEqual({ x: 0, y: 0 });
    expect(anchorPoint(BOX, "n", 2)).toEqual({ x: 1_000_000, y: 0 });
  });

  it("falls back to the midpoint for non-finite t (NaN/Infinity)", () => {
    expect(anchorPoint(BOX, "n", Number.NaN)).toEqual({ x: 500_000, y: 0 });
    expect(anchorPoint(BOX, "e", Number.POSITIVE_INFINITY)).toEqual({ x: 1_000_000, y: 300_000 });
  });

  it("ignores t for the center anchor (always exact centre)", () => {
    expect(anchorPoint(BOX, "center", 0.25)).toEqual({ x: 500_000, y: 300_000 });
    expect(anchorPoint(BOX, "center", 0.99)).toEqual({ x: 500_000, y: 300_000 });
  });
});
