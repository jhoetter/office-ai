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

describe("anchorPoint rotation", () => {
  // 90° clockwise about the centre of a 1000×600 box at the origin
  // sends north → east, east → south, south → west, west → north.
  // Centre stays put. The arithmetic uses the *real* (un-rounded)
  // centre, not the rounded `Math.round` from the legacy helper, so
  // an even-dimensioned box round-trips exactly.
  const ROT_BOX: Box = { x: 0, y: 0, cx: 1_000_000, cy: 600_000 };

  it("90° rotation: cardinal midpoints rotate clockwise about the centre", () => {
    // Centre at (500_000, 300_000). North midpoint (500_000, 0) →
    // (500_000 + 0, 300_000 + (-300_000)*0 - (-500_000? no))…
    // Concretely: after 90° CW the north dot lands where the east
    // dot used to be: (500_000 + 300_000, 300_000 + 0) =
    // (800_000, 300_000).
    expect(closeTo(anchorPoint(ROT_BOX, "n", 0.5, 90), { x: 800_000, y: 300_000 })).toBe(true);
    expect(closeTo(anchorPoint(ROT_BOX, "e", 0.5, 90), { x: 500_000, y: 800_000 })).toBe(true);
    expect(closeTo(anchorPoint(ROT_BOX, "s", 0.5, 90), { x: 200_000, y: 300_000 })).toBe(true);
    expect(closeTo(anchorPoint(ROT_BOX, "w", 0.5, 90), { x: 500_000, y: -200_000 })).toBe(true);
  });

  it("180° rotation: the centre is invariant, opposite cardinals swap", () => {
    expect(closeTo(anchorPoint(ROT_BOX, "n", 0.5, 180), { x: 500_000, y: 600_000 })).toBe(true);
    expect(closeTo(anchorPoint(ROT_BOX, "s", 0.5, 180), { x: 500_000, y: 0 })).toBe(true);
    expect(closeTo(anchorPoint(ROT_BOX, "center", 0.5, 180), { x: 500_000, y: 300_000 })).toBe(true);
  });

  it("0° rotation matches the un-rotated arithmetic (no rounding drift)", () => {
    expect(anchorPoint(ROT_BOX, "n", 0.5, 0)).toEqual({ x: 500_000, y: 0 });
    expect(anchorPoint(ROT_BOX, "e", 0.25, 0)).toEqual({ x: 1_000_000, y: 150_000 });
  });
});

function closeTo(a: { x: number; y: number }, b: { x: number; y: number }, eps = 1): boolean {
  return Math.abs(a.x - b.x) <= eps && Math.abs(a.y - b.y) <= eps;
}
