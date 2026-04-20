import { describe, expect, it } from "vitest";
import type { BoundingBox } from "./shape.js";
import { resolveRotatedResize, type ResizeHandle } from "./resize.js";

/** Origin box used in most cases: 100×100 EMU centred at (150, 150). */
const O: BoundingBox = { x: 100, y: 100, cx: 100, cy: 100 };

/** Approximate equality for floating-point bbox math. */
function expectClose(actual: BoundingBox, expected: BoundingBox, eps = 1e-6): void {
  expect(actual.x).toBeCloseTo(expected.x, 6);
  expect(actual.y).toBeCloseTo(expected.y, 6);
  expect(actual.cx).toBeCloseTo(expected.cx, 6);
  expect(actual.cy).toBeCloseTo(expected.cy, 6);
  void eps;
}

/**
 * Compute the screen-space corner of the rotated body at a given
 * local handle. Used by the round-trip tests below to assert that
 * the OPPOSITE corner stays anchored.
 */
function rotatedCorner(box: BoundingBox, rotDeg: number, h: ResizeHandle): { x: number; y: number } {
  const r = (rotDeg * Math.PI) / 180;
  const cos = Math.cos(r);
  const sin = Math.sin(r);
  const cx = box.x + box.cx / 2;
  const cy = box.y + box.cy / 2;
  let lx = 0;
  let ly = 0;
  if (h.includes("e")) lx = box.cx / 2;
  if (h.includes("w")) lx = -box.cx / 2;
  if (h.includes("s")) ly = box.cy / 2;
  if (h.includes("n")) ly = -box.cy / 2;
  return {
    x: cx + (lx * cos - ly * sin),
    y: cy + (lx * sin + ly * cos),
  };
}

const ALL_HANDLES: ResizeHandle[] = ["n", "ne", "e", "se", "s", "sw", "w", "nw"];

const opposite: Record<ResizeHandle, ResizeHandle> = {
  n: "s",
  ne: "sw",
  e: "w",
  se: "nw",
  s: "n",
  sw: "ne",
  w: "e",
  nw: "se",
};

describe("resolveRotatedResize", () => {
  describe("rot = 0 parity with the legacy axis-aligned formulas", () => {
    // The rotated helper must match the unrotated formulas at rot=0
    // so callers can route every shape through it without breaking
    // the byte-identical no-touch save invariant.
    const cases: Array<{ h: ResizeHandle; dx: number; dy: number; want: BoundingBox }> = [
      { h: "e", dx: 30, dy: 0, want: { x: 100, y: 100, cx: 130, cy: 100 } },
      { h: "w", dx: -20, dy: 0, want: { x: 80, y: 100, cx: 120, cy: 100 } },
      { h: "s", dx: 0, dy: 40, want: { x: 100, y: 100, cx: 100, cy: 140 } },
      { h: "n", dx: 0, dy: -25, want: { x: 100, y: 75, cx: 100, cy: 125 } },
      { h: "ne", dx: 10, dy: -10, want: { x: 100, y: 90, cx: 110, cy: 110 } },
      { h: "nw", dx: -10, dy: -10, want: { x: 90, y: 90, cx: 110, cy: 110 } },
      { h: "se", dx: 10, dy: 10, want: { x: 100, y: 100, cx: 110, cy: 110 } },
      { h: "sw", dx: -10, dy: 10, want: { x: 90, y: 100, cx: 110, cy: 110 } },
    ];

    it.each(cases)("$h handle drag (dx=$dx, dy=$dy)", ({ h, dx, dy, want }) => {
      const got = resolveRotatedResize({
        o: O,
        rotDeg: 0,
        h,
        dxEmu: dx,
        dyEmu: dy,
        minSize: 0,
      });
      expectClose(got, want);
    });
  });

  describe("rot != 0: opposite corner stays anchored in screen space", () => {
    // For every (rotation, handle) combination, the opposite corner
    // of the rotated body must NOT move regardless of the cursor
    // delta. This is the screen-space anchor invariant the
    // PowerPoint-style behaviour is built on.
    for (const rotDeg of [15, 45, 90, 135, 200, 350]) {
      for (const h of ALL_HANDLES) {
        it(`rot=${rotDeg}°, handle=${h}`, () => {
          const anchorBefore = rotatedCorner(O, rotDeg, opposite[h]);
          // Pick a cursor delta that grows the box on at least one
          // axis so the test exercises real motion.
          const got = resolveRotatedResize({
            o: O,
            rotDeg,
            h,
            dxEmu: 25,
            dyEmu: -15,
            minSize: 0,
          });
          const anchorAfter = rotatedCorner(got, rotDeg, opposite[h]);
          expect(anchorAfter.x).toBeCloseTo(anchorBefore.x, 6);
          expect(anchorAfter.y).toBeCloseTo(anchorBefore.y, 6);
        });
      }
    }
  });

  describe("rot = 90: local axes are rotated 90° CW", () => {
    // Drag the "e" handle by (0, +30) screen-space EMU. With rot=90,
    // local +x points down on screen, so a downward cursor delta
    // grows the LOCAL width (cx). Visually the rotated body extends
    // further south on screen.
    it("e-handle: downward screen drag grows local cx", () => {
      const got = resolveRotatedResize({
        o: O,
        rotDeg: 90,
        h: "e",
        dxEmu: 0,
        dyEmu: 30,
        minSize: 0,
      });
      expect(got.cx).toBeCloseTo(130, 6);
      expect(got.cy).toBeCloseTo(100, 6);
    });

    // Drag the "s" handle by (-30, 0) screen-space EMU. With rot=90,
    // local +y points west on screen, so a leftward cursor delta
    // grows the LOCAL height (cy).
    it("s-handle: leftward screen drag grows local cy", () => {
      const got = resolveRotatedResize({
        o: O,
        rotDeg: 90,
        h: "s",
        dxEmu: -30,
        dyEmu: 0,
        minSize: 0,
      });
      expect(got.cx).toBeCloseTo(100, 6);
      expect(got.cy).toBeCloseTo(130, 6);
    });
  });

  describe("rot = 45: corner drag along local diagonal", () => {
    it("se-handle: drag along screen +y grows the local diagonal proportionally", () => {
      // At 45°, local +x is screen (+x, +y)/√2 and local +y is
      // screen (-x, +y)/√2. A pure (+y) drag projects equally onto
      // both local axes (each gets +30/√2), so cx and cy grow by
      // the same amount.
      const got = resolveRotatedResize({
        o: O,
        rotDeg: 45,
        h: "se",
        dxEmu: 0,
        dyEmu: 30,
        minSize: 0,
      });
      const grow = 30 / Math.SQRT2;
      expect(got.cx).toBeCloseTo(100 + grow, 6);
      expect(got.cy).toBeCloseTo(100 + grow, 6);
      // And the screen-space NW corner (opposite of se) must stay
      // pinned (already covered by the round-trip suite, asserted
      // here for documentation).
      const before = rotatedCorner(O, 45, "nw");
      const after = rotatedCorner(got, 45, "nw");
      expect(after.x).toBeCloseTo(before.x, 6);
      expect(after.y).toBeCloseTo(before.y, 6);
    });
  });

  describe("minSize floor", () => {
    it("does not shrink the local size below minSize", () => {
      const got = resolveRotatedResize({
        o: O,
        rotDeg: 30,
        h: "e",
        dxEmu: -1000,
        dyEmu: -1000,
        minSize: 40,
      });
      expect(got.cx).toBeGreaterThanOrEqual(40);
      // cy is untouched by the "e" handle.
      expect(got.cy).toBeCloseTo(100, 6);
    });

    it("clamps both axes when both are dragged inward at minSize", () => {
      const got = resolveRotatedResize({
        o: O,
        rotDeg: 30,
        h: "se",
        dxEmu: -10000,
        dyEmu: -10000,
        minSize: 25,
      });
      expect(got.cx).toBeGreaterThanOrEqual(25);
      expect(got.cy).toBeGreaterThanOrEqual(25);
    });

    it("minSize=0 (line-shape policy) lets dimensions collapse to zero", () => {
      const got = resolveRotatedResize({
        o: O,
        rotDeg: 30,
        h: "e",
        dxEmu: -10000,
        dyEmu: 0,
        minSize: 0,
      });
      expect(got.cx).toBeGreaterThanOrEqual(0);
    });
  });
});
