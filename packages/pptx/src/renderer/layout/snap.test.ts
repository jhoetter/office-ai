import { describe, expect, it } from "vitest";
import type { SlideSize } from "../../model/types.js";
import { computeSnap } from "./snap.js";

const SLIDE: SlideSize = { cxEmu: 9_144_000, cyEmu: 6_858_000 };
const T = 50_000; // ≈ 5 px @ 96 DPI

describe("computeSnap", () => {
  it("returns zero deltas + no guides when nothing is within threshold", () => {
    const r = computeSnap(
      { x: 1_000_000, y: 1_000_000, cx: 800_000, cy: 800_000 },
      [{ id: "other", box: { x: 5_000_000, y: 5_000_000, cx: 800_000, cy: 800_000 } }],
      SLIDE,
      T
    );
    expect(r.snapDx).toBe(0);
    expect(r.snapDy).toBe(0);
    expect(r.guides.length).toBe(0);
  });

  it("snaps left edge to another shape's left edge when only that edge is within threshold", () => {
    // Moving cx=500k; ref cx=2_000_000 (different sizes) so centres
    // and right edges DON'T align → only the left-edge snap fires.
    const moving = { x: 1_000_010, y: 2_000_000, cx: 500_000, cy: 500_000 };
    const others = [
      { id: "ref", box: { x: 1_000_000, y: 4_000_000, cx: 2_000_000, cy: 500_000 } },
    ];
    const r = computeSnap(moving, others, SLIDE, T);
    expect(r.snapDx).toBe(-10);
    expect(r.snapDy).toBe(0);
    expect(r.guides.length).toBe(1);
    expect(r.guides[0].axis).toBe("vertical");
    expect(r.guides[0].value).toBe(1_000_000);
    expect(r.guides[0].kind).toBe("edge");
  });

  it("snaps to slide horizontal centre", () => {
    const moving = { x: 4_572_000 - 500_000 / 2 + 30, y: 0, cx: 500_000, cy: 500_000 };
    const r = computeSnap(moving, [], SLIDE, T);
    expect(r.snapDx).toBe(-30);
    expect(r.guides.some((g) => g.axis === "vertical" && g.kind === "slide")).toBe(true);
  });

  it("prefers the centre snap over an edge snap when both are within threshold and tied", () => {
    // Layout: ref shape's right edge is at x = 1.0M; ref shape's centre is at 0.5M.
    // Moving box (cx = 1.0M, x ≈ 0): its centre is at ~0.5M (matches ref centre)
    // and its right edge is at ~1.0M (matches ref right). Both within threshold.
    const moving = { x: 5, y: 2_000_000, cx: 1_000_000, cy: 500_000 };
    const others = [
      { id: "ref", box: { x: 0, y: 4_000_000, cx: 1_000_000, cy: 500_000 } },
    ];
    const r = computeSnap(moving, others, SLIDE, T);
    expect(r.snapDx).toBe(-5);
    // The chosen guide should reference the centre value (500_000), not the edge.
    expect(r.guides.some((g) => g.axis === "vertical" && g.value === 500_000)).toBe(true);
  });

  it("snaps both axes independently when both are eligible", () => {
    const moving = { x: 1_000_020, y: 2_000_015, cx: 500_000, cy: 500_000 };
    const others = [
      { id: "ref", box: { x: 1_000_000, y: 2_000_000, cx: 500_000, cy: 500_000 } },
    ];
    const r = computeSnap(moving, others, SLIDE, T);
    expect(r.snapDx).toBe(-20);
    expect(r.snapDy).toBe(-15);
    expect(r.guides.filter((g) => g.axis === "vertical").length).toBeGreaterThan(0);
    expect(r.guides.filter((g) => g.axis === "horizontal").length).toBeGreaterThan(0);
  });
});
