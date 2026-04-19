import { describe, expect, it } from "vitest";
import { anchorsFor, snapToAnchor } from "./anchors.js";

describe("anchorsFor", () => {
  it("exposes three points per cardinal edge (quarter / mid / three-quarter) plus the centre", () => {
    const a = anchorsFor("rect-1", { x: 0, y: 0, cx: 1_000_000, cy: 600_000 });
    // 4 sides * 3 points + 1 centre = 13.
    expect(a.length).toBe(13);
    const north = a.filter((x) => x.side === "n");
    expect(north.map((p) => p.t).sort()).toEqual([0.25, 0.5, 0.75]);
    expect(north.map((p) => p.x).sort((x, y) => x - y)).toEqual([250_000, 500_000, 750_000]);
    expect(north.every((p) => p.y === 0)).toBe(true);
    expect(a.find((x) => x.side === "center")).toEqual({
      shapeId: "rect-1",
      side: "center",
      x: 500_000,
      y: 300_000,
      t: 0.5,
    });
  });
});

describe("snapToAnchor", () => {
  const SHAPES = [
    { id: "a", box: { x: 0, y: 0, cx: 1_000_000, cy: 1_000_000 } },
    { id: "b", box: { x: 5_000_000, y: 5_000_000, cx: 1_000_000, cy: 1_000_000 } },
  ];

  it("returns no anchor when nothing is within threshold", () => {
    const r = snapToAnchor({ x: 3_000_000, y: 3_000_000 }, SHAPES, 100_000);
    expect(r.anchor).toBeNull();
    expect(r.dx).toBe(0);
    expect(r.dy).toBe(0);
    expect(r.nearby.length).toBe(0);
  });

  it("snaps to the closest anchor within threshold", () => {
    // Endpoint near shape "a"'s east mid-edge anchor (1_000_000, 500_000).
    const r = snapToAnchor({ x: 1_000_010, y: 500_005 }, SHAPES, 50_000);
    expect(r.anchor).not.toBeNull();
    expect(r.anchor?.shapeId).toBe("a");
    expect(r.anchor?.side).toBe("e");
    expect(r.anchor?.t).toBe(0.5);
    expect(r.dx).toBe(-10);
    expect(r.dy).toBe(-5);
  });

  it("snaps to a quarter-point anchor when it's the closest", () => {
    // Aim at shape "a"'s east edge at t=0.25 → (1_000_000, 250_000).
    const r = snapToAnchor({ x: 1_000_000, y: 260_000 }, SHAPES, 50_000);
    expect(r.anchor?.side).toBe("e");
    expect(r.anchor?.t).toBe(0.25);
  });

  it("collects edge candidates inside the threshold (for UI hinting)", () => {
    // Aim near the middle of shape "a"'s east edge so multiple
    // edge anchors (mid + the two quarter-points) sit within the
    // threshold. The closer one wins but every edge candidate
    // appears in `nearby`. Centre anchors are intentionally hidden
    // when any edge is in range — they are not snap targets.
    const r = snapToAnchor({ x: 1_000_000, y: 500_000 }, SHAPES, 350_000);
    expect(r.nearby.length).toBeGreaterThanOrEqual(2);
    expect(r.nearby.every((a) => a.side !== "center")).toBe(true);
    expect(r.anchor?.shapeId).toBe("a");
    expect(r.anchor?.side).toBe("e");
  });

  it("prefers an edge anchor over a closer centre anchor", () => {
    // Cursor sits at the centre of shape "a" (500k, 500k). The centre
    // anchor is exactly here (distance 0); the closest edge anchors
    // are 500k EMU away. With a generous threshold both are in range,
    // but edges must still win — connectors that snap to the centre
    // produce routes that visibly cut through the shape.
    const r = snapToAnchor({ x: 500_000, y: 500_000 }, SHAPES, 600_000);
    expect(r.anchor).not.toBeNull();
    expect(r.anchor?.shapeId).toBe("a");
    expect(r.anchor?.side).not.toBe("center");
  });

  it("falls back to centre when no edge anchor is in range", () => {
    // Cursor at shape "a"'s centroid with a tight threshold that only
    // covers the centre anchor itself.
    const r = snapToAnchor({ x: 500_000, y: 500_000 }, SHAPES, 1);
    expect(r.anchor?.shapeId).toBe("a");
    expect(r.anchor?.side).toBe("center");
  });
});
