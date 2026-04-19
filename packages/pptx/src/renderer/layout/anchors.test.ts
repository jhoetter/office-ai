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

  it("collects every anchor inside the threshold (for UI hinting)", () => {
    // Halfway between centre and east anchor of shape "a" — both within
    // a generous threshold; the closer one wins but both appear in `nearby`.
    const r = snapToAnchor({ x: 750_000, y: 500_000 }, SHAPES, 300_000);
    expect(r.nearby.length).toBeGreaterThanOrEqual(2);
    expect(r.anchor?.shapeId).toBe("a");
  });
});
