import { describe, expect, it } from "vitest";
import { anchorsFor, snapToAnchor } from "./anchors.js";

describe("anchorsFor", () => {
  it("returns 5 anchors per shape (4 cardinal mid-edges + centre)", () => {
    const a = anchorsFor("rect-1", { x: 0, y: 0, cx: 1_000_000, cy: 600_000 });
    expect(a.length).toBe(5);
    expect(a.find((x) => x.side === "n")).toEqual({
      shapeId: "rect-1",
      side: "n",
      x: 500_000,
      y: 0,
    });
    expect(a.find((x) => x.side === "s")?.y).toBe(600_000);
    expect(a.find((x) => x.side === "w")?.x).toBe(0);
    expect(a.find((x) => x.side === "e")?.x).toBe(1_000_000);
    expect(a.find((x) => x.side === "center")).toEqual({
      shapeId: "rect-1",
      side: "center",
      x: 500_000,
      y: 300_000,
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
    // Endpoint near shape "a"'s east anchor (1_000_000, 500_000).
    const r = snapToAnchor({ x: 1_000_010, y: 500_005 }, SHAPES, 50_000);
    expect(r.anchor).not.toBeNull();
    expect(r.anchor?.shapeId).toBe("a");
    expect(r.anchor?.side).toBe("e");
    expect(r.dx).toBe(-10);
    expect(r.dy).toBe(-5);
  });

  it("collects every anchor inside the threshold (for UI hinting)", () => {
    // Halfway between centre and east anchor of shape "a" — both within
    // a generous threshold; the closer one wins but both appear in `nearby`.
    const r = snapToAnchor({ x: 750_000, y: 500_000 }, SHAPES, 300_000);
    expect(r.nearby.length).toBeGreaterThanOrEqual(2);
    expect(r.anchor?.shapeId).toBe("a");
  });
});
