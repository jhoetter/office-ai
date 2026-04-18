import { describe, expect, it } from "vitest";
import { boxesIntersect, pointInBox, shapeBoundingBox } from "./shape.js";
import type { TextShape } from "../../model/types.js";

const ts = (x: number, y: number, cx: number, cy: number): TextShape => ({
  kind: "text",
  id: "x",
  cNvPrId: 1,
  name: "n",
  position: { xEmu: x, yEmu: y },
  size: { cxEmu: cx, cyEmu: cy },
  txBody: { paragraphs: [] },
  nvSpPrTail: [],
  spPrTail: [],
});

describe("shape geometry helpers", () => {
  it("returns a bounding box when both pos and size are present", () => {
    expect(shapeBoundingBox(ts(1, 2, 3, 4))).toEqual({ x: 1, y: 2, cx: 3, cy: 4 });
  });

  it("box intersection is symmetric", () => {
    const a = { x: 0, y: 0, cx: 10, cy: 10 };
    const b = { x: 5, y: 5, cx: 10, cy: 10 };
    const c = { x: 100, y: 100, cx: 1, cy: 1 };
    expect(boxesIntersect(a, b)).toBe(true);
    expect(boxesIntersect(b, a)).toBe(true);
    expect(boxesIntersect(a, c)).toBe(false);
  });

  it("point-in-box hits and misses", () => {
    const b = { x: 0, y: 0, cx: 10, cy: 10 };
    expect(pointInBox(b, 5, 5)).toBe(true);
    expect(pointInBox(b, 11, 5)).toBe(false);
  });
});
