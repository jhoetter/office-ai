import type { Shape } from "../../model/types.js";

export interface BoundingBox {
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
}

/**
 * Return the bounding box of a shape in EMU, or null when it has no
 * explicit position/size (e.g. a placeholder inheriting from layout/master).
 */
export function shapeBoundingBox(shape: Shape): BoundingBox | null {
  if (!shape.position || !shape.size) return null;
  return {
    x: shape.position.xEmu,
    y: shape.position.yEmu,
    cx: shape.size.cxEmu,
    cy: shape.size.cyEmu,
  };
}

/** Whether two boxes intersect (used for hit-testing & selection). */
export function boxesIntersect(a: BoundingBox, b: BoundingBox): boolean {
  return !(a.x + a.cx < b.x || b.x + b.cx < a.x || a.y + a.cy < b.y || b.y + b.cy < a.y);
}

/** Whether a point lies within a box. */
export function pointInBox(box: BoundingBox, x: number, y: number): boolean {
  return x >= box.x && x <= box.x + box.cx && y >= box.y && y <= box.y + box.cy;
}
