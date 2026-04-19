import type { RouterBox, RouterObstacle, RouterPoint } from "./types.js";

/**
 * How much to inflate every obstacle box so connectors visibly clear
 * the shape edge instead of grazing it. Roughly ~0.094 inch — close
 * to the visual "halo" PowerPoint leaves between the routed line and
 * the source/target shape.
 */
export const OBSTACLE_PAD_EMU = 90_000;

/**
 * Inflate a bounding box by `pad` EMU on every side. Negative pads
 * are allowed (used by the self-loop branch to nudge the loop just
 * outside the shape outline rather than far away).
 */
export function inflateBox(box: RouterBox, pad: number): RouterBox {
  return {
    x: box.x - pad,
    y: box.y - pad,
    cx: box.cx + pad * 2,
    cy: box.cy + pad * 2,
  };
}

/** True when `pt` lies inside `box` (inclusive on the right/bottom). */
export function pointInBox(pt: RouterPoint, box: RouterBox): boolean {
  return (
    pt.x >= box.x &&
    pt.x <= box.x + box.cx &&
    pt.y >= box.y &&
    pt.y <= box.y + box.cy
  );
}

/**
 * True when an axis-aligned segment from `a` to `b` crosses any of
 * the inflated obstacle boxes. We accept either orthogonal segments
 * (the elbow router's primary case) or arbitrary ones (used by the
 * straight / curved sanity checks).
 *
 * The check is conservative: anything that touches the box's interior
 * counts as a hit. Endpoints that fall ON an obstacle's perimeter do
 * NOT count, which matches PowerPoint's "you can land an arrow on
 * a shape's edge" feel — endpoints anchored to a shape would otherwise
 * always count as colliding with their own target.
 */
export function segmentHitsObstacle(
  a: RouterPoint,
  b: RouterPoint,
  obstacles: ReadonlyArray<RouterObstacle>,
  exemptIds?: ReadonlySet<string>
): boolean {
  for (const obs of obstacles) {
    if (exemptIds && exemptIds.has(obs.id)) continue;
    if (segmentIntersectsBox(a, b, obs.box)) return true;
  }
  return false;
}

/**
 * Segment-vs-axis-aligned-box intersection using parametric (Liang-
 * Barsky) clipping. Returns true when the segment ENTERS the open
 * interior of the box. For perfectly axis-aligned segments that lie
 * on a box edge we return false — the segment grazes but doesn't
 * enter, which matches the "endpoint anchored to a shape's edge
 * doesn't count as colliding with that shape" feel.
 */
export function segmentIntersectsBox(
  a: RouterPoint,
  b: RouterPoint,
  box: RouterBox
): boolean {
  const dx = b.x - a.x;
  const dy = b.y - a.y;
  const xMin = box.x;
  const xMax = box.x + box.cx;
  const yMin = box.y;
  const yMax = box.y + box.cy;
  // Standard Liang-Barsky pairs: for each plane (x=xMin, x=xMax,
  // y=yMin, y=yMax) we have p_i = component-of-direction toward the
  // plane and q_i = signed distance from the start point to the
  // plane. The valid t-window is shrunk on each side by max(t for
  // entries) and min(t for exits).
  const ps = [-dx, dx, -dy, dy];
  const qs = [a.x - xMin, xMax - a.x, a.y - yMin, yMax - a.y];
  let tEnter = 0;
  let tExit = 1;
  for (let i = 0; i < 4; i++) {
    const p = ps[i];
    const q = qs[i];
    if (p === 0) {
      // Segment parallel to this plane — only hits if the start point
      // is already inside the corresponding half-plane.
      if (q < 0) return false;
      continue;
    }
    const t = q / p;
    if (p < 0) {
      if (t > tEnter) tEnter = t;
    } else {
      if (t < tExit) tExit = t;
    }
  }
  // Strict interior crossing — endpoints on an edge / corner give
  // tExit - tEnter ≈ 0 and don't count.
  return tExit - tEnter > 1e-6;
}

/**
 * True when at least one segment of the polyline crosses an obstacle.
 * Single-segment paths fall through to `segmentHitsObstacle`.
 */
export function polylineHitsObstacle(
  pts: ReadonlyArray<RouterPoint>,
  obstacles: ReadonlyArray<RouterObstacle>,
  exemptIds?: ReadonlySet<string>
): boolean {
  for (let i = 0; i < pts.length - 1; i++) {
    if (segmentHitsObstacle(pts[i], pts[i + 1], obstacles, exemptIds)) return true;
  }
  return false;
}
