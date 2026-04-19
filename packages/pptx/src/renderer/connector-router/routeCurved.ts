import { chordPerpendicular, exitDirection } from "./exitVector.js";
import { polylineHitsObstacle } from "./obstacles.js";
import type { RouterObstacle, RouterPoint, RouterSide } from "./types.js";

/**
 * Maximum reach (in EMU) for the cubic Bezier control points pulled
 * out from each endpoint. ~1.5 inch — long enough for the curve to
 * read as a real bow without long connectors ballooning off-slide.
 */
const MAX_REACH_EMU = 1_500_000;

/**
 * Cubic Bezier control polygon for a curved connector. Anchored
 * endpoints push the control out along their side normal; free
 * endpoints push along the chord perpendicular so both ends bow to
 * the same side (clean arc, never an S that crosses the chord).
 *
 * If `obstacles` is provided AND the default bow direction puts the
 * curve into a shape, the bow direction flips to the opposite
 * perpendicular for the free endpoints. This is a cheap heuristic —
 * we don't actually intersect the Bezier with shapes, we just check
 * that the chord-midpoint-offset control polygon clears them. Good
 * enough to avoid the "curve dives into a textbox" case the user
 * reported and zero risk of breaking existing rendering.
 */
export function routeCurved(
  sp: RouterPoint,
  ep: RouterPoint,
  startSide: RouterSide,
  endSide: RouterSide,
  options: {
    readonly obstacles?: ReadonlyArray<RouterObstacle>;
    readonly exemptObstacleIds?: ReadonlySet<string>;
  } = {}
): readonly [RouterPoint, RouterPoint, RouterPoint, RouterPoint] {
  const { obstacles = [], exemptObstacleIds } = options;
  const perp = chordPerpendicular(sp, ep);
  const flipped: RouterPoint = { x: -perp.x, y: -perp.y };

  // Try the default bow direction first; flip if it visibly cuts into
  // an obstacle. We compare the simplified control polygon against
  // obstacles rather than the actual Bezier — fast and conservative.
  const candidate = controlPolygon(sp, ep, startSide, endSide, perp);
  if (obstacles.length === 0 || !polylineHitsObstacle(candidate, obstacles, exemptObstacleIds)) {
    return candidate;
  }
  return controlPolygon(sp, ep, startSide, endSide, flipped);
}

function controlPolygon(
  sp: RouterPoint,
  ep: RouterPoint,
  startSide: RouterSide,
  endSide: RouterSide,
  perp: RouterPoint
): readonly [RouterPoint, RouterPoint, RouterPoint, RouterPoint] {
  const dx = ep.x - sp.x;
  const dy = ep.y - sp.y;
  const len = Math.max(1, Math.hypot(dx, dy));
  const sNorm = exitDirection(startSide, perp);
  const eNorm = exitDirection(endSide, perp);
  const reach = Math.min(len * 0.4, MAX_REACH_EMU);
  const c1: RouterPoint = { x: sp.x + sNorm.x * reach, y: sp.y + sNorm.y * reach };
  const c2: RouterPoint = { x: ep.x + eNorm.x * reach, y: ep.y + eNorm.y * reach };
  return [sp, c1, c2, ep] as const;
}

/**
 * SVG `d` attribute for the cubic Bezier defined by a control
 * polygon. `xform` lets the caller pre-scale coordinates (EMU → user
 * units) without the router needing to know about render-space.
 */
export function curvedPathD(
  pts: readonly [RouterPoint, RouterPoint, RouterPoint, RouterPoint],
  xform: (n: number) => number
): string {
  const [a, c1, c2, b] = pts;
  return `M ${xform(a.x)} ${xform(a.y)} C ${xform(c1.x)} ${xform(c1.y)} ${xform(c2.x)} ${xform(c2.y)} ${xform(b.x)} ${xform(b.y)}`;
}
