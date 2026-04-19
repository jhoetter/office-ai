import type { RouterPoint, RouterSide } from "./types.js";

/**
 * Outward unit vector for an anchored side. `center` and `null`
 * collapse to the zero vector — callers detect "no information" and
 * substitute a chord-derived perpendicular.
 */
export function sideUnitVector(side: RouterSide): RouterPoint {
  switch (side) {
    case "n":
      return { x: 0, y: -1 };
    case "s":
      return { x: 0, y: 1 };
    case "e":
      return { x: 1, y: 0 };
    case "w":
      return { x: -1, y: 0 };
    case "center":
    case null:
      return { x: 0, y: 0 };
    default: {
      const _exhaustive: never = side;
      void _exhaustive;
      return { x: 0, y: 0 };
    }
  }
}

/**
 * Pick the perpendicular bow direction for a free curved-connector
 * endpoint. Rotates the chord 90° counter-clockwise so when both
 * endpoints fall through this branch they push the curve to the SAME
 * side (a clean bow rather than an S that crosses the chord). When
 * the chord is a single point we return (0, 1) so the curve has SOME
 * direction to bow into rather than collapsing.
 */
export function chordPerpendicular(sp: RouterPoint, ep: RouterPoint): RouterPoint {
  const dx = ep.x - sp.x;
  const dy = ep.y - sp.y;
  const len = Math.hypot(dx, dy);
  if (len < 1) return { x: 0, y: 1 };
  return { x: -dy / len, y: dx / len };
}

/**
 * Resolve the outward direction the router should treat the endpoint
 * as exiting along. Anchored sides return their side normal; free
 * endpoints fall back to `fallback` (typically `chordPerpendicular`).
 */
export function exitDirection(side: RouterSide, fallback: RouterPoint): RouterPoint {
  const v = sideUnitVector(side);
  if (v.x !== 0 || v.y !== 0) return v;
  return fallback;
}
