import type { ConnectorType } from "../../model/types.js";
import { cacheGet, cachePut, routeCacheKey } from "./cache.js";
import { routeElbow } from "./routeElbow.js";
import { routeCurved } from "./routeCurved.js";
import type { RouteResult, RouterObstacle, RouterOptions, RouterPoint, RouterSide } from "./types.js";

/**
 * Public entry point for the connector router. Both the SVG renderer
 * (`svg/shapes.ts:connectorToSvg`) and the React chrome
 * (`react/SlideCanvas.tsx:computeRoutePoints`) call into here so the
 * preview, the rendered SVG, and the selection halo all agree.
 *
 * The signature is intentionally minimal: pure geometry in, pure
 * geometry out. No SVG strings, no React, no agent / command
 * machinery. That makes it cheap to test, cheap to call from a
 * tight render loop, and lets us cache aggressively.
 *
 * Side handling:
 *   - `null` / `"center"` → no perpendicular exit; router infers
 *     direction from the chord (free endpoints).
 *   - `n | s | e | w` → exit perpendicular to that side via a short
 *     lead segment.
 */
export function routeConnector(
  type: ConnectorType,
  start: RouterPoint,
  end: RouterPoint,
  startSide: RouterSide,
  endSide: RouterSide,
  options: RouterOptions = {}
): RouteResult {
  const key = routeCacheKey(type, start, end, startSide, endSide, options.waypoints, options.obstacles);
  const hit = cacheGet(key);
  if (hit) return hit;
  const fresh = computeRoute(type, start, end, startSide, endSide, options);
  cachePut(key, fresh);
  return fresh;
}

function computeRoute(
  type: ConnectorType,
  start: RouterPoint,
  end: RouterPoint,
  startSide: RouterSide,
  endSide: RouterSide,
  options: RouterOptions
): RouteResult {
  switch (type) {
    case "straight":
    case "unsupported":
      return { kind: "polyline", points: [start, end] };
    case "curved": {
      const pts = routeCurved(start, end, startSide, endSide, {
        obstacles: options.obstacles,
      });
      return { kind: "cubic", points: pts };
    }
    case "elbow": {
      const pts = routeElbow(start, end, startSide, endSide, {
        waypoints: options.waypoints,
        obstacles: options.obstacles,
      });
      return { kind: "polyline", points: pts };
    }
    default: {
      const _exhaustive: never = type;
      void _exhaustive;
      return { kind: "polyline", points: [start, end] };
    }
  }
}

/**
 * Convenience for chrome callers that always want polyline points,
 * even for curved (where we expose the control polygon as the
 * "route"). Straight and elbow pass through unchanged.
 */
export function routeAsPoints(
  type: ConnectorType,
  start: RouterPoint,
  end: RouterPoint,
  startSide: RouterSide,
  endSide: RouterSide,
  options: RouterOptions = {}
): ReadonlyArray<RouterPoint> {
  const result = routeConnector(type, start, end, startSide, endSide, options);
  return result.points;
}

/** Re-export key shapes so consumers can import everything from one place. */
export type { RouteResult, RouterObstacle, RouterOptions, RouterPoint, RouterSide };
