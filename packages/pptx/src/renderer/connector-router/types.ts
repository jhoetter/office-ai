/**
 * Public types for the connector router.
 *
 * The router is intentionally agnostic to the SVG layer and the React
 * chrome — it consumes pure geometry (endpoints, anchored sides,
 * obstacles) and returns either a polyline (straight, elbow) or a
 * cubic Bezier control polygon (curved). Both the SVG renderer
 * (`svg/shapes.ts`) and the interactive chrome (`react/SlideCanvas`)
 * call into the same entry point so what the user sees while drawing
 * is byte-for-byte the route that gets committed.
 */

import type { ConnectorType } from "../../model/types.js";

/** A point in slide-EMU space. */
export interface RouterPoint {
  readonly x: number;
  readonly y: number;
}

/** A bounding box in slide-EMU space. */
export interface RouterBox {
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
}

/**
 * Side an anchored endpoint exits from. Free endpoints pass `null`
 * here and the router infers a direction from the chord.
 */
export type RouterSide = "n" | "s" | "e" | "w" | "center" | null;

/**
 * Obstacle the router must route around. Boxes are expected to be
 * pre-inflated by the caller (typically `OBSTACLE_PAD_EMU`).
 *
 * `id` is opaque — the router uses it only for cache invalidation /
 * debugging and never compares it against shape IDs.
 */
export interface RouterObstacle {
  readonly id: string;
  readonly box: RouterBox;
}

/**
 * Optional inputs for the router. Most callers leave these empty and
 * the router falls back to the simple monotonic heuristic.
 */
export interface RouterOptions {
  /**
   * Bend offsets the user has manually set on individual interior
   * segments. Index 0 corresponds to the first interior segment
   * (after the lead stub from `start`). Currently honoured only by
   * the elbow router; straight / curved ignore.
   *
   * Values are ABSOLUTE coordinates: x for vertical bridges,
   * y for horizontal bridges. See `routeElbow` for the clamping rules.
   */
  readonly waypoints?: ReadonlyArray<number>;
  /**
   * Inflated obstacle boxes the router must avoid. Should NOT include
   * the connector's own anchored shape(s) for non-self-loop cases —
   * see `obstacles.ts` for collection logic.
   */
  readonly obstacles?: ReadonlyArray<RouterObstacle>;
}

/**
 * Result of routing. `kind` discriminates how the consumer should
 * draw it:
 *   - "polyline" → straight (2 pts) or elbow (>=2 pts) sequence;
 *     consumers may apply corner rounding for the visible path.
 *   - "cubic" → exactly 4 control points (sp, c1, c2, ep) for a
 *     cubic Bezier. Used by `connectorType: "curved"`.
 */
export type RouteResult =
  | { readonly kind: "polyline"; readonly points: ReadonlyArray<RouterPoint> }
  | {
      readonly kind: "cubic";
      readonly points: readonly [RouterPoint, RouterPoint, RouterPoint, RouterPoint];
    };

/**
 * Stable, side-effect-free routing entry point.
 *
 * @see ./routeConnector.ts for the implementation.
 */
export type RouteConnector = (
  type: ConnectorType,
  start: RouterPoint,
  end: RouterPoint,
  startSide: RouterSide,
  endSide: RouterSide,
  options?: RouterOptions
) => RouteResult;
