/**
 * Connector routing engine. Pure-geometry, side-effect-free, cached.
 *
 * Both the SVG renderer and the React chrome import from here so the
 * preview the user sees while drawing is byte-for-byte identical to
 * the route that gets committed.
 */

export { routeConnector, routeAsPoints } from "./routeConnector.js";
export { routeElbow, routeCost, LEAD_EMU } from "./routeElbow.js";
export { routeCurved, curvedPathD } from "./routeCurved.js";
export {
  collectObstacles,
  endpointCNvPrId,
} from "./collectObstacles.js";
export {
  OBSTACLE_PAD_EMU,
  inflateBox,
  pointInBox,
  segmentHitsObstacle,
  segmentIntersectsBox,
  polylineHitsObstacle,
} from "./obstacles.js";
export {
  sideUnitVector,
  chordPerpendicular,
  exitDirection,
} from "./exitVector.js";
export { __clearRouteCache } from "./cache.js";
export type {
  RouteResult,
  RouterBox,
  RouterObstacle,
  RouterOptions,
  RouterPoint,
  RouterSide,
  RouteConnector,
} from "./types.js";
