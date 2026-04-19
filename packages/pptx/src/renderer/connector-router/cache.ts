import type { ConnectorType } from "../../model/types.js";
import type { RouteResult, RouterObstacle, RouterPoint, RouterSide } from "./types.js";

/**
 * Tiny insertion-ordered LRU. We don't need a fancy data structure —
 * cache misses just rerun the router (which is cheap for the common
 * heuristic path), and capping the entry count keeps memory bounded
 * across long editing sessions where every drag yields a new key.
 */
const CAPACITY = 256;
const store = new Map<string, RouteResult>();

/**
 * Compute a stable, side-effect-free key for a route call. Anything
 * that can change the routed polyline / Bezier must be reflected here
 * — endpoints, sides, type, waypoints, AND the obstacle set. The
 * obstacle list is hashed as a sorted concatenation of `id|x|y|cx|cy`
 * tuples so order doesn't change the key.
 */
export function routeCacheKey(
  type: ConnectorType,
  start: RouterPoint,
  end: RouterPoint,
  startSide: RouterSide,
  endSide: RouterSide,
  waypoints: ReadonlyArray<number> | undefined,
  obstacles: ReadonlyArray<RouterObstacle> | undefined
): string {
  const wp = waypoints?.map((w) => w.toFixed(0)).join(",") ?? "";
  const obs = obstacles
    ? [...obstacles]
        .map((o) => `${o.id}|${o.box.x}|${o.box.y}|${o.box.cx}|${o.box.cy}`)
        .sort()
        .join(";")
    : "";
  return [
    type,
    start.x.toFixed(0),
    start.y.toFixed(0),
    end.x.toFixed(0),
    end.y.toFixed(0),
    startSide ?? "f",
    endSide ?? "f",
    wp,
    obs,
  ].join("/");
}

export function cacheGet(key: string): RouteResult | undefined {
  const hit = store.get(key);
  if (hit === undefined) return undefined;
  // Touch the entry so it moves to the end of the insertion order
  // (Map iterates in insertion order; re-setting the same key bumps
  // it). This is what makes the eviction below behave LRU-ish.
  store.delete(key);
  store.set(key, hit);
  return hit;
}

export function cachePut(key: string, value: RouteResult): void {
  if (store.has(key)) store.delete(key);
  store.set(key, value);
  if (store.size > CAPACITY) {
    const oldest = store.keys().next().value;
    if (oldest !== undefined) store.delete(oldest);
  }
}

/** Test-only: drop everything. Production code never needs this. */
export function __clearRouteCache(): void {
  store.clear();
}
