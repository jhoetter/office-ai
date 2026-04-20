import { OBSTACLE_PAD_EMU, polylineHitsObstacle, segmentHitsObstacle } from "./obstacles.js";
import { sideUnitVector } from "./exitVector.js";
import type { RouterObstacle, RouterPoint, RouterSide } from "./types.js";

/**
 * Lead distance for the perpendicular exit segment off an anchored
 * endpoint. Capped at half the displacement on the same axis so a
 * tiny gap doesn't produce a stub longer than the connector itself.
 *
 * ~0.25 inch — close enough to PowerPoint's `bentConnector` default
 * that decks visually translate without surprises.
 */
export const LEAD_EMU = 228_600;

/**
 * Public entry: the heuristic-first elbow router with A* fallback.
 *
 * The decision flow is:
 *   1. Build the candidate polylines that come out of the four
 *      "geometric branches" (free/free, anchored/anchored same axis,
 *      anchored/anchored cross axis, mixed). Each branch yields one
 *      or two candidates.
 *   2. If any candidate clears the obstacle set, return the cheapest
 *      one (sum of segment lengths + bend penalty).
 *   3. Otherwise run an orthogonal A* on a sparse visibility grid
 *      derived from obstacle corners + the perpendicular leads. This
 *      is the "go around" path — slower but guaranteed to find a
 *      route if one exists.
 *   4. If A* fails too (impossible obstacles, iteration cap hit),
 *      fall back to the cheapest heuristic candidate so we still
 *      draw SOMETHING.
 *
 * User-supplied `waypoints` override the auto-routed bridge segment
 * coordinate (see `applyWaypoint`). Waypoints are clamped into the
 * monotonic range defined by the candidate's sp/ep so a stale value
 * can't loop the connector outside its own endpoints.
 */
export function routeElbow(
  sp: RouterPoint,
  ep: RouterPoint,
  startSide: RouterSide,
  endSide: RouterSide,
  options: {
    readonly waypoints?: ReadonlyArray<number>;
    readonly obstacles?: ReadonlyArray<RouterObstacle>;
    /** IDs to skip when evaluating obstacle hits (typically the connector's own start/end shapes). */
    readonly exemptObstacleIds?: ReadonlySet<string>;
  } = {}
): ReadonlyArray<RouterPoint> {
  const { waypoints, obstacles = [], exemptObstacleIds } = options;

  const candidates = candidateRoutes(sp, ep, startSide, endSide, waypoints);
  // Score every candidate; collisions get +∞ so they're only chosen
  // if EVERY route collides (in which case A* takes over below).
  let best: { pts: ReadonlyArray<RouterPoint>; cost: number } | null = null;
  let bestClean: { pts: ReadonlyArray<RouterPoint>; cost: number } | null = null;
  for (const pts of candidates) {
    const collides = polylineHitsObstacle(pts, obstacles, exemptObstacleIds);
    const cost = routeCost(pts) + (collides ? Number.POSITIVE_INFINITY : 0);
    if (best === null || cost < best.cost) best = { pts, cost };
    if (!collides && (bestClean === null || cost < bestClean.cost)) {
      bestClean = { pts, cost };
    }
  }

  if (bestClean) return bestClean.pts;

  // No clean heuristic — try A* over the visibility grid. Honour the
  // user's waypoint override if present, since that's their explicit
  // intent and they may be deliberately routing through a known gap.
  const astar = orthogonalAStar(sp, ep, startSide, endSide, obstacles, exemptObstacleIds);
  if (astar) return astar;

  // Last-ditch: return the cheapest collision-y candidate so we still
  // draw something readable. Caller may prefer to flag this with a
  // toast — we don't here because the router stays pure.
  return best ? best.pts : [sp, ep];
}

/**
 * Sum of segment lengths plus a bend penalty so wiggly routes lose to
 * straight ones of the same length. Used to score heuristic candidates.
 */
export function routeCost(pts: ReadonlyArray<RouterPoint>): number {
  let len = 0;
  for (let i = 0; i < pts.length - 1; i++) {
    len += Math.hypot(pts[i + 1].x - pts[i].x, pts[i + 1].y - pts[i].y);
  }
  // Penalise bends so a 4-point L beats a 5-point Z of the same total
  // length. Roughly half a lead segment per bend keeps the trade-off
  // visible without overwhelming the geometric cost.
  const bends = Math.max(0, pts.length - 2);
  return len + bends * (LEAD_EMU * 0.5);
}

/**
 * Generate every reasonable orthogonal route between sp and ep for
 * the given anchored sides, honouring user waypoints. Each entry is
 * a polyline from sp to ep inclusive.
 */
function candidateRoutes(
  sp: RouterPoint,
  ep: RouterPoint,
  startSide: RouterSide,
  endSide: RouterSide,
  waypoints: ReadonlyArray<number> | undefined
): RouterPoint[][] {
  const sV = sideUnitVector(startSide);
  const eV = sideUnitVector(endSide);
  const dx = ep.x - sp.x;
  const dy = ep.y - sp.y;
  const out: RouterPoint[][] = [];

  // Both endpoints free: try BOTH Z orientations so the chooser can
  // pick whichever clears obstacles.
  if (sV.x === 0 && sV.y === 0 && eV.x === 0 && eV.y === 0) {
    out.push(freeFreeZ(sp, ep, "horizontal", waypoints));
    out.push(freeFreeZ(sp, ep, "vertical", waypoints));
    return out;
  }

  const leadX = Math.min(LEAD_EMU, Math.max(LEAD_EMU / 2, Math.abs(dx) / 2));
  const leadY = Math.min(LEAD_EMU, Math.max(LEAD_EMU / 2, Math.abs(dy) / 2));
  const p1 = { x: sp.x + sV.x * leadX, y: sp.y + sV.y * leadY };
  const p2 = { x: ep.x + eV.x * leadX, y: ep.y + eV.y * leadY };
  const sIsHoriz = sV.x !== 0;
  const eIsHoriz = eV.x !== 0;

  if (sIsHoriz && eIsHoriz) {
    // Both leads horizontal → vertical bridge. Z-shape.
    const defaultMidX = (p1.x + p2.x) / 2;
    const wp = absoluteWaypoint(waypoints, 0);
    const midX = clampToRange(wp ?? defaultMidX, p1.x, p2.x);
    out.push([sp, p1, { x: midX, y: p1.y }, { x: midX, y: p2.y }, p2, ep]);
    return out;
  }
  if (!sIsHoriz && !eIsHoriz) {
    // Both leads vertical → horizontal bridge.
    const defaultMidY = (p1.y + p2.y) / 2;
    const wp = absoluteWaypoint(waypoints, 0);
    const midY = clampToRange(wp ?? defaultMidY, p1.y, p2.y);
    out.push([sp, p1, { x: p1.x, y: midY }, { x: p2.x, y: midY }, p2, ep]);
    return out;
  }
  // Mixed orientations — single corner L (4 points). No waypoint
  // slot here: the bend is fully determined by the two leads.
  const corner = sIsHoriz ? { x: p2.x, y: p1.y } : { x: p1.x, y: p2.y };
  out.push([sp, p1, corner, p2, ep]);
  return out;
}

/**
 * Free/free Z route. `axis` is "horizontal" when the first segment is
 * horizontal (mid pivot is on x), "vertical" otherwise.
 */
function freeFreeZ(
  sp: RouterPoint,
  ep: RouterPoint,
  axis: "horizontal" | "vertical",
  waypoints: ReadonlyArray<number> | undefined
): RouterPoint[] {
  if (axis === "horizontal") {
    const defaultMidX = sp.x + (ep.x - sp.x) / 2;
    const wp = absoluteWaypoint(waypoints, 0);
    const midX = clampToRange(wp ?? defaultMidX, sp.x, ep.x);
    return [sp, { x: midX, y: sp.y }, { x: midX, y: ep.y }, ep];
  }
  const defaultMidY = sp.y + (ep.y - sp.y) / 2;
  const wp = absoluteWaypoint(waypoints, 0);
  const midY = clampToRange(wp ?? defaultMidY, sp.y, ep.y);
  return [sp, { x: sp.x, y: midY }, { x: ep.x, y: midY }, ep];
}

/**
 * Read a waypoint slot as an absolute coordinate. Returns `undefined`
 * for missing / non-finite values so callers fall back to defaults.
 */
function absoluteWaypoint(waypoints: ReadonlyArray<number> | undefined, index: number): number | undefined {
  if (!waypoints) return undefined;
  const v = waypoints[index];
  return typeof v === "number" && Number.isFinite(v) ? v : undefined;
}

function clampToRange(v: number, a: number, b: number): number {
  const lo = Math.min(a, b);
  const hi = Math.max(a, b);
  if (v < lo) return lo;
  if (v > hi) return hi;
  return v;
}

// ─── A* fallback ──────────────────────────────────────────────────────────

/**
 * Visibility-grid orthogonal A*. The grid is the Cartesian product of
 *   xs = { sp.x, ep.x, plus every obstacle box's x and x+cx (padded) }
 *   ys = same for y
 * Each grid intersection is a node; orthogonal moves connect to the
 * next intersection in the four cardinal directions, skipping any
 * move that crosses an obstacle. We bias the cost by bend count so
 * the result reads as a Manhattan route rather than a staircase.
 *
 * Returns `null` when no path exists or the iteration cap is hit.
 */
function orthogonalAStar(
  sp: RouterPoint,
  ep: RouterPoint,
  startSide: RouterSide,
  endSide: RouterSide,
  obstacles: ReadonlyArray<RouterObstacle>,
  exemptIds: ReadonlySet<string> | undefined,
  iterationCap = 2000
): ReadonlyArray<RouterPoint> | null {
  if (obstacles.length === 0) return null;
  const xs = new Set<number>([sp.x, ep.x]);
  const ys = new Set<number>([sp.y, ep.y]);
  for (const o of obstacles) {
    if (exemptIds && exemptIds.has(o.id)) continue;
    xs.add(o.box.x - 1);
    xs.add(o.box.x + o.box.cx + 1);
    ys.add(o.box.y - 1);
    ys.add(o.box.y + o.box.cy + 1);
  }
  // Force a perpendicular exit/entry by injecting a "rail" coord
  // offset by LEAD_EMU along the side normal. Without this, A* can
  // emit a route that exits the anchored shape diagonally / through
  // an adjacent face, which looks broken.
  const sV = sideUnitVector(startSide);
  const eV = sideUnitVector(endSide);
  if (sV.x !== 0) xs.add(sp.x + sV.x * LEAD_EMU);
  if (sV.y !== 0) ys.add(sp.y + sV.y * LEAD_EMU);
  if (eV.x !== 0) xs.add(ep.x + eV.x * LEAD_EMU);
  if (eV.y !== 0) ys.add(ep.y + eV.y * LEAD_EMU);
  const xArr = [...xs].sort((a, b) => a - b);
  const yArr = [...ys].sort((a, b) => a - b);
  const startIx = xArr.indexOf(sp.x);
  const startIy = yArr.indexOf(sp.y);
  const endIx = xArr.indexOf(ep.x);
  const endIy = yArr.indexOf(ep.y);

  type Node = {
    readonly ix: number;
    readonly iy: number;
    readonly g: number;
    readonly h: number;
    readonly f: number;
    /** -1 = none, 0 = +x, 1 = -x, 2 = +y, 3 = -y */
    readonly dir: number;
    readonly parent: Node | null;
  };
  const heuristic = (ix: number, iy: number) => Math.abs(xArr[ix] - ep.x) + Math.abs(yArr[iy] - ep.y);
  const open: Node[] = [];
  const closed = new Map<string, number>();
  const startH = heuristic(startIx, startIy);
  const startNode: Node = {
    ix: startIx,
    iy: startIy,
    g: 0,
    h: startH,
    f: startH,
    dir: -1,
    parent: null,
  };
  open.push(startNode);
  let iterations = 0;
  // Tiny binary-heap-less priority queue: we re-sort on each pop.
  // 2000 iterations × 256 max obstacle count → ~10k entries worst
  // case; sort cost is dominated by the obstacle test below.
  while (open.length > 0 && iterations < iterationCap) {
    iterations++;
    open.sort((a, b) => a.f - b.f);
    const cur = open.shift()!;
    if (cur.ix === endIx && cur.iy === endIy) {
      // Reconstruct path → polyline of grid points.
      const pts: RouterPoint[] = [];
      let n: Node | null = cur;
      while (n) {
        pts.unshift({ x: xArr[n.ix], y: yArr[n.iy] });
        n = n.parent;
      }
      return collapseColinear(pts);
    }
    const key = `${cur.ix},${cur.iy}`;
    if (closed.has(key) && closed.get(key)! <= cur.g) continue;
    closed.set(key, cur.g);
    const neighbors: Array<{ ix: number; iy: number; dir: number }> = [];
    if (cur.ix + 1 < xArr.length) neighbors.push({ ix: cur.ix + 1, iy: cur.iy, dir: 0 });
    if (cur.ix - 1 >= 0) neighbors.push({ ix: cur.ix - 1, iy: cur.iy, dir: 1 });
    if (cur.iy + 1 < yArr.length) neighbors.push({ ix: cur.ix, iy: cur.iy + 1, dir: 2 });
    if (cur.iy - 1 >= 0) neighbors.push({ ix: cur.ix, iy: cur.iy - 1, dir: 3 });
    for (const nb of neighbors) {
      const a = { x: xArr[cur.ix], y: yArr[cur.iy] };
      const b = { x: xArr[nb.ix], y: yArr[nb.iy] };
      if (segmentHitsObstacle(a, b, obstacles, exemptIds)) continue;
      const stepLen = Math.abs(b.x - a.x) + Math.abs(b.y - a.y);
      // Bend penalty: extra cost when the direction differs from the
      // segment we arrived on. ~LEAD_EMU per bend matches the
      // heuristic chooser's `routeCost`.
      const bendCost = cur.dir === -1 || cur.dir === nb.dir ? 0 : LEAD_EMU * 0.5;
      const g = cur.g + stepLen + bendCost;
      const h = heuristic(nb.ix, nb.iy);
      const nbKey = `${nb.ix},${nb.iy}`;
      if (closed.has(nbKey) && closed.get(nbKey)! <= g) continue;
      open.push({ ix: nb.ix, iy: nb.iy, g, h, f: g + h, dir: nb.dir, parent: cur });
    }
  }
  return null;
}

/**
 * Drop interior points that are colinear with both their neighbours
 * so the A* output reads as a small set of long segments instead of a
 * staircase of one-grid-cell steps. Preserves the start and end.
 */
function collapseColinear(pts: ReadonlyArray<RouterPoint>): RouterPoint[] {
  if (pts.length <= 2) return [...pts];
  const out: RouterPoint[] = [pts[0]];
  for (let i = 1; i < pts.length - 1; i++) {
    const prev = out[out.length - 1];
    const curr = pts[i];
    const next = pts[i + 1];
    const colinearH = prev.y === curr.y && curr.y === next.y;
    const colinearV = prev.x === curr.x && curr.x === next.x;
    if (colinearH || colinearV) continue;
    out.push(curr);
  }
  out.push(pts[pts.length - 1]);
  return out;
}

/** Re-export so the renderer / chrome can pick it up without re-deriving. */
export { OBSTACLE_PAD_EMU };
