/**
 * Table tests for the connector routing engine.
 *
 * The router has two visible behaviours we need to lock down:
 *
 *   1. Side-pair geometry: the polyline starts at `sp`, ends at `ep`,
 *      and exits perpendicular to the anchored side(s). We exhaustively
 *      check every (startSide × endSide) pair so a future tweak to the
 *      heuristic can't silently regress one quadrant.
 *
 *   2. Obstacle avoidance: when an obstacle blocks the natural route,
 *      the polyline must NOT intersect the inflated obstacle box. The
 *      A* fallback should kick in and find a path around.
 *
 * Curves and waypoints are exercised separately at the bottom.
 */

import { beforeEach, describe, expect, it } from "vitest";
import {
  __clearRouteCache,
  routeConnector,
  type RouterObstacle,
  type RouterPoint,
  type RouterSide,
} from "./index.js";
import { polylineHitsObstacle } from "./obstacles.js";

const SIDES: ReadonlyArray<RouterSide> = ["n", "s", "e", "w"];

beforeEach(() => __clearRouteCache());

describe("routeConnector — straight", () => {
  it("returns exactly the two endpoints regardless of sides", () => {
    const sp = { x: 0, y: 0 };
    const ep = { x: 1_000_000, y: 500_000 };
    for (const a of SIDES) {
      for (const b of SIDES) {
        const r = routeConnector("straight", sp, ep, a, b);
        expect(r.kind).toBe("polyline");
        expect(r.points).toEqual([sp, ep]);
      }
    }
  });
});

describe("routeConnector — elbow side pairs", () => {
  // For every side pair we check three invariants on the polyline:
  //   - it starts at sp
  //   - it ends at ep
  //   - the first segment moves perpendicular to startSide
  //   - the last segment moves perpendicular to endSide
  const sp: RouterPoint = { x: 1_000_000, y: 1_000_000 };
  const ep: RouterPoint = { x: 5_000_000, y: 3_000_000 };

  function isPerpendicular(side: RouterSide, from: RouterPoint, to: RouterPoint): boolean {
    if (side === "n" || side === "s") return from.x === to.x;
    if (side === "e" || side === "w") return from.y === to.y;
    return true;
  }

  for (const a of SIDES) {
    for (const b of SIDES) {
      it(`exits/enters perpendicular for sides ${a} → ${b}`, () => {
        const r = routeConnector("elbow", sp, ep, a, b);
        expect(r.kind).toBe("polyline");
        const pts = r.points;
        expect(pts[0]).toEqual(sp);
        expect(pts[pts.length - 1]).toEqual(ep);
        expect(pts.length).toBeGreaterThanOrEqual(2);
        expect(isPerpendicular(a, pts[0], pts[1])).toBe(true);
        expect(isPerpendicular(b, pts[pts.length - 2], pts[pts.length - 1])).toBe(true);
      });
    }
  }
});

describe("routeConnector — elbow obstacle avoidance", () => {
  it("routes around an obstacle that sits on the natural straight path", () => {
    // sp on the left, ep on the right, an inflated box sitting between
    // them. The natural Z route passes through the obstacle so the A*
    // fallback should fire and produce a polyline that misses it.
    const sp: RouterPoint = { x: 0, y: 1_000_000 };
    const ep: RouterPoint = { x: 4_000_000, y: 1_000_000 };
    const blocker: RouterObstacle = {
      id: "blocker",
      box: { x: 1_500_000, y: 500_000, cx: 1_000_000, cy: 1_000_000 },
    };
    const r = routeConnector("elbow", sp, ep, "e", "w", { obstacles: [blocker] });
    expect(r.kind).toBe("polyline");
    expect(polylineHitsObstacle(r.points, [blocker])).toBe(false);
  });

  it("ignores obstacles passed via exemptObstacleIds (router doesn't crash on empty list)", () => {
    // Smoke check: empty obstacle list still routes fine.
    const sp: RouterPoint = { x: 0, y: 0 };
    const ep: RouterPoint = { x: 1_000_000, y: 1_000_000 };
    const r = routeConnector("elbow", sp, ep, "e", "w", { obstacles: [] });
    expect(r.kind).toBe("polyline");
    expect(r.points[0]).toEqual(sp);
  });

  it("falls back to a heuristic candidate when A* can't find a clear path", () => {
    // Box wraps both endpoints — there's no clear orthogonal route.
    // Router should still return SOMETHING (cheapest collision-y
    // candidate) rather than throwing.
    const sp: RouterPoint = { x: 100_000, y: 100_000 };
    const ep: RouterPoint = { x: 200_000, y: 200_000 };
    const blocker: RouterObstacle = {
      id: "wrap",
      box: { x: 0, y: 0, cx: 500_000, cy: 500_000 },
    };
    const r = routeConnector("elbow", sp, ep, "e", "w", { obstacles: [blocker] });
    expect(r.kind).toBe("polyline");
    expect(r.points[0]).toEqual(sp);
    expect(r.points[r.points.length - 1]).toEqual(ep);
  });
});

describe("routeConnector — elbow waypoints", () => {
  it("honours an absolute waypoint coordinate for the bridge segment", () => {
    const sp: RouterPoint = { x: 0, y: 0 };
    const ep: RouterPoint = { x: 4_000_000, y: 0 };
    // Both leads horizontal → vertical bridge whose x is the waypoint.
    const r = routeConnector("elbow", sp, ep, "e", "w", { waypoints: [3_000_000] });
    expect(r.kind).toBe("polyline");
    // The bridge segment should sit at x=3_000_000 (clamped to the
    // monotonic range between the lead endpoints, which spans the
    // whole sp→ep chord here).
    const pts = r.points;
    const hasBridgeAtX = pts.some((p) => p.x === 3_000_000);
    expect(hasBridgeAtX).toBe(true);
  });

  it("clamps a waypoint that's outside the lead range to the nearest bound", () => {
    const sp: RouterPoint = { x: 0, y: 0 };
    const ep: RouterPoint = { x: 4_000_000, y: 0 };
    // Pass a waypoint well past ep — should clamp to ep.x - leadX, not
    // wrap the connector outside its own endpoints.
    const r = routeConnector("elbow", sp, ep, "e", "w", { waypoints: [10_000_000] });
    expect(r.kind).toBe("polyline");
    const xs = r.points.map((p) => p.x);
    for (const x of xs) {
      expect(x).toBeLessThanOrEqual(ep.x);
    }
  });
});

describe("routeConnector — curved", () => {
  it("returns a 4-point cubic with control points offset from the chord", () => {
    const sp: RouterPoint = { x: 0, y: 0 };
    const ep: RouterPoint = { x: 4_000_000, y: 0 };
    const r = routeConnector("curved", sp, ep, "e", "w");
    expect(r.kind).toBe("cubic");
    expect(r.points.length).toBe(4);
    expect(r.points[0]).toEqual(sp);
    expect(r.points[3]).toEqual(ep);
    // c1 / c2 must NOT lie on the chord (which would render as a line).
    const c1 = r.points[1];
    const c2 = r.points[2];
    const offChord = c1.y !== 0 || c2.y !== 0 || c1.x !== sp.x || c2.x !== ep.x;
    expect(offChord).toBe(true);
  });

  it("does not crash when both endpoints are free and obstacle-free", () => {
    const sp: RouterPoint = { x: 0, y: 0 };
    const ep: RouterPoint = { x: 1_000_000, y: 1_000_000 };
    const r = routeConnector("curved", sp, ep, null, null);
    expect(r.kind).toBe("cubic");
    expect(r.points.length).toBe(4);
  });
});

describe("routeConnector — caching", () => {
  it("returns the same RouteResult instance for repeated calls with the same inputs", () => {
    const sp: RouterPoint = { x: 0, y: 0 };
    const ep: RouterPoint = { x: 4_000_000, y: 2_000_000 };
    const a = routeConnector("elbow", sp, ep, "e", "w");
    const b = routeConnector("elbow", sp, ep, "e", "w");
    expect(b).toBe(a);
  });
});
