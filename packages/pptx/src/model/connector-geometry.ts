/**
 * Pure geometry helpers for `ConnectorShape`. Lives in the model layer
 * so commands (which can't reach into the renderer) and the renderer
 * can both resolve endpoints + bounding boxes consistently.
 */

import type { ConnectorEndpoint, ConnectorSide, Shape } from "./types.js";

export interface Point {
  readonly x: number;
  readonly y: number;
}

export interface Box {
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
}

/**
 * Resolve a connector endpoint to an absolute slide-coordinate point.
 *
 * Anchored endpoints walk back to the target shape's bounding box and
 * pick the side anchor (n/s/e/w/center). When the target can't be
 * found, returns `null` — the renderer falls back to the connector's
 * stored bounding-box corner so a deleted target doesn't make the
 * line vanish entirely.
 */
export function resolveEndpoint(
  endpoint: ConnectorEndpoint,
  shapesByCNvPrId: ReadonlyMap<number, Shape>
): Point | null {
  if (endpoint.kind === "free") {
    return { x: endpoint.xEmu, y: endpoint.yEmu };
  }
  const target = shapesByCNvPrId.get(endpoint.targetCNvPrId);
  if (!target || !target.position || !target.size) return null;
  const box: Box = {
    x: target.position.xEmu,
    y: target.position.yEmu,
    cx: target.size.cxEmu,
    cy: target.size.cyEmu,
  };
  return anchorPoint(box, endpoint.side, endpoint.t);
}

/**
 * The slide-coordinate location of an anchor on a shape's bounding box.
 *
 * `t` (clamped to [0, 1], default 0.5) interpolates along the picked
 * edge so quarter-points land cleanly: t=0 hits the corner closest to
 * the prior side in the n→e→s→w cycle (left for n/s, top for w/e),
 * t=1 the opposite corner. `center` ignores `t`.
 */
export function anchorPoint(box: Box, side: ConnectorSide, t?: number): Point {
  const cx = box.x + box.cx / 2;
  const cy = box.y + box.cy / 2;
  const u = clampT(t);
  switch (side) {
    case "n":
      return { x: box.x + box.cx * u, y: box.y };
    case "s":
      return { x: box.x + box.cx * u, y: box.y + box.cy };
    case "w":
      return { x: box.x, y: box.y + box.cy * u };
    case "e":
      return { x: box.x + box.cx, y: box.y + box.cy * u };
    case "center":
      return { x: cx, y: cy };
    default: {
      const _exhaustive: never = side;
      return { x: cx + (_exhaustive as unknown as number) * 0, y: cy };
    }
  }
}

function clampT(t: number | undefined): number {
  if (t === undefined || !Number.isFinite(t)) return 0.5;
  if (t < 0) return 0;
  if (t > 1) return 1;
  return t;
}

/**
 * Compute the bounding box of a connector from its two resolved
 * endpoint points. The minimum dimension is 1 EMU so the rendered
 * `<a:xfrm>` always has positive `cx`/`cy` (PowerPoint renders zero-
 * extent lines, but other consumers may not).
 */
export function bboxFromEndpoints(start: Point, end: Point): Box {
  const x = Math.min(start.x, end.x);
  const y = Math.min(start.y, end.y);
  const cx = Math.max(1, Math.abs(end.x - start.x));
  const cy = Math.max(1, Math.abs(end.y - start.y));
  return { x, y, cx, cy };
}

/**
 * `<a:xfrm>` for a `<p:cxnSp>` carries `flipH` / `flipV` instead of
 * a negative `cx`/`cy` when the start point is to the right of / below
 * the end point. This helper mirrors PowerPoint's emit behaviour so
 * round-trip stays clean.
 */
export interface ConnectorXfrm {
  readonly box: Box;
  readonly flipH: boolean;
  readonly flipV: boolean;
}

export function connectorXfrm(start: Point, end: Point): ConnectorXfrm {
  return {
    box: bboxFromEndpoints(start, end),
    flipH: start.x > end.x,
    flipV: start.y > end.y,
  };
}
