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
  // PowerPoint stores connector endpoints relative to the target's
  // *unrotated* bbox (the anchor index is invariant under rotation),
  // so the wire frame in the file does not bake the rotation in.
  // Apply it here at resolution time so the rendered line meets the
  // shape on its actually-rendered edge — same pivot as
  // `wrapWithRotation` in `renderer/svg/shapes.ts`.
  const rotation =
    "rotation" in target && typeof (target as { rotation?: unknown }).rotation === "number"
      ? (target as { rotation: number }).rotation
      : 0;
  return anchorPoint(box, endpoint.side, endpoint.t, rotation);
}

/**
 * The slide-coordinate location of an anchor on a shape's bounding box.
 *
 * `t` (clamped to [0, 1], default 0.5) interpolates along the picked
 * edge so quarter-points land cleanly: t=0 hits the corner closest to
 * the prior side in the n→e→s→w cycle (left for n/s, top for w/e),
 * t=1 the opposite corner. `center` ignores `t`.
 *
 * `rotationDeg` (default 0) rotates the resulting point clockwise
 * about the box centre so anchors track the shape's rendered edges
 * for rotated shapes. Matches the SVG `rotate(deg cx cy)` transform
 * that `wrapWithRotation` emits, so the visual port dot, the
 * connector endpoint, and the snap target line up exactly.
 */
export function anchorPoint(box: Box, side: ConnectorSide, t?: number, rotationDeg: number = 0): Point {
  const cx = box.x + box.cx / 2;
  const cy = box.y + box.cy / 2;
  const u = clampT(t);
  let local: Point;
  switch (side) {
    case "n":
      local = { x: box.x + box.cx * u, y: box.y };
      break;
    case "s":
      local = { x: box.x + box.cx * u, y: box.y + box.cy };
      break;
    case "w":
      local = { x: box.x, y: box.y + box.cy * u };
      break;
    case "e":
      local = { x: box.x + box.cx, y: box.y + box.cy * u };
      break;
    case "center":
      return { x: cx, y: cy };
    default: {
      const _exhaustive: never = side;
      return { x: cx + (_exhaustive as unknown as number) * 0, y: cy };
    }
  }
  return rotateAroundCenter(local, { x: cx, y: cy }, rotationDeg);
}

/**
 * Rotate a point about a centre by `deg` degrees clockwise (Y-down,
 * matching SVG's `rotate()`). Pure helper used by anchor / snap
 * geometry to keep rotated shapes' connection points on their visible
 * edges.
 */
export function rotateAroundCenter(p: Point, center: Point, deg: number): Point {
  if (deg === 0 || !Number.isFinite(deg)) return p;
  const rad = (deg * Math.PI) / 180;
  const c = Math.cos(rad);
  const s = Math.sin(rad);
  const dx = p.x - center.x;
  const dy = p.y - center.y;
  return { x: center.x + dx * c - dy * s, y: center.y + dx * s + dy * c };
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
