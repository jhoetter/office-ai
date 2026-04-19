import type { ConnectorEndpoint, GroupShape, Shape } from "../../model/types.js";
import { shapeBoundingBox } from "../layout/shape.js";
import { OBSTACLE_PAD_EMU, inflateBox } from "./obstacles.js";
import type { RouterObstacle } from "./types.js";

/**
 * Walk a slide's shape tree and produce the inflated obstacle list
 * the router consumes. Group children are folded into slide-space
 * coordinates so the router never needs to know about groups.
 *
 * Connector shapes are skipped (they aren't obstacles to other
 * connectors) and any shape whose `cNvPrId` matches one of the IDs
 * in `excludeCNvPrIds` is also skipped — typically the source and
 * target shapes of the connector being routed, so a connector
 * doesn't try to route around the very thing it's anchored to.
 */
export function collectObstacles(
  shapes: ReadonlyArray<Shape>,
  excludeCNvPrIds: ReadonlySet<number>,
  pad: number = OBSTACLE_PAD_EMU
): RouterObstacle[] {
  const out: RouterObstacle[] = [];
  walk(shapes, 0, 0, excludeCNvPrIds, pad, out);
  return out;
}

function walk(
  shapes: ReadonlyArray<Shape>,
  offsetX: number,
  offsetY: number,
  exclude: ReadonlySet<number>,
  pad: number,
  out: RouterObstacle[]
): void {
  for (const sh of shapes) {
    if (sh.kind === "group") {
      const g = sh as GroupShape;
      const childOffsetX = offsetX + (g.position?.xEmu ?? 0);
      const childOffsetY = offsetY + (g.position?.yEmu ?? 0);
      walk(g.children, childOffsetX, childOffsetY, exclude, pad, out);
      continue;
    }
    if (sh.kind === "connector") continue;
    if (exclude.has(sh.cNvPrId)) continue;
    if (sh.cNvPrId <= 0) continue;
    const local = shapeBoundingBox(sh);
    if (!local) continue;
    const inflated = inflateBox(
      { x: local.x + offsetX, y: local.y + offsetY, cx: local.cx, cy: local.cy },
      pad
    );
    out.push({ id: sh.id, box: inflated });
  }
}

/**
 * Helper — pull the `cNvPrId` of an anchored endpoint, or `null` for
 * free endpoints. Used by callers to assemble the `excludeCNvPrIds`
 * set without reaching into the endpoint shape directly.
 */
export function endpointCNvPrId(ep: ConnectorEndpoint): number | null {
  return ep.kind === "anchored" ? ep.targetCNvPrId : null;
}
