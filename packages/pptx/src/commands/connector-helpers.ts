import type { ConnectorShape, Shape } from "../model/types.js";
import { bboxFromEndpoints, resolveEndpoint } from "../model/connector-geometry.js";
import { collectConnectorsReferencing, replaceShape } from "./helpers.js";

/**
 * Recompute the stored `position`/`size` (the derived bounding box) of
 * every connector on `shapes` whose endpoint resolves through
 * `shapesByCNvPrId`. Returns the (possibly unchanged) shapes array; we
 * splice each updated connector back into the tree using its captured
 * path so group-nested connectors are also handled.
 *
 * Called by `set-position` / `set-size` after the moving shape has been
 * applied so anchored connectors stay glued to the new position. Idempotent
 * when no connectors target the moved shape (returns the same reference).
 */
export function reflowConnectorsForCNvPrId(shapes: ReadonlyArray<Shape>, cNvPrId: number): Shape[] {
  const refs = collectConnectorsReferencing(shapes, cNvPrId);
  if (refs.length === 0) return [...shapes];
  const map = new Map<number, Shape>();
  walkAll(shapes, map);
  let next: Shape[] = [...shapes];
  for (const r of refs) {
    const updated = recomputeConnectorBox(r.shape, map);
    next = replaceShape(next, r.path, updated);
  }
  return next;
}

/**
 * Recompute the bounding box of a single connector from its (possibly
 * anchored) endpoints. When an endpoint can't be resolved (target shape
 * was deleted) we leave the existing box untouched so the connector
 * stays where the user last saw it instead of jumping to (0, 0).
 */
export function recomputeConnectorBox(
  connector: ConnectorShape,
  shapesByCNvPrId: ReadonlyMap<number, Shape>
): ConnectorShape {
  const start = resolveEndpoint(connector.start, shapesByCNvPrId);
  const end = resolveEndpoint(connector.end, shapesByCNvPrId);
  if (!start || !end) return connector;
  const box = bboxFromEndpoints(start, end);
  return {
    ...connector,
    position: { xEmu: box.x, yEmu: box.y },
    size: { cxEmu: box.cx, cyEmu: box.cy },
  };
}

function walkAll(shapes: ReadonlyArray<Shape>, out: Map<number, Shape>): void {
  for (const s of shapes) {
    if (s.cNvPrId > 0) out.set(s.cNvPrId, s);
    if (s.kind === "group") walkAll(s.children, out);
  }
}
