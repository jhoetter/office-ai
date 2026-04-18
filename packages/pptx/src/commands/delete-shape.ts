import type { CommandHandler } from "@officeai/core";
import type { ConnectorEndpoint, ConnectorShape, PptxSnapshot, Shape } from "../model/types.js";
import { resolveEndpoint } from "../model/connector-geometry.js";
import {
  buildDiff,
  collectConnectorsReferencing,
  evolveSnapshot,
  findShapeInSlide,
  findSlide,
  makeError,
  replaceShape,
  withSlide,
} from "./helpers.js";
import type { DeleteShapePayload } from "./payloads.js";

/**
 * Removes a shape from the slide. Top-level shapes only — nested children
 * inside a `GroupShape` aren't removable yet (a future "ungroup" command
 * would be the right place to handle that). Opaque shapes are deletable
 * because the renderer would otherwise leave them as ghost placeholders.
 */
export const deleteShapeHandler: CommandHandler<DeleteShapePayload, PptxSnapshot> = {
  type: "pptx:delete-shape",
  apply(snapshot, payload) {
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const { shape, path } = findShapeInSlide(slide, payload.shapeId);
    if (path.length !== 1) {
      throw makeError("not-applicable", "deleting nested group children is not supported");
    }
    const idx = path[0]!;
    // Before removing the shape, resolve every dangling connector
    // endpoint to its current absolute position and convert it to a
    // free endpoint. This way the connector stays where the user last
    // saw it instead of jumping to (0, 0) when the target disappears.
    const detachedShapes = detachConnectorsTargeting(slide.shapes, shape.cNvPrId);
    const shapes = detachedShapes.filter((_, i) => i !== idx);
    const root = withSlide(snapshot.root, sIdx, (s) => ({ ...s, shapes }));
    const next = evolveSnapshot(snapshot, root, { slides: [slide.partPath] });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-deleted",
        nodeId: shape.id,
        path: ["slides", sIdx, "shapes", idx],
        summary: shape.kind,
      }),
    };
  },
};

function detachConnectorsTargeting(shapes: ReadonlyArray<Shape>, cNvPrId: number): Shape[] {
  if (cNvPrId <= 0) return [...shapes];
  const refs = collectConnectorsReferencing(shapes, cNvPrId);
  if (refs.length === 0) return [...shapes];
  const map = new Map<number, Shape>();
  walkShapes(shapes, map);
  let out: Shape[] = [...shapes];
  for (const r of refs) {
    out = replaceShape(out, r.path, detachConnector(r.shape, cNvPrId, map));
  }
  return out;
}

function detachConnector(
  c: ConnectorShape,
  cNvPrId: number,
  map: ReadonlyMap<number, Shape>
): ConnectorShape {
  return {
    ...c,
    start: detachIfMatches(c.start, cNvPrId, map),
    end: detachIfMatches(c.end, cNvPrId, map),
  };
}

function detachIfMatches(
  ep: ConnectorEndpoint,
  cNvPrId: number,
  map: ReadonlyMap<number, Shape>
): ConnectorEndpoint {
  if (ep.kind !== "anchored" || ep.targetCNvPrId !== cNvPrId) return ep;
  const pt = resolveEndpoint(ep, map);
  return { kind: "free", xEmu: pt?.x ?? 0, yEmu: pt?.y ?? 0 };
}

function walkShapes(shapes: ReadonlyArray<Shape>, out: Map<number, Shape>): void {
  for (const s of shapes) {
    if (s.cNvPrId > 0) out.set(s.cNvPrId, s);
    if (s.kind === "group") walkShapes(s.children, out);
  }
}
