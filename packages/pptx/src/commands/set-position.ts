import type { CommandHandler } from "@officeai/core";
import type { PptxSnapshot, Shape } from "../model/types.js";
import {
  buildDiff,
  evolveSnapshot,
  findShapeInSlide,
  findSlide,
  makeError,
  replaceShape,
  withSlide,
} from "./helpers.js";
import type { SetPositionPayload } from "./payloads.js";

export const setPositionHandler: CommandHandler<SetPositionPayload, PptxSnapshot> = {
  type: "pptx:set-position",
  apply(snapshot, payload) {
    if (!Number.isFinite(payload.x) || !Number.isFinite(payload.y)) {
      throw makeError("invalid-payload", `x and y must be finite numbers`);
    }
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const { shape, path } = findShapeInSlide(slide, payload.shapeId);
    if (shape.kind === "opaque") {
      throw makeError("not-applicable", "cannot move opaque shape");
    }
    const xEmu = Math.round(payload.x);
    const yEmu = Math.round(payload.y);
    const updated: Shape = { ...shape, position: { xEmu, yEmu } };

    const root = withSlide(snapshot.root, sIdx, (s) => ({
      ...s,
      shapes: replaceShape(s.shapes, path, updated),
    }));

    const next = evolveSnapshot(snapshot, root, { slides: [slide.partPath] });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: shape.id,
        path: ["slides", sIdx, "shapes", ...path],
        field: "position",
        summary: `(${xEmu},${yEmu})`,
      }),
    };
  },
};
