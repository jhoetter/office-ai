import type { CommandHandler } from "@officeai/core";
import type { PptxSnapshot, Shape } from "../model/types.js";
import { reflowConnectorsForCNvPrId } from "./connector-helpers.js";
import {
  buildDiff,
  evolveSnapshot,
  findShapeInSlide,
  findSlide,
  makeError,
  replaceShape,
  withSlide,
} from "./helpers.js";
import type { SetSizePayload } from "./payloads.js";

export const setSizeHandler: CommandHandler<SetSizePayload, PptxSnapshot> = {
  type: "pptx:set-size",
  apply(snapshot, payload) {
    if (!Number.isFinite(payload.width) || !Number.isFinite(payload.height)) {
      throw makeError("invalid-payload", `width and height must be finite numbers`);
    }
    if (payload.width <= 0 || payload.height <= 0) {
      throw makeError("invalid-payload", `width and height must be > 0`);
    }
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const { shape, path } = findShapeInSlide(slide, payload.shapeId);
    if (shape.kind === "opaque") {
      throw makeError("not-applicable", "cannot resize opaque shape");
    }
    const cxEmu = Math.round(payload.width);
    const cyEmu = Math.round(payload.height);
    const updated: Shape = { ...shape, size: { cxEmu, cyEmu } };

    const root = withSlide(snapshot.root, sIdx, (s) => {
      const resizedShapes = replaceShape(s.shapes, path, updated);
      const reflowed = reflowConnectorsForCNvPrId(resizedShapes, shape.cNvPrId);
      return { ...s, shapes: reflowed };
    });

    const next = evolveSnapshot(snapshot, root, { slides: [slide.partPath] });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: shape.id,
        path: ["slides", sIdx, "shapes", ...path],
        field: "size",
        summary: `${cxEmu}×${cyEmu}`,
      }),
    };
  },
};
