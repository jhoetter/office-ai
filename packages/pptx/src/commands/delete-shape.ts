import type { CommandHandler } from "@officeai/core";
import type { PptxSnapshot } from "../model/types.js";
import { buildDiff, evolveSnapshot, findShapeInSlide, findSlide, makeError, withSlide } from "./helpers.js";
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
    const shapes = slide.shapes.filter((_, i) => i !== idx);
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
