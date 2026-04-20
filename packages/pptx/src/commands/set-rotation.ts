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
import type { SetRotationPayload } from "./payloads.js";

/**
 * Set a shape's rotation in degrees. We normalise into `[0, 360)` and
 * carry `0` as "rotation cleared" (the serializer omits the OOXML
 * `rot` attr in that case), so toggling between rotated and
 * unrotated states round-trips to the same XML PowerPoint emits.
 *
 * `connector`, `group`, and `opaque` shapes are explicitly out of
 * scope: connectors describe orientation through endpoints +
 * `flipH`/`flipV`, groups have their own child-coord transforms, and
 * opaque shapes are unparsed XML we don't introspect. Returning
 * `not-applicable` lets multi-select callers (e.g. the toolbar) skip
 * those shapes without aborting the whole batch.
 */
export const setRotationHandler: CommandHandler<SetRotationPayload, PptxSnapshot> = {
  type: "pptx:set-rotation",
  apply(snapshot, payload) {
    if (!Number.isFinite(payload.degrees)) {
      throw makeError("invalid-payload", `degrees must be a finite number`);
    }
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const { shape, path } = findShapeInSlide(slide, payload.shapeId);
    if (shape.kind === "opaque") {
      throw makeError("not-applicable", "cannot rotate opaque shape");
    }
    if (shape.kind === "connector") {
      throw makeError(
        "not-applicable",
        "rotate is not supported for connector shapes (use endpoints / flip)"
      );
    }
    if (shape.kind === "group") {
      throw makeError("not-applicable", "rotate is not supported for group shapes yet");
    }

    const normalised = ((payload.degrees % 360) + 360) % 360;
    const rounded = Math.round(normalised * 1000) / 1000;
    // Treat 0 as "no rotation" so the serializer drops the attr.
    const updated: Shape =
      rounded === 0
        ? (() => {
            const { rotation: _omit, ...rest } = shape as Shape & { rotation?: number };
            return rest as Shape;
          })()
        : { ...shape, rotation: rounded };

    if ((shape.rotation ?? 0) === rounded) {
      // No-op: short-circuit so we don't bump the revision for nothing.
      return {
        next: snapshot,
        diff: buildDiff(snapshot.revision, snapshot.revision, {
          kind: "node-updated",
          nodeId: shape.id,
          path: ["slides", sIdx, "shapes", ...path],
          field: "rotation",
          summary: `${rounded}°`,
        }),
      };
    }

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
        field: "rotation",
        summary: `${rounded}°`,
      }),
    };
  },
};
