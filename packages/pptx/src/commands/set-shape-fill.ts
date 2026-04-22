import type { CommandHandler } from "@officeai/core";
import type { Picture, PptxSnapshot, Shape, TextShape } from "../model/types.js";
import {
  normaliseFillSpec,
  spliceFillIntoSpPr,
  type FillSpec,
} from "../model/fill.js";
import {
  buildDiff,
  evolveSnapshot,
  findShapeInSlide,
  findSlide,
  makeError,
  replaceShape,
  withSlide,
} from "./helpers.js";
import type { SetShapeFillPayload } from "./payloads.js";

/** Shapes whose `<p:spPr>` carries a typed `spPrTail` we can splice fills into. */
type FillCapableShape = TextShape | Picture;

function isFillCapable(shape: Shape): shape is FillCapableShape {
  return shape.kind === "text" || shape.kind === "pic";
}

/**
 * Apply any kind of fill — solid, gradient, pattern, picture or none —
 * to a shape's `<p:spPr>` block. Accepts the legacy `string | null`
 * shorthand for back-compat with callers (and tests) that only knew
 * about solid colours; under the hood everything goes through
 * `FillSpec`.
 */
export const setShapeFillHandler: CommandHandler<SetShapeFillPayload, PptxSnapshot> = {
  type: "pptx:set-shape-fill",
  apply(snapshot, payload) {
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const { shape, path } = findShapeInSlide(slide, payload.shapeId);
    if (!isFillCapable(shape)) {
      throw makeError("not-applicable", `cannot set fill on shape of kind ${shape.kind}`);
    }

    const spec = coerceFillPayload(payload.fill);
    const next: Shape = applyFill(shape, spec);

    const root = withSlide(snapshot.root, sIdx, (s) => ({
      ...s,
      shapes: replaceShape(s.shapes, path, next),
    }));
    const evolved = evolveSnapshot(snapshot, root, { slides: [slide.partPath] });

    return {
      next: evolved,
      diff: buildDiff(snapshot.revision, evolved.revision, {
        kind: "node-updated",
        nodeId: shape.id,
        path: ["slides", sIdx, "shapes", ...path],
        field: "fill",
        summary: summariseFill(spec),
      }),
    };
  },
};

function applyFill<S extends FillCapableShape>(shape: S, spec: FillSpec): S {
  return { ...shape, spPrTail: spliceFillIntoSpPr(shape.spPrTail, spec) };
}

/**
 * Normalise the legacy `string | null` shorthand into a `FillSpec` and
 * validate everything in one place. `null` maps to `noFill` to match
 * the original handler's "transparent" semantics; `string` → solid.
 */
function coerceFillPayload(input: SetShapeFillPayload["fill"]): FillSpec {
  if (input === null) return { type: "none" };
  if (typeof input === "string") {
    try {
      return normaliseFillSpec({ type: "solid", color: input });
    } catch (err) {
      throw makeError("invalid-payload", err instanceof Error ? err.message : String(err));
    }
  }
  try {
    return normaliseFillSpec(input);
  } catch (err) {
    throw makeError("invalid-payload", err instanceof Error ? err.message : String(err));
  }
}

function summariseFill(spec: FillSpec): string {
  switch (spec.type) {
    case "none":
      return "(none)";
    case "solid":
      return `#${spec.color}${spec.alpha !== undefined ? ` @${Math.round(spec.alpha * 100)}%` : ""}`;
    case "gradient":
      return `gradient(${spec.kind}, ${spec.stops.length} stops)`;
    case "pattern":
      return `pattern(${spec.preset})`;
    case "picture":
      return `picture(${spec.embedRelId})`;
    default: {
      const _exhaustive: never = spec;
      void _exhaustive;
      return "(unknown)";
    }
  }
}
