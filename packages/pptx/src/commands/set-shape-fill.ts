import type { CommandHandler } from "@officeai/core";
import type { OpaqueXml, Picture, PptxSnapshot, Shape, TextShape } from "../model/types.js";
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
 * Replaces (or removes) the `<a:solidFill>` entry in a shape's `spPrTail`.
 *
 * D13 — used to be TextShape-only, which left the toolbar swatch
 * permanently disabled for inserted pictures (and any future shape
 * subtype that exposes `spPrTail`). Now also operates on `Picture`,
 * matching PowerPoint's behaviour where a picture frame can carry an
 * outline-style background fill.
 */
export const setShapeFillHandler: CommandHandler<SetShapeFillPayload, PptxSnapshot> = {
  type: "pptx:set-shape-fill",
  apply(snapshot, payload) {
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const { shape, path } = findShapeInSlide(slide, payload.shapeId);
    if (!isFillCapable(shape)) {
      throw makeError("not-applicable", `cannot set fill on shape of kind ${shape.kind}`);
    }

    const next: Shape = applyFill(shape, payload.fill);

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
        summary: payload.fill ? `#${normaliseHex(payload.fill)}` : "(none)",
      }),
    };
  },
};

function applyFill<S extends FillCapableShape>(shape: S, fill: string | null): S {
  return { ...shape, spPrTail: spliceFill(shape.spPrTail, fill) };
}

function spliceFill(tail: ReadonlyArray<OpaqueXml>, fill: string | null): ReadonlyArray<OpaqueXml> {
  // Strip every existing fill-related entry first; PowerPoint allows at
  // most one of solidFill / noFill / gradFill / pattFill at the spPr
  // level, so dropping all of them keeps the result well-formed.
  const filtered = tail.filter(
    (c) => c.tag !== "a:solidFill" && c.tag !== "a:noFill" && c.tag !== "a:gradFill" && c.tag !== "a:pattFill"
  );
  const replacement: OpaqueXml = fill
    ? {
        tag: "a:solidFill",
        attrs: {},
        rawAttrs: {},
        subtree: [{ "a:srgbClr": [], ":@": { "@_val": normaliseHex(fill) } }],
      }
    : { tag: "a:noFill", attrs: {}, rawAttrs: {}, subtree: [] };

  // The fill must come AFTER prstGeom (otherwise the rendered shape's
  // path wouldn't pick up the fill in PowerPoint's serializer). Insert
  // immediately after the first prstGeom, otherwise prepend.
  const idx = filtered.findIndex((c) => c.tag === "a:prstGeom");
  return idx >= 0
    ? [...filtered.slice(0, idx + 1), replacement, ...filtered.slice(idx + 1)]
    : [replacement, ...filtered];
}

function normaliseHex(input: string): string {
  const v = input.trim().replace(/^#/, "").toUpperCase();
  if (!/^[0-9A-F]{6}$/.test(v)) {
    throw makeError("invalid-payload", `invalid hex color: ${input}`);
  }
  return v;
}
