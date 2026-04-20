import type { CommandHandler } from "@officeai/core";
import type { OpaqueXml, PptxSnapshot, TextShape } from "../model/types.js";
import {
  buildDiff,
  evolveSnapshot,
  findShapeInSlide,
  findSlide,
  isTextShape,
  makeError,
  replaceShape,
  withSlide,
} from "./helpers.js";
import type { SetTextAnchorPayload, TextAnchor } from "./payloads.js";

const ANCHOR_MAP: Record<TextAnchor, string> = {
  top: "t",
  middle: "ctr",
  bottom: "b",
};

const VALID_ANCHORS: ReadonlySet<string> = new Set(Object.keys(ANCHOR_MAP));

/**
 * Set (or clear) a text shape's vertical anchor (`<a:bodyPr anchor>`).
 *
 * PowerPoint's "Align Text" → Top/Middle/Bottom is shape-wide, so the
 * payload doesn't take a paragraph subset. The anchor lives on
 * `<a:bodyPr>`, which the model preserves opaquely; this handler
 * mutates the cached attrs in place (or synthesises a default
 * `bodyPr` when one wasn't on disk) so the renderer's `readBodyAnchor`
 * picks up the change without anyone needing to round-trip the XML.
 */
export const setTextAnchorHandler: CommandHandler<SetTextAnchorPayload, PptxSnapshot> = {
  type: "pptx:set-text-anchor",
  apply(snapshot, payload) {
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const { shape, path } = findShapeInSlide(slide, payload.shapeId);
    if (!isTextShape(shape)) {
      throw makeError("not-applicable", `shape is not a text shape`);
    }
    if (payload.anchor !== null && !VALID_ANCHORS.has(payload.anchor)) {
      throw makeError("invalid-payload", `unknown anchor: ${payload.anchor}`);
    }

    const nextBodyPr = applyAnchor(shape.txBody.bodyPrRaw, payload.anchor);
    const updatedShape: TextShape = {
      ...shape,
      txBody: { ...shape.txBody, bodyPrRaw: nextBodyPr },
    };

    const root = withSlide(snapshot.root, sIdx, (s) => ({
      ...s,
      shapes: replaceShape(s.shapes, path, updatedShape),
    }));
    const next = evolveSnapshot(snapshot, root, { slides: [slide.partPath] });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: shape.id,
        path: ["slides", sIdx, "shapes", ...path, "txBody", "bodyPrRaw"],
        field: "anchor",
        summary: payload.anchor === null ? "clear-anchor" : `anchor=${payload.anchor}`,
      }),
    };
  },
};

function applyAnchor(current: OpaqueXml | undefined, anchor: TextAnchor | null): OpaqueXml | undefined {
  const base: OpaqueXml = current ?? defaultBodyPr();
  const attrs: Record<string, string> = { ...base.attrs };
  const rawAttrs: Record<string, string> = { ...base.rawAttrs };
  if (anchor === null) {
    delete attrs.anchor;
    delete rawAttrs["@_anchor"];
  } else {
    attrs.anchor = ANCHOR_MAP[anchor];
    rawAttrs["@_anchor"] = ANCHOR_MAP[anchor];
  }
  // If the source had no bodyPr and we're asked to clear, leave it
  // missing — there's nothing to encode and PowerPoint defaults to
  // "top" anyway.
  if (anchor === null && current === undefined) return undefined;
  return { ...base, attrs, rawAttrs };
}

function defaultBodyPr(): OpaqueXml {
  return {
    tag: "a:bodyPr",
    attrs: { wrap: "square", rtlCol: "0" },
    rawAttrs: { "@_wrap": "square", "@_rtlCol": "0" },
    subtree: [],
  };
}
