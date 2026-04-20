import type { CommandHandler } from "@officeai/core";
import type { PptxSnapshot, TextParagraph, TextParagraphProperties, TextShape } from "../model/types.js";
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
import type { SetParagraphAlignmentPayload } from "./payloads.js";

const ALGN_MAP: Record<NonNullable<TextParagraphProperties["alignment"]>, string> = {
  left: "l",
  center: "ctr",
  right: "r",
  justify: "just",
};

const VALID_ALIGNMENTS: ReadonlySet<string> = new Set(Object.keys(ALGN_MAP));

/**
 * Set (or clear) a paragraph's horizontal alignment on a `TextShape`.
 *
 * Mirrors PowerPoint's "Align Text" → Left/Center/Right/Justify in the
 * Home ribbon: per-paragraph and shape-wide when no `paragraphs`
 * subset is specified. The serializer prefers `properties.opaqueAttrs`
 * over the typed `alignment` field (so that round-trips of unmodified
 * paragraphs stay byte-faithful), so we have to keep the cached
 * `algn` attribute in sync — otherwise a typed change here would be
 * silently lost on re-emit.
 */
export const setParagraphAlignmentHandler: CommandHandler<SetParagraphAlignmentPayload, PptxSnapshot> = {
  type: "pptx:set-paragraph-alignment",
  apply(snapshot, payload) {
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const { shape, path } = findShapeInSlide(slide, payload.shapeId);
    if (!isTextShape(shape)) {
      throw makeError("not-applicable", `shape is not a text shape`);
    }
    if (payload.alignment !== null && !VALID_ALIGNMENTS.has(payload.alignment)) {
      throw makeError("invalid-payload", `unknown alignment: ${payload.alignment}`);
    }

    const total = shape.txBody.paragraphs.length;
    if (total === 0) {
      throw makeError("not-applicable", `shape has no paragraphs to align`);
    }
    const targets = resolveTargets(payload.paragraphs, total);

    const updatedParagraphs = shape.txBody.paragraphs.map((p, i) =>
      targets.has(i) ? applyAlignment(p, payload.alignment) : p
    );
    const updatedShape: TextShape = {
      ...shape,
      txBody: { ...shape.txBody, paragraphs: updatedParagraphs },
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
        path: ["slides", sIdx, "shapes", ...path, "txBody", "paragraphs"],
        field: "alignment",
        summary:
          payload.alignment === null
            ? `clear-alignment[${[...targets].join(",")}]`
            : `align=${payload.alignment}[${[...targets].join(",")}]`,
      }),
    };
  },
};

function resolveTargets(requested: ReadonlyArray<number> | undefined, total: number): Set<number> {
  if (requested === undefined) {
    const out = new Set<number>();
    for (let i = 0; i < total; i++) out.add(i);
    return out;
  }
  const out = new Set<number>();
  for (const idx of requested) {
    if (!Number.isInteger(idx) || idx < 0 || idx >= total) {
      throw makeError("invalid-payload", `paragraph index ${idx} out of range (0..${total})`);
    }
    out.add(idx);
  }
  if (out.size === 0) {
    throw makeError("invalid-payload", `paragraphs must contain at least one index`);
  }
  return out;
}

function applyAlignment(
  para: TextParagraph,
  alignment: SetParagraphAlignmentPayload["alignment"]
): TextParagraph {
  const props: { -readonly [K in keyof TextParagraphProperties]: TextParagraphProperties[K] } = {
    ...para.properties,
  };
  // Keep `opaqueAttrs.algn` in sync — the serializer reaches for it
  // first when re-emitting a paragraph (see serialize.ts ~1120).
  const opaqueAttrs: Record<string, string> = { ...(para.properties.opaqueAttrs ?? {}) };
  if (alignment === null) {
    delete props.alignment;
    delete opaqueAttrs.algn;
  } else {
    props.alignment = alignment;
    opaqueAttrs.algn = ALGN_MAP[alignment];
  }
  if (Object.keys(opaqueAttrs).length > 0) {
    props.opaqueAttrs = opaqueAttrs;
  } else {
    delete props.opaqueAttrs;
  }
  return { ...para, properties: props };
}
