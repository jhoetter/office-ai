import type { CommandHandler } from "@officeai/core";
import type { PptxSnapshot, TextParagraph, TextRun, TextShape } from "../model/types.js";
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
import type { SetTextPayload, SetTextParagraphPatch, SetTextRunPatch } from "./payloads.js";

/**
 * D12 — set-text now accepts either:
 *   • `text` — legacy plain-string commit; collapses to one run/par.
 *   • `paragraphs` — structured replacement that preserves per-run
 *     formatting (bold/italic/colour spans + line breaks).
 *
 * The structured form is what the contenteditable overlay should
 * commit on blur; the plain-text form remains for the agent / CLI /
 * tests that want a one-line shape value.
 */
export const setTextHandler: CommandHandler<SetTextPayload, PptxSnapshot> = {
  type: "pptx:set-text",
  apply(snapshot, payload, ctx) {
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const { shape, path } = findShapeInSlide(slide, payload.shapeId);
    if (!isTextShape(shape)) {
      throw makeError("not-applicable", `shape ${payload.shapeId} is not a text shape`);
    }

    const newParagraphs = payload.paragraphs
      ? buildFromParagraphs(payload.paragraphs, shape, ctx)
      : buildFromText(payload.text ?? "", shape, ctx);

    const updated: TextShape = {
      ...shape,
      txBody: {
        ...shape.txBody,
        paragraphs: newParagraphs,
      },
    };

    const root = withSlide(snapshot.root, sIdx, (s) => ({
      ...s,
      shapes: replaceShape(s.shapes, path, updated),
    }));

    const next = evolveSnapshot(snapshot, root, { slides: [slide.partPath] });

    const summaryText = payload.paragraphs ? paragraphsPlain(newParagraphs) : (payload.text ?? "");

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: shape.id,
        path: ["slides", sIdx, "shapes", ...path, "txBody"],
        field: "text",
        summary: `+${JSON.stringify(summaryText)}`,
      }),
    };
  },
};

function buildFromText(
  text: string,
  shape: TextShape,
  ctx: { mintNodeId: () => import("@officeai/core").NodeId }
): TextParagraph[] {
  const lines = text.split("\n");
  const firstPara = shape.txBody.paragraphs[0];
  const firstRun = firstPara?.runs[0];
  const inheritedParaProps = firstPara?.properties ?? {};
  const inheritedRunProps = firstRun?.properties ?? {};

  return lines.map((line, i) => {
    const run: TextRun = {
      id: ctx.mintNodeId(),
      properties: { ...inheritedRunProps },
      text: line,
    };
    const p: TextParagraph = {
      id: ctx.mintNodeId(),
      properties: { ...inheritedParaProps },
      runs: [run],
      ...(i === 0 && firstPara?.endParaRPrRaw ? { endParaRPrRaw: firstPara.endParaRPrRaw } : {}),
    };
    return p;
  });
}

function buildFromParagraphs(
  patches: ReadonlyArray<SetTextParagraphPatch>,
  shape: TextShape,
  ctx: { mintNodeId: () => import("@officeai/core").NodeId }
): TextParagraph[] {
  const existing = shape.txBody.paragraphs;
  const fallbackParaProps = existing[0]?.properties ?? {};
  const fallbackRunProps = existing[0]?.runs[0]?.properties ?? {};

  return patches.map((pp, pi) => {
    const original = existing[pi];
    const paraProps = pp.properties ?? original?.properties ?? fallbackParaProps;
    const runs: TextRun[] = pp.runs.map((rp) => buildRun(rp, original, fallbackRunProps, ctx));
    const p: TextParagraph = {
      id: ctx.mintNodeId(),
      properties: { ...paraProps },
      runs,
      ...(original?.endParaRPrRaw ? { endParaRPrRaw: original.endParaRPrRaw } : {}),
    };
    return p;
  });
}

function buildRun(
  rp: SetTextRunPatch,
  originalParagraph: TextParagraph | undefined,
  fallbackRunProps: import("../model/types.js").TextRunProperties,
  ctx: { mintNodeId: () => import("@officeai/core").NodeId }
): TextRun {
  const inherited =
    rp.properties ??
    (rp.inheritFromRun !== undefined ? originalParagraph?.runs[rp.inheritFromRun]?.properties : undefined) ??
    originalParagraph?.runs[0]?.properties ??
    fallbackRunProps;
  return {
    id: ctx.mintNodeId(),
    properties: { ...inherited },
    text: rp.text,
    ...(rp.isLineBreak ? { isLineBreak: true } : {}),
  };
}

function paragraphsPlain(paragraphs: ReadonlyArray<TextParagraph>): string {
  return paragraphs.map((p) => p.runs.map((r) => (r.isLineBreak ? "\n" : r.text)).join("")).join("\n");
}
