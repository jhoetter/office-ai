import type { CommandHandler } from "@officeai/core";
import type {
  PptxSnapshot,
  TextParagraph,
  TextRun,
  TextShape,
} from "../model/types.js";
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
import type { SetTextPayload } from "./payloads.js";

export const setTextHandler: CommandHandler<SetTextPayload, PptxSnapshot> = {
  type: "pptx:set-text",
  apply(snapshot, payload, ctx) {
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const { shape, path } = findShapeInSlide(slide, payload.shapeId);
    if (!isTextShape(shape)) {
      throw makeError("not-applicable", `shape ${payload.shapeId} is not a text shape`);
    }
    const text = payload.text;
    const lines = text.split("\n");
    const firstPara = shape.txBody.paragraphs[0];
    const firstRun = firstPara?.runs[0];
    const inheritedParaProps = firstPara?.properties ?? {};
    const inheritedRunProps = firstRun?.properties ?? {};

    const newParagraphs: TextParagraph[] = lines.map((line, i) => {
      const run: TextRun = {
        id: ctx.mintNodeId(),
        properties: { ...inheritedRunProps },
        text: line,
      };
      const p: TextParagraph = {
        id: ctx.mintNodeId(),
        properties: { ...inheritedParaProps },
        runs: [run],
        ...(i === 0 && firstPara?.endParaRPrRaw
          ? { endParaRPrRaw: firstPara.endParaRPrRaw }
          : {}),
      };
      return p;
    });

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

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: shape.id,
        path: ["slides", sIdx, "shapes", ...path, "txBody"],
        field: "text",
        summary: `+${JSON.stringify(text)}`,
      }),
    };
  },
};
