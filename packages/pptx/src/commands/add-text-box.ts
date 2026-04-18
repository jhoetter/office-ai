import type { CommandHandler } from "@officeai/core";
import type { OpaqueXml, PptxSnapshot, Slide, TextParagraph, TextRun, TextShape } from "../model/types.js";
import { buildDiff, evolveSnapshot, findSlide, makeError, maxCNvPrId } from "./helpers.js";
import type { AddTextBoxPayload } from "./payloads.js";

export const addTextBoxHandler: CommandHandler<AddTextBoxPayload, PptxSnapshot> = {
  type: "pptx:add-text-box",
  apply(snapshot, payload, ctx) {
    if (!Number.isFinite(payload.x) || !Number.isFinite(payload.y)) {
      throw makeError("invalid-payload", "x and y must be finite numbers");
    }
    if (payload.width <= 0 || payload.height <= 0) {
      throw makeError("invalid-payload", "width and height must be > 0");
    }
    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const cNvPrId = maxCNvPrId(slide.shapes) + 1;
    const name = payload.name ?? `Text Box ${cNvPrId}`;

    const run: TextRun = { id: ctx.mintNodeId(), properties: {}, text: payload.text };
    const para: TextParagraph = {
      id: ctx.mintNodeId(),
      properties: {},
      runs: [run],
    };

    const shape: TextShape = {
      kind: "text",
      id: ctx.mintNodeId(),
      cNvPrId,
      name,
      position: { xEmu: Math.round(payload.x), yEmu: Math.round(payload.y) },
      size: { cxEmu: Math.round(payload.width), cyEmu: Math.round(payload.height) },
      nvSpPrTail: defaultNvSpPrTail(),
      spPrTail: defaultSpPrTail(),
      txBody: {
        bodyPrRaw: defaultBodyPr(),
        paragraphs: [para],
      },
    };

    const newSlide: Slide = { ...slide, shapes: [...slide.shapes, shape] };
    const root = {
      ...snapshot.root,
      slides: snapshot.root.slides.map((s, i) => (i === sIdx ? newSlide : s)),
    };
    const next = evolveSnapshot(snapshot, root, { slides: [slide.partPath] });

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-inserted",
        nodeId: shape.id,
        path: ["slides", sIdx, "shapes", newSlide.shapes.length - 1],
        summary: "text-box",
      }),
    };
  },
};

function defaultNvSpPrTail(): OpaqueXml[] {
  return [
    {
      tag: "p:cNvPr",
      attrs: { id: "0", name: "" },
      rawAttrs: { "@_id": "0", "@_name": "" },
      subtree: [],
    },
    {
      tag: "p:cNvSpPr",
      attrs: { txBox: "1" },
      rawAttrs: { "@_txBox": "1" },
      subtree: [],
    },
    { tag: "p:nvPr", attrs: {}, rawAttrs: {}, subtree: [] },
  ];
}

function defaultSpPrTail(): OpaqueXml[] {
  return [
    {
      tag: "a:prstGeom",
      attrs: { prst: "rect" },
      rawAttrs: { "@_prst": "rect" },
      subtree: [{ "a:avLst": [] }],
    },
    { tag: "a:noFill", attrs: {}, rawAttrs: {}, subtree: [] },
  ];
}

function defaultBodyPr(): OpaqueXml {
  return {
    tag: "a:bodyPr",
    attrs: { wrap: "square", rtlCol: "0" },
    rawAttrs: { "@_wrap": "square", "@_rtlCol": "0" },
    subtree: [],
  };
}
