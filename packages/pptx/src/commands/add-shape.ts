import type { CommandHandler } from "@officeai/core";
import type { OpaqueXml, PptxSnapshot, Slide, TextParagraph, TextRun, TextShape } from "../model/types.js";
import { buildDiff, evolveSnapshot, findSlide, makeError, maxCNvPrId } from "./helpers.js";
import type { AddShapePayload, ShapePreset } from "./payloads.js";

/**
 * Inserts a typed `TextShape` configured as a decorative shape (rectangle,
 * ellipse, rounded-rect, line, triangle, or arrow). The renderer already
 * draws `prstGeom` + `solidFill` + `ln` from `spPrTail`, so the model layer
 * doesn't need a new shape kind — we just emit a `TextShape` whose tail
 * contains the right opaque XML. An empty `txBody` keeps the shape purely
 * graphical; subsequent `pptx:set-text` / `pptx:format-text` calls can add
 * a label without a separate command.
 */
export const addShapeHandler: CommandHandler<AddShapePayload, PptxSnapshot> = {
  type: "pptx:add-shape",
  apply(snapshot, payload, ctx) {
    if (!Number.isFinite(payload.x) || !Number.isFinite(payload.y)) {
      throw makeError("invalid-payload", "x and y must be finite numbers");
    }
    // Lines may legitimately have a zero extent in one axis (a horizontal
    // line has `cy=0`, a vertical one has `cx=0`); other shapes need both.
    const isLine = payload.preset === "line";
    if (payload.width < 0 || payload.height < 0) {
      throw makeError("invalid-payload", "width and height must be ≥ 0");
    }
    if (!isLine && (payload.width <= 0 || payload.height <= 0)) {
      throw makeError("invalid-payload", "width and height must be > 0");
    }
    if (isLine && payload.width === 0 && payload.height === 0) {
      throw makeError("invalid-payload", "line must have width or height > 0");
    }
    const preset = payload.preset;
    if (!isKnownPreset(preset)) {
      throw makeError("invalid-payload", `unknown shape preset: ${preset}`);
    }

    // Lines never render a fill (we strip `solidFill` for them in
    // `buildSpPrTail`), so don't validate one for that case — the caller
    // shouldn't be forced to invent a placeholder hex.
    const fill = isLine
      ? normaliseHex(payload.fill ?? "000000")
      : normaliseHex(payload.fill ?? defaultFillFor(preset));
    const stroke = payload.stroke ? normaliseHex(payload.stroke) : null;
    const strokeWidthEmu = payload.strokeWidthEmu ?? defaultStrokeWidth(preset);

    const { slide, index: sIdx } = findSlide(snapshot, payload.slideIndex);
    const cNvPrId = maxCNvPrId(slide.shapes) + 1;
    const name = payload.name ?? defaultName(preset, cNvPrId);

    const text = payload.text ?? "";
    const para: TextParagraph = {
      id: ctx.mintNodeId(),
      properties: {},
      runs: text.length > 0 ? [{ id: ctx.mintNodeId(), properties: {}, text } satisfies TextRun] : [],
    };

    const shape: TextShape = {
      kind: "text",
      id: ctx.mintNodeId(),
      cNvPrId,
      name,
      position: { xEmu: Math.round(payload.x), yEmu: Math.round(payload.y) },
      size: { cxEmu: Math.round(payload.width), cyEmu: Math.round(payload.height) },
      nvSpPrTail: defaultNvSpPrTail(),
      spPrTail: buildSpPrTail(preset, fill, stroke, strokeWidthEmu),
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
        summary: `shape:${preset}`,
      }),
    };
  },
};

const KNOWN_PRESETS: ReadonlySet<ShapePreset> = new Set<ShapePreset>([
  "rect",
  "roundRect",
  "ellipse",
  "triangle",
  "rtTriangle",
  "diamond",
  "line",
  "rightArrow",
]);

function isKnownPreset(p: string): p is ShapePreset {
  return KNOWN_PRESETS.has(p as ShapePreset);
}

function defaultFillFor(_preset: ShapePreset): string {
  return "4F81BD"; // PowerPoint's classic "Accent 1"
}

function defaultStrokeWidth(preset: ShapePreset): number {
  // 9525 EMU ≈ 1 px @ 96 DPI; lines need a visible default weight.
  return preset === "line" ? 19050 : 0;
}

function defaultName(preset: ShapePreset, cNvPrId: number): string {
  const human: Record<ShapePreset, string> = {
    rect: "Rectangle",
    roundRect: "Rounded Rectangle",
    ellipse: "Ellipse",
    triangle: "Triangle",
    rtTriangle: "Right Triangle",
    diamond: "Diamond",
    line: "Line",
    rightArrow: "Arrow",
  };
  return `${human[preset]} ${cNvPrId}`;
}

function defaultNvSpPrTail(): OpaqueXml[] {
  return [
    {
      tag: "p:cNvPr",
      attrs: { id: "0", name: "" },
      rawAttrs: { "@_id": "0", "@_name": "" },
      subtree: [],
    },
    { tag: "p:cNvSpPr", attrs: {}, rawAttrs: {}, subtree: [] },
    { tag: "p:nvPr", attrs: {}, rawAttrs: {}, subtree: [] },
  ];
}

function defaultBodyPr(): OpaqueXml {
  return {
    tag: "a:bodyPr",
    attrs: { wrap: "square", rtlCol: "0", anchor: "ctr" },
    rawAttrs: { "@_wrap": "square", "@_rtlCol": "0", "@_anchor": "ctr" },
    subtree: [],
  };
}

function buildSpPrTail(
  preset: ShapePreset,
  fillHex: string,
  strokeHex: string | null,
  strokeWidthEmu: number
): OpaqueXml[] {
  const tail: OpaqueXml[] = [
    {
      tag: "a:prstGeom",
      attrs: { prst: preset },
      rawAttrs: { "@_prst": preset },
      subtree: [{ "a:avLst": [] }],
    },
  ];

  // Lines must rely on the stroke; a filled rect would draw a thin band.
  if (preset !== "line") {
    tail.push({
      tag: "a:solidFill",
      attrs: {},
      rawAttrs: {},
      subtree: [{ "a:srgbClr": [], ":@": { "@_val": fillHex } }],
    });
  }

  if (strokeHex || preset === "line") {
    const color = strokeHex ?? fillHex;
    tail.push({
      tag: "a:ln",
      attrs: { w: String(strokeWidthEmu) },
      rawAttrs: { "@_w": String(strokeWidthEmu) },
      subtree: [
        {
          "a:solidFill": [{ "a:srgbClr": [], ":@": { "@_val": color } }],
        },
      ],
    });
  }

  return tail;
}

function normaliseHex(input: string): string {
  const v = input.trim().replace(/^#/, "").toUpperCase();
  if (!/^[0-9A-F]{6}$/.test(v)) {
    throw makeError("invalid-payload", `invalid hex color: ${input}`);
  }
  return v;
}
