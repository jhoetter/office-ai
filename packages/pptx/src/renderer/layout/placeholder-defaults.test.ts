import { describe, expect, it } from "vitest";

import type {
  OpaqueXml,
  PlaceholderSpec,
  Shape,
  SlideLayout,
  TextBody,
  TextShape,
} from "../../model/types.js";
import {
  findLayoutPlaceholder,
  resolvePlaceholderTextDefaults,
  resolvedShapeBoundingBox,
} from "./placeholder-defaults.js";

const EMPTY_XML: OpaqueXml = { tag: "", attrs: {}, rawAttrs: {}, subtree: [] };

function emptyTextBody(anchor?: "ctr" | "t" | "b"): TextBody {
  return {
    bodyPrRaw: anchor
      ? { tag: "a:bodyPr", attrs: { anchor }, rawAttrs: { "@_anchor": anchor }, subtree: [] }
      : undefined,
    paragraphs: [
      {
        id: "p0",
        properties: {},
        runs: [{ id: "r0", properties: {}, text: "" }],
      },
    ],
  };
}

function makeShape(opts: {
  id?: string;
  type?: string;
  idx?: number;
  position?: { xEmu: number; yEmu: number };
  size?: { cxEmu: number; cyEmu: number };
  body?: TextBody;
}): TextShape {
  return {
    id: opts.id ?? "s1",
    cNvPrId: 1,
    name: "ph",
    kind: "text",
    placeholder:
      opts.type !== undefined || opts.idx !== undefined
        ? { type: opts.type, idx: opts.idx }
        : undefined,
    position: opts.position,
    size: opts.size,
    txBody: opts.body ?? emptyTextBody(),
    nvSpPrTail: [],
    spPrTail: [],
  };
}

function makeLayout(placeholders: ReadonlyArray<PlaceholderSpec>): SlideLayout {
  return {
    partPath: "ppt/slideLayouts/slideLayout1.xml",
    kind: "title",
    name: "Title Slide",
    placeholders,
    raw: EMPTY_XML,
  };
}

describe("resolvePlaceholderTextDefaults", () => {
  it("returns 36pt left-anchored Calibri for an empty title placeholder", () => {
    const shape = makeShape({ type: "title", idx: 0 });
    const defaults = resolvePlaceholderTextDefaults(shape, undefined);
    expect(defaults.fontSizePt).toBe(36);
    expect(defaults.align).toBe("left");
    expect(defaults.anchor).toBe("middle");
    expect(defaults.fontFamily).toContain("Calibri");
  });

  it("returns 40pt center-anchored for a centered-title placeholder", () => {
    const shape = makeShape({ type: "ctrTitle", idx: 0 });
    const defaults = resolvePlaceholderTextDefaults(shape, undefined);
    expect(defaults.fontSizePt).toBe(40);
    expect(defaults.align).toBe("center");
    expect(defaults.anchor).toBe("middle");
  });

  it("returns 18pt top-anchored for a body placeholder", () => {
    const shape = makeShape({ type: "body", idx: 1 });
    const defaults = resolvePlaceholderTextDefaults(shape, undefined);
    expect(defaults.fontSizePt).toBe(18);
    expect(defaults.align).toBe("left");
    expect(defaults.anchor).toBe("top");
  });

  it("falls back to body defaults for unknown placeholder types", () => {
    const shape = makeShape({ type: "weirdType" });
    const defaults = resolvePlaceholderTextDefaults(shape, undefined);
    expect(defaults.fontSizePt).toBe(18);
    expect(defaults.anchor).toBe("top");
  });

  it("respects an explicit run fontSizeHundredths over the type default", () => {
    const body: TextBody = {
      paragraphs: [
        {
          id: "p0",
          properties: {},
          runs: [
            {
              id: "r0",
              properties: { fontSizeHundredths: 4400, fontFamily: "Arial" },
              text: "Hello",
            },
          ],
        },
      ],
    };
    const shape = makeShape({ type: "body", body });
    const defaults = resolvePlaceholderTextDefaults(shape, undefined);
    expect(defaults.fontSizePt).toBe(44);
    expect(defaults.fontFamily).toContain("Arial");
  });

  it("uses the bodyPr anchor attribute over the type default", () => {
    const shape = makeShape({ type: "body", body: emptyTextBody("ctr") });
    expect(resolvePlaceholderTextDefaults(shape, undefined).anchor).toBe("middle");
    const bottom = makeShape({ type: "body", body: emptyTextBody("b") });
    expect(resolvePlaceholderTextDefaults(bottom, undefined).anchor).toBe("bottom");
  });

  it("uses the paragraph alignment over the type default", () => {
    const body: TextBody = {
      paragraphs: [
        {
          id: "p0",
          properties: { alignment: "center" },
          runs: [{ id: "r0", properties: {}, text: "x" }],
        },
      ],
    };
    const shape = makeShape({ type: "body", body });
    expect(resolvePlaceholderTextDefaults(shape, undefined).align).toBe("center");
  });

  it("returns bold weight when the first run is bold", () => {
    const body: TextBody = {
      paragraphs: [
        {
          id: "p0",
          properties: {},
          runs: [{ id: "r0", properties: { bold: true }, text: "x" }],
        },
      ],
    };
    const shape = makeShape({ type: "body", body });
    expect(resolvePlaceholderTextDefaults(shape, undefined).fontWeight).toBe(700);
  });
});

describe("findLayoutPlaceholder", () => {
  const layout = makeLayout([
    { type: "title", idx: 0 },
    { type: "body", idx: 1 },
    { type: "body", idx: 2 },
  ]);

  it("matches by (type, idx) first", () => {
    const shape = makeShape({ type: "body", idx: 2 });
    expect(findLayoutPlaceholder(shape, layout)?.idx).toBe(2);
  });

  it("falls back to idx-only match", () => {
    const shape = makeShape({ type: "wrong", idx: 1 });
    expect(findLayoutPlaceholder(shape, layout)?.type).toBe("body");
  });

  it("falls back to type-only match", () => {
    const shape = makeShape({ type: "title", idx: 99 });
    expect(findLayoutPlaceholder(shape, layout)?.type).toBe("title");
  });

  it("returns undefined when no layout is supplied", () => {
    const shape = makeShape({ type: "title", idx: 0 });
    expect(findLayoutPlaceholder(shape, undefined)).toBeUndefined();
  });

  it("returns undefined for a non-placeholder shape", () => {
    const shape: Shape = { ...makeShape({}), placeholder: undefined };
    expect(findLayoutPlaceholder(shape, layout)).toBeUndefined();
  });
});

describe("resolvedShapeBoundingBox", () => {
  it("uses the shape's own geometry when present", () => {
    const shape = makeShape({
      type: "title",
      idx: 0,
      position: { xEmu: 100, yEmu: 200 },
      size: { cxEmu: 300, cyEmu: 400 },
    });
    expect(resolvedShapeBoundingBox(shape, undefined)).toEqual({
      x: 100,
      y: 200,
      cx: 300,
      cy: 400,
    });
  });

  it("falls back to the layout placeholder's geometry", () => {
    const layout = makeLayout([
      {
        type: "title",
        idx: 0,
        position: { xEmu: 914400, yEmu: 1828800 },
        size: { cxEmu: 7315200, cyEmu: 1143000 },
      },
    ]);
    const shape = makeShape({ type: "title", idx: 0 });
    expect(resolvedShapeBoundingBox(shape, layout)).toEqual({
      x: 914400,
      y: 1828800,
      cx: 7315200,
      cy: 1143000,
    });
  });

  it("returns null when neither the shape nor the layout supply geometry", () => {
    const layout = makeLayout([{ type: "title", idx: 0 }]);
    const shape = makeShape({ type: "title", idx: 0 });
    expect(resolvedShapeBoundingBox(shape, layout)).toBeNull();
  });
});
