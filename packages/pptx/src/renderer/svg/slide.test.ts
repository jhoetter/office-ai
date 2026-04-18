import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { parsePptx } from "../../parser/parse.js";
import { shapeToSvg } from "./shapes.js";
import { slideToSvgString } from "./slide.js";
import type { TextShape } from "../../model/types.js";

const FIXTURES_DIR = new URL("../../../../../fixtures/pptx/synthetic/", import.meta.url);

async function load(name: string) {
  return parsePptx(await readFile(join(FIXTURES_DIR.pathname, name)));
}

describe("slideToSvgString", () => {
  it("renders a blank slide as a well-formed SVG with white background", async () => {
    const snap = await load("01-blank.pptx");
    const slide = snap.root.slides[0];
    const svg = slideToSvgString(slide, { slideSize: snap.root.slideSize });
    expect(svg.startsWith("<svg")).toBe(true);
    expect(svg.endsWith("</svg>")).toBe(true);
    expect(svg).toContain('viewBox="0 0 ');
    expect(svg).toContain('fill="white"');
  });

  it("renders text shapes via foreignObject HTML so the browser word-wraps long runs", async () => {
    const snap = await load("04-multi-shape.pptx");
    const slide = snap.root.slides[0];
    const svg = slideToSvgString(slide, { slideSize: snap.root.slideSize });
    expect(svg).toContain('class="shape text"');
    expect(svg).toContain("<foreignObject");
    expect(svg).toContain('xmlns="http://www.w3.org/1999/xhtml"');
    expect(svg).toContain("white-space:pre-wrap");
  });

  it("renders pictures as <image> when a media URL is provided", async () => {
    const snap = await load("05-with-image.pptx");
    const slide = snap.root.slides[0];
    const pic = slide.shapes.find((s) => s.kind === "pic");
    expect(pic).toBeDefined();
    if (!pic || pic.kind !== "pic") return;
    const url = "data:image/png;base64,iVBORw0KGgo=";
    const map = new Map([[pic.mediaPartPath, url]]);
    const svg = slideToSvgString(slide, {
      slideSize: snap.root.slideSize,
      mediaUrls: map,
    });
    expect(svg).toContain(`href="${url}"`);
  });

  it("renders a placeholder rect when picture media URL is missing", async () => {
    const snap = await load("05-with-image.pptx");
    const slide = snap.root.slides[0];
    const svg = slideToSvgString(slide, { slideSize: snap.root.slideSize });
    expect(svg).toContain(">image<");
  });

  it("renders a typed bar chart with native SVG bars", async () => {
    const snap = await load("09-with-chart.pptx");
    const slide = snap.root.slides[0];
    const svg = slideToSvgString(slide, {
      slideSize: snap.root.slideSize,
      charts: snap.root.charts,
      theme: snap.root.themeDefault,
    });
    expect(svg).toContain('class="shape chart"');
    // bar / line / area should at least render rect or polyline glyphs.
    const hasBars = svg.includes("<rect") || svg.includes("<polyline") || svg.includes("<polygon");
    expect(hasBars).toBe(true);
  });

  it("resolves <a:schemeClr val='accent1'> via the supplied theme", () => {
    const theme = {
      tx1: "111111",
      bg1: "EEEEEE",
      tx2: "222222",
      bg2: "DDDDDD",
      accent1: "FF0000",
      accent2: "00FF00",
      accent3: "0000FF",
      accent4: "112233",
      accent5: "445566",
      accent6: "778899",
      hlink: "0563C1",
      folHlink: "954F72",
    };
    const shape: TextShape = {
      id: "node-1",
      kind: "text",
      cNvPrId: 1,
      name: "test",
      position: { xEmu: 0, yEmu: 0 },
      size: { cxEmu: 1_000_000, cyEmu: 200_000 },
      nvSpPrTail: [],
      spPrTail: [],
      txBody: {
        paragraphs: [
          {
            id: "node-2",
            properties: {},
            runs: [
              {
                id: "node-3",
                text: "hello",
                properties: {
                  opaqueChildren: [
                    {
                      tag: "a:solidFill",
                      attrs: {},
                      rawAttrs: {},
                      subtree: [
                        {
                          "a:schemeClr": [],
                          ":@": { "@_val": "accent1" },
                        },
                      ],
                    },
                  ],
                },
              },
            ],
          },
        ],
      },
    };
    const svg = shapeToSvg(shape, {
      slideSize: { cxEmu: 9144000, cyEmu: 6858000 },
      theme,
    });
    expect(svg).toContain("color:#FF0000");
  });

  it("keeps font-size constant when the same shape is resized (text reflows, doesn't stretch)", () => {
    const make = (cxEmu: number, cyEmu: number): TextShape => ({
      id: "node-1",
      kind: "text",
      cNvPrId: 1,
      name: "title",
      position: { xEmu: 0, yEmu: 0 },
      size: { cxEmu, cyEmu },
      nvSpPrTail: [],
      spPrTail: [],
      txBody: {
        paragraphs: [
          {
            id: "node-2",
            properties: {},
            runs: [
              {
                id: "node-3",
                text: "Welcome to officeAI",
                properties: { fontSizeHundredths: 4400 },
              },
            ],
          },
        ],
      },
    });
    const ctx = { slideSize: { cxEmu: 9144000, cyEmu: 6858000 } };
    const wide = shapeToSvg(make(7_468_742, 1_389_742), ctx);
    const narrow = shapeToSvg(make(800_000, 600_000), ctx);
    // The run-level font-size MUST be identical across both renders —
    // it's derived purely from `fontSizeHundredths` (44pt → "44pt") and
    // doesn't depend on the geometry. This is the invariant a user
    // notices when they drag a resize handle: glyphs stay the same
    // height, the text just rewraps at the new width — same as
    // PowerPoint and Google Slides. A regression to a CSS scale-stretch
    // path would either change this token or wrap the whole shape in a
    // `transform="scale(...)"` that visually distorts the text.
    const runFontWide = /<span style="[^"]*font-size:(\d+(?:\.\d+)?)pt/.exec(wide);
    const runFontNarrow = /<span style="[^"]*font-size:(\d+(?:\.\d+)?)pt/.exec(narrow);
    expect(runFontWide?.[1]).toBe("44");
    expect(runFontNarrow?.[1]).toBe("44");
    expect(wide).not.toContain('transform="scale');
    expect(narrow).not.toContain('transform="scale');
    // foreignObject viewport scales with the box so the HTML inside
    // gets a smaller wrap width — that's what triggers the reflow.
    const widthWide = /<foreignObject[^>]*width="(\d+(?:\.\d+)?)"/.exec(wide);
    const widthNarrow = /<foreignObject[^>]*width="(\d+(?:\.\d+)?)"/.exec(narrow);
    expect(Number(widthWide?.[1])).toBeGreaterThan(700);
    expect(Number(widthNarrow?.[1])).toBeLessThan(100);
  });

  it("wraps long text inside a narrow text shape (no overflow)", () => {
    const longText = "This is a very long sentence that should wrap automatically across multiple lines inside a narrow text box.";
    const shape: TextShape = {
      id: "node-1",
      kind: "text",
      cNvPrId: 1,
      name: "narrow",
      position: { xEmu: 0, yEmu: 0 },
      size: { cxEmu: 1_500_000, cyEmu: 2_000_000 },
      nvSpPrTail: [],
      spPrTail: [],
      txBody: {
        paragraphs: [
          {
            id: "node-2",
            properties: {},
            runs: [{ id: "node-3", text: longText, properties: {} }],
          },
        ],
      },
    };
    const svg = shapeToSvg(shape, {
      slideSize: { cxEmu: 9144000, cyEmu: 6858000 },
    });
    expect(svg).toContain("<foreignObject");
    expect(svg).toContain(longText);
    expect(svg).toContain("white-space:pre-wrap");
    expect(svg).toContain("word-wrap:break-word");
  });
});
