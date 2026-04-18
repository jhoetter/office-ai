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

  it("renders text shapes with text content escaped", async () => {
    const snap = await load("04-multi-shape.pptx");
    const slide = snap.root.slides[0];
    const svg = slideToSvgString(slide, { slideSize: snap.root.slideSize });
    expect(svg).toContain('class="shape text"');
    expect(svg).toContain("<text");
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
    expect(svg).toContain('fill="#FF0000"');
  });
});
