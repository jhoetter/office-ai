import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { parsePptx } from "../../parser/parse.js";
import { slideToSvgString } from "./slide.js";

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
});
