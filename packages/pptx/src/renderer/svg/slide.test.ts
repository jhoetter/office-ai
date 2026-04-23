import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { parsePptx } from "../../parser/parse.js";
import { shapeToSvg } from "./shapes.js";
import { slideToSvgString } from "./slide.js";
import type { ConnectorShape, TextShape } from "../../model/types.js";

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

  it("present/export SVG keeps the slide-only viewBox (off-slide content is cropped)", async () => {
    // The interactive editor canvas uses `slideStageViewBox` so the user
    // can see shapes that live in the scratch margin, but the artefact
    // we hand to present mode and the .pptx writer must still crop to
    // the slide bounds — exactly like real PowerPoint, where the
    // delivered slideshow never shows scratch-area objects.
    const snap = await load("01-blank.pptx");
    const slide = snap.root.slides[0];
    const svg = slideToSvgString(slide, { slideSize: snap.root.slideSize });
    const match = /viewBox="(-?\d+(?:\.\d+)?) (-?\d+(?:\.\d+)?) /.exec(svg);
    expect(match).not.toBeNull();
    const x = Number(match?.[1]);
    const y = Number(match?.[2]);
    expect(x).toBe(0);
    expect(y).toBe(0);
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
                text: "Welcome to office-ai",
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

  it("paints a dashed ghost outline + prompt label for an empty text placeholder (no real text)", () => {
    // Mirrors what `clonePlaceholdersIntoSlide` produces for a fresh
    // "Title and Content" layout: a body placeholder with one empty
    // paragraph and zero runs. Without ghost rendering the SVG would
    // be a transparent rect — visually invisible — which is exactly
    // the "blank slide" bug we're fixing here.
    const shape: TextShape = {
      id: "ph-body",
      kind: "text",
      cNvPrId: 2,
      name: "Body 1",
      position: { xEmu: 500_000, yEmu: 1_500_000 },
      size: { cxEmu: 8_000_000, cyEmu: 4_000_000 },
      nvSpPrTail: [],
      spPrTail: [],
      placeholder: { type: "body", idx: 1 },
      txBody: {
        paragraphs: [{ id: "p-0", properties: {}, runs: [] }],
      },
    };
    const svg = shapeToSvg(shape, {
      slideSize: { cxEmu: 9144000, cyEmu: 6858000 },
    });
    expect(svg).toContain('class="placeholder-hint"');
    expect(svg).toContain("stroke-dasharray=");
    expect(svg).toContain("Click to add text");
    // The hint must never absorb pointer events — selection / drag /
    // resize hit-testing depends on the underlying transparent rect
    // staying clickable. If this regresses the user can no longer
    // grab an empty placeholder by clicking inside it.
    expect(svg).toContain('pointer-events="none"');
  });

  it("renders a picture-icon glyph for an empty pic placeholder", () => {
    const shape: TextShape = {
      id: "ph-pic",
      kind: "text",
      cNvPrId: 3,
      name: "Picture Placeholder 1",
      position: { xEmu: 500_000, yEmu: 1_500_000 },
      size: { cxEmu: 4_000_000, cyEmu: 3_000_000 },
      nvSpPrTail: [],
      spPrTail: [],
      placeholder: { type: "pic", idx: 2 },
      txBody: {
        paragraphs: [{ id: "p-0", properties: {}, runs: [] }],
      },
    };
    const svg = shapeToSvg(shape, {
      slideSize: { cxEmu: 9144000, cyEmu: 6858000 },
    });
    expect(svg).toContain("Click to add picture");
    // Mountain-and-sun glyph: a polyline (the mountain) + a circle
    // (the sun) inside the dashed frame. Asserting both ensures the
    // icon path actually fires — a stray refactor that drops the
    // helper would still satisfy "Click to add picture".
    expect(svg).toContain("<polyline");
    expect(svg).toContain("<circle");
  });

  it("does NOT paint a ghost when the placeholder already has text", () => {
    // Once the user types into a placeholder the ghost UI must
    // disappear, otherwise the slide reads as cluttered (real title
    // text *plus* a dashed "Click to add title" overlay).
    const shape: TextShape = {
      id: "ph-title",
      kind: "text",
      cNvPrId: 4,
      name: "Title 1",
      position: { xEmu: 500_000, yEmu: 500_000 },
      size: { cxEmu: 8_000_000, cyEmu: 1_000_000 },
      nvSpPrTail: [],
      spPrTail: [],
      placeholder: { type: "title", idx: 0 },
      txBody: {
        paragraphs: [
          {
            id: "p-0",
            properties: {},
            runs: [{ id: "r-0", text: "Real title", properties: {} }],
          },
        ],
      },
    };
    const svg = shapeToSvg(shape, {
      slideSize: { cxEmu: 9144000, cyEmu: 6858000 },
    });
    expect(svg).not.toContain('class="placeholder-hint"');
    expect(svg).not.toContain("Click to add title");
    expect(svg).toContain("Real title");
  });

  it("respects renderPlaceholderHints=false (suppresses the ghost UI)", () => {
    // Export-style previews opt out of the authoring affordances so
    // the rendered slide matches what's in the saved .pptx — empty
    // placeholders should look invisible, not dashed-and-prompty.
    const shape: TextShape = {
      id: "ph-body",
      kind: "text",
      cNvPrId: 5,
      name: "Body 1",
      position: { xEmu: 500_000, yEmu: 1_500_000 },
      size: { cxEmu: 8_000_000, cyEmu: 4_000_000 },
      nvSpPrTail: [],
      spPrTail: [],
      placeholder: { type: "body", idx: 1 },
      txBody: { paragraphs: [{ id: "p-0", properties: {}, runs: [] }] },
    };
    const svg = shapeToSvg(shape, {
      slideSize: { cxEmu: 9144000, cyEmu: 6858000 },
      renderPlaceholderHints: false,
    });
    expect(svg).not.toContain('class="placeholder-hint"');
    expect(svg).not.toContain("Click to add text");
  });

  it("wraps long text inside a narrow text shape (no overflow)", () => {
    const longText =
      "This is a very long sentence that should wrap automatically across multiple lines inside a narrow text box.";
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

  it("renders a curved connector as a cubic Bezier with offset control points (not a degenerate straight line)", () => {
    const shape: ConnectorShape = {
      id: "cxn-1",
      kind: "connector",
      cNvPrId: 10,
      name: "Curved 1",
      position: { xEmu: 0, yEmu: 0 },
      size: { cxEmu: 4_000_000, cyEmu: 0 },
      nvCxnSpPrTail: [],
      spPrTail: [],
      connectorType: "curved",
      start: { kind: "free", xEmu: 0, yEmu: 0 },
      end: { kind: "free", xEmu: 4_000_000, yEmu: 0 },
    };
    const svg = shapeToSvg(shape, { slideSize: { cxEmu: 9_144_000, cyEmu: 6_858_000 } });
    // Cubic Bezier path is required — a quadratic with control on the
    // chord midpoint (the previous bug) renders identically to a
    // straight line, so users couldn't tell "curved" from "straight".
    expect(svg).toMatch(/<path d="M 0 0 C \S+ \S+ \S+ \S+ /);
    // Endpoints must remain exactly on sp/ep so anchored markers and
    // selection dots line up with the visible tip.
    expect(svg).toContain(' 419.95 0"');
  });

  it("uses placeholder-typed font size in the foreignObject when runs carry no explicit size", () => {
    // Synthesise a title placeholder with a single user-typed run that
    // has neither `fontSizeHundredths` nor `fontFamily` set — exactly
    // what `set-shape-text` produces when the user clicks an empty
    // title placeholder and types "test". Before the placeholder-
    // defaults wiring this rendered at the 18pt fallback, so a freshly
    // typed title looked like body copy.
    const title: TextShape = {
      id: "txt-title",
      kind: "text",
      cNvPrId: 7,
      name: "Title 1",
      position: { xEmu: 0, yEmu: 0 },
      size: { cxEmu: 4_000_000, cyEmu: 1_000_000 },
      placeholder: { type: "title" },
      spPrTail: [],
      txBody: {
        bodyPrRaw: undefined,
        paragraphs: [
          {
            properties: {},
            runs: [{ text: "test", isLineBreak: false, properties: {} }],
          },
        ],
      },
    } as unknown as TextShape;
    const svg = shapeToSvg(title, { slideSize: { cxEmu: 9_144_000, cyEmu: 6_858_000 } });
    // 36pt → 36 * 96/72 = 48 px. The shorthand string is what the
    // foreignObject inline style emits; assert the px value rather
    // than parsing the style declaration so the test stays portable
    // even if the property order changes.
    expect(svg).toContain("font-size:48px");
    // The placeholder font family is now wrapped via `wrapFontFamily`
    // so `Calibri` composes with the `@font-face` aliases in
    // `apps/web/app/globals.css` (which redefine `Calibri` → bundled
    // Carlito). The trailing `system-ui, sans-serif` is the
    // unknown-font fallback contract from the helper.
    expect(svg).toContain("font-family:Calibri, system-ui, sans-serif");
  });

  it("subtitle placeholder defaults to 24pt centered text in the foreignObject", () => {
    const sub: TextShape = {
      id: "txt-sub",
      kind: "text",
      cNvPrId: 8,
      name: "Subtitle 1",
      position: { xEmu: 0, yEmu: 0 },
      size: { cxEmu: 4_000_000, cyEmu: 1_000_000 },
      placeholder: { type: "subTitle" },
      spPrTail: [],
      txBody: {
        bodyPrRaw: undefined,
        paragraphs: [
          {
            properties: {},
            runs: [{ text: "hello", isLineBreak: false, properties: {} }],
          },
        ],
      },
    } as unknown as TextShape;
    const svg = shapeToSvg(sub, { slideSize: { cxEmu: 9_144_000, cyEmu: 6_858_000 } });
    expect(svg).toContain("font-size:32px"); // 24 * 96/72
    expect(svg).toContain("text-align:center");
  });

  it("hides entrance-animation targets in the markup when ctx.hiddenCNvPrIds includes them", () => {
    const shape: TextShape = {
      id: "txt-anim",
      kind: "text",
      cNvPrId: 99,
      name: "Title 1",
      position: { xEmu: 0, yEmu: 0 },
      size: { cxEmu: 4_000_000, cyEmu: 1_000_000 },
      placeholder: { type: "title" },
      spPrTail: [],
      txBody: {
        bodyPrRaw: undefined,
        paragraphs: [{ properties: {}, runs: [{ text: "x", isLineBreak: false, properties: {} }] }],
      },
    } as unknown as TextShape;
    const svg = shapeToSvg(shape, {
      slideSize: { cxEmu: 9_144_000, cyEmu: 6_858_000 },
      hiddenCNvPrIds: new Set([99]),
    });
    // The wrapping <g> should carry visibility:hidden so the very first
    // paint already conceals the shape — otherwise present mode flashes
    // the entrance target before `prepare()` runs in a useEffect.
    expect(svg).toMatch(/<g class="anim-target" data-cnvprid="99" style="visibility:hidden;opacity:0">/);
    // Sibling shapes that aren't entrance targets still render with
    // the bare wrapper — no inline style.
    const otherSvg = shapeToSvg({ ...shape, cNvPrId: 100 } as unknown as TextShape, {
      slideSize: { cxEmu: 9_144_000, cyEmu: 6_858_000 },
      hiddenCNvPrIds: new Set([99]),
    });
    expect(otherSvg).toMatch(/<g class="anim-target" data-cnvprid="100">/);
  });

  it("emits the four connector end-shape markers in <defs> exactly once per slide", async () => {
    const snap = await load("01-blank.pptx");
    const slide = snap.root.slides[0];
    const svg = slideToSvgString(slide, { slideSize: snap.root.slideSize });
    // Each marker definition should appear exactly once — duplicate
    // <marker id="…"> entries would make some browsers ignore later
    // references with a console warning.
    expect((svg.match(/id="cxn-arrow"/g) ?? []).length).toBe(1);
    expect((svg.match(/id="cxn-triangle"/g) ?? []).length).toBe(1);
    expect((svg.match(/id="cxn-oval"/g) ?? []).length).toBe(1);
    expect((svg.match(/id="cxn-none"/g) ?? []).length).toBe(1);
  });

  it("selects per-end markers based on headEnd / tailEnd on the connector", () => {
    const shape: ConnectorShape = {
      id: "cxn-marker",
      kind: "connector",
      cNvPrId: 12,
      name: "Marker 1",
      position: { xEmu: 0, yEmu: 0 },
      size: { cxEmu: 1_000_000, cyEmu: 0 },
      nvCxnSpPrTail: [],
      spPrTail: [],
      connectorType: "straight",
      start: { kind: "free", xEmu: 0, yEmu: 0 },
      end: { kind: "free", xEmu: 1_000_000, yEmu: 0 },
      headEnd: "triangle",
      tailEnd: "oval",
    };
    const svg = shapeToSvg(shape, { slideSize: { cxEmu: 9_144_000, cyEmu: 6_858_000 } });
    // headEnd → marker-end, tailEnd → marker-start. Mismatching either
    // side would visually swap the arrowhead.
    expect(svg).toContain('marker-end="url(#cxn-triangle)"');
    expect(svg).toContain('marker-start="url(#cxn-oval)"');
  });

  it("omits marker references when an end is set to 'none'", () => {
    const shape: ConnectorShape = {
      id: "cxn-marker-none",
      kind: "connector",
      cNvPrId: 13,
      name: "Marker 2",
      position: { xEmu: 0, yEmu: 0 },
      size: { cxEmu: 1_000_000, cyEmu: 0 },
      nvCxnSpPrTail: [],
      spPrTail: [],
      connectorType: "straight",
      start: { kind: "free", xEmu: 0, yEmu: 0 },
      end: { kind: "free", xEmu: 1_000_000, yEmu: 0 },
      headEnd: "none",
      tailEnd: "none",
    };
    const svg = shapeToSvg(shape, { slideSize: { cxEmu: 9_144_000, cyEmu: 6_858_000 } });
    expect(svg).not.toContain("marker-end=");
    expect(svg).not.toContain("marker-start=");
  });

  it("renders an elbow connector with rounded corners (smooth path, not raw polyline)", () => {
    const shape: ConnectorShape = {
      id: "cxn-2",
      kind: "connector",
      cNvPrId: 11,
      name: "Elbow 1",
      position: { xEmu: 0, yEmu: 0 },
      size: { cxEmu: 2_000_000, cyEmu: 2_000_000 },
      nvCxnSpPrTail: [],
      spPrTail: [],
      connectorType: "elbow",
      start: { kind: "free", xEmu: 0, yEmu: 0 },
      end: { kind: "free", xEmu: 2_000_000, yEmu: 2_000_000 },
    };
    const svg = shapeToSvg(shape, { slideSize: { cxEmu: 9_144_000, cyEmu: 6_858_000 } });
    // Visible layer is now a `<path>` with quadratic rounds, not a
    // raw `<polyline>`. The hit-testing layer remains a polyline so
    // we should see exactly one polyline (hit) and one path (visible).
    const polylineCount = (svg.match(/<polyline/g) ?? []).length;
    const pathCount = (svg.match(/<path/g) ?? []).length;
    expect(polylineCount).toBe(1);
    expect(pathCount).toBeGreaterThanOrEqual(1);
    // Rounded corners produce `Q` commands inside the visible path.
    expect(svg).toMatch(/<path d="M [^"]*Q /);
  });
});
