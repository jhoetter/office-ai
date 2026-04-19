import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { ooxml } from "@officeai/core";
import { PptxAgent } from "../agent/agent.js";
import type { TextShape } from "../model/types.js";

const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);

async function loadAgent(name: string): Promise<PptxAgent> {
  const buf = await readFile(join(FIXTURES_DIR.pathname, name));
  return PptxAgent.fromBuffer(buf);
}

describe("slide layouts (typed promotion)", () => {
  it("classifies the deck's existing layouts into one of the standard kinds", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const root = agent.getSnapshot().root;
    expect(root.layouts.size).toBeGreaterThan(0);
    for (const layout of root.layouts.values()) {
      expect(typeof layout.kind).toBe("string");
      expect(typeof layout.name).toBe("string");
      expect(Array.isArray(layout.placeholders)).toBe(true);
    }
  });
});

describe("pptx:add-slide with layoutKind", () => {
  it("synthesises a built-in layout part when the deck doesn't already have one of that kind", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const beforeLayouts = agent.getSnapshot().root.layouts.size;
    const m = await agent.applyCommand({
      type: "pptx:add-slide",
      payload: { layoutKind: "twoContent" },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const after = agent.getSnapshot();
    // A new layout part was added (synthesised on demand).
    expect(after.root.layouts.size).toBeGreaterThanOrEqual(beforeLayouts);
    // The new slide points at a `twoContent` layout.
    const newSlide = after.root.slides[after.root.slides.length - 1];
    const layout = after.root.layouts.get(newSlide.layoutPartPath ?? "");
    expect(layout?.kind).toBe("twoContent");
    // Placeholders were stamped onto the slide.
    const placeholderShapes = newSlide.shapes.filter(
      (s) => s.kind === "text" && (s as TextShape).placeholder
    );
    expect(placeholderShapes.length).toBe(layout!.placeholders.length);
  });

  it("survives a serialize → re-parse roundtrip with the synthesised layout part still present", async () => {
    const agent = await loadAgent("01-blank.pptx");
    await agent.applyCommand({
      type: "pptx:add-slide",
      payload: { layoutKind: "titleAndContent" },
      source: "human",
    });
    const out = await agent.exportFile();
    const c = await ooxml.OoxmlContainer.load(out);
    // The synthesised layout part should be present in the package.
    const layoutPaths = [...c.parts.keys()].filter((p) => p.startsWith("ppt/slideLayouts/"));
    expect(layoutPaths.length).toBeGreaterThan(0);
    // The new slide's rels file must point at one of them.
    const slideRelPaths = [...c.parts.keys()].filter((p) => p.startsWith("ppt/slides/_rels/"));
    const lastRelXml = c.readText(slideRelPaths[slideRelPaths.length - 1]);
    expect(lastRelXml).toContain("slideLayout");

    // Round-trip the bytes back through the parser and confirm classification holds.
    const reAgent = await PptxAgent.fromBuffer(out);
    const reRoot = reAgent.getSnapshot().root;
    const lastSlide = reRoot.slides[reRoot.slides.length - 1];
    const layout = reRoot.layouts.get(lastSlide.layoutPartPath ?? "");
    expect(layout?.kind).toBe("titleAndContent");
  });
});

describe("pptx:set-slide-layout", () => {
  it("re-points a slide's layout and preserves placeholder text where idx matches", async () => {
    const agent = await loadAgent("01-blank.pptx");
    // First, add a slide using the title layout so we have known placeholders.
    await agent.applyCommand({
      type: "pptx:add-slide",
      payload: { layoutKind: "title" },
      source: "human",
    });
    const sIdx = agent.getSnapshot().root.slides.length - 1;
    const slide = agent.getSnapshot().root.slides[sIdx];
    const titleShape = slide.shapes.find(
      (s) => s.kind === "text" && (s as TextShape).placeholder?.idx === 0
    ) as TextShape;
    expect(titleShape).toBeDefined();
    // Type something into the title.
    await agent.applyCommand({
      type: "pptx:set-text",
      payload: { slideIndex: sIdx, shapeId: titleShape.id, text: "Hello" },
      source: "human",
    });
    // Switch layout.
    const m = await agent.applyCommand({
      type: "pptx:set-slide-layout",
      payload: { slideIndex: sIdx, layoutKind: "titleAndContent" },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const after = agent.getSnapshot().root.slides[sIdx];
    const layout = agent.getSnapshot().root.layouts.get(after.layoutPartPath ?? "");
    expect(layout?.kind).toBe("titleAndContent");
    // Title (idx 0) survived the layout swap.
    const newTitle = after.shapes.find(
      (s) => s.kind === "text" && (s as TextShape).placeholder?.idx === 0
    ) as TextShape;
    expect(newTitle).toBeDefined();
    const text = newTitle.txBody.paragraphs.flatMap((p) => p.runs.map((r) => r.text)).join("");
    expect(text).toBe("Hello");
  });

  it("falls back to built-in placeholders when the deck's matching layout has none", async () => {
    // Construct a deck whose `title` layout is intentionally placeholder-
    // less (mirrors `apps/web/app/lib/sample-pptx.ts` and other minimal
    // decks). Picking "Title Slide" from the menu must still stamp the
    // built-in `ctrTitle` + `subTitle` placeholders so the user sees the
    // ghost-hint UI on the new slide.
    const agent = await loadAgent("01-blank.pptx");
    const snap = agent.getSnapshot();
    const titleLayout = [...snap.root.layouts.values()].find((l) => l.kind === "title");
    if (!titleLayout) {
      // Stub a title layout into the deck so the test stays meaningful
      // even if the fixture stops shipping one.
      const partPath = "ppt/slideLayouts/__empty_title.xml";
      const enriched = new Map(snap.root.layouts);
      enriched.set(partPath, {
        partPath,
        kind: "title",
        name: "Title Slide",
        placeholders: [],
        raw: { tag: "p:sldLayout", attrs: {}, rawAttrs: {}, subtree: [] },
      });
      // @ts-expect-error – test-only mutation to seed the scenario
      snap.root.layouts = enriched;
    } else {
      // Force the existing title layout to have zero placeholders.
      // @ts-expect-error – test-only mutation to seed the scenario
      titleLayout.placeholders = [];
    }
    await agent.applyCommand({
      type: "pptx:add-slide",
      payload: { layoutKind: "title" },
      source: "human",
    });
    const after = agent.getSnapshot();
    const newSlide = after.root.slides[after.root.slides.length - 1];
    const placeholderShapes = newSlide.shapes.filter(
      (s) => s.kind === "text" && (s as TextShape).placeholder
    ) as TextShape[];
    expect(placeholderShapes.length).toBeGreaterThan(0);
    const types = new Set(placeholderShapes.map((s) => s.placeholder?.type));
    expect(types.has("ctrTitle")).toBe(true);
    expect(types.has("subTitle")).toBe(true);
  });

  it("rejects when neither layoutPartPath nor layoutKind is supplied", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const m = await agent.applyCommand({
      type: "pptx:set-slide-layout",
      payload: { slideIndex: 0 },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-payload");
  });
});
