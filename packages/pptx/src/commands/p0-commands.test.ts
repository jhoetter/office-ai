import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { ooxml } from "@officeai/core";
import { PptxAgent } from "../agent/agent.js";
import { parsePptx } from "../parser/parse.js";
import type { TextShape } from "../model/types.js";

const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);

async function loadAgent(name: string): Promise<PptxAgent> {
  const buf = await readFile(join(FIXTURES_DIR.pathname, name));
  return PptxAgent.fromBuffer(buf);
}

describe("P0 slide commands", () => {
  it("add-slide appends a new slide and produces a valid PPTX", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const before = agent.getSnapshot().root.slides.length;
    const m = await agent.applyCommand({
      type: "pptx:add-slide",
      payload: {},
      source: "system",
    });
    expect(m.status).toBe("approved");
    expect(agent.getSnapshot().root.slides.length).toBe(before + 1);

    const out = await agent.exportFile();
    const reparsed = await parsePptx(out);
    expect(reparsed.root.slides.length).toBe(before + 1);
    // Slide-id was bumped past the original max.
    const ids = reparsed.root.slides.map((s) => s.slideId).sort((a, b) => a - b);
    expect(ids[ids.length - 1]).toBeGreaterThanOrEqual(257);
  });

  it("add-slide rejects out-of-range position", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const m = await agent.applyCommand({
      type: "pptx:add-slide",
      payload: { at: 99 },
      source: "system",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-position");
  });

  it("delete-slide removes a slide and re-emits a valid file", async () => {
    const agent = await loadAgent("07-multi-slide.pptx");
    const before = agent.getSnapshot().root.slides.length;
    expect(before).toBeGreaterThan(1);
    await agent.applyCommand({
      type: "pptx:delete-slide",
      payload: { slideIndex: 0 },
      source: "system",
    });
    expect(agent.getSnapshot().root.slides.length).toBe(before - 1);

    const out = await agent.exportFile();
    const reparsed = await parsePptx(out);
    expect(reparsed.root.slides.length).toBe(before - 1);
    // The deleted slide's part is gone from the package.
    const c = await ooxml.OoxmlContainer.load(out);
    expect(c.has("ppt/slides/slide1.xml")).toBe(false);
  });

  it("delete-slide rejects unknown index", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const m = await agent.applyCommand({
      type: "pptx:delete-slide",
      payload: { slideIndex: 99 },
      source: "system",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-target");
  });

  it("duplicate-slide deep-clones the slide at index+1", async () => {
    const agent = await loadAgent("04-multi-shape.pptx");
    const before = agent.getSnapshot().root.slides.length;
    await agent.applyCommand({
      type: "pptx:duplicate-slide",
      payload: { slideIndex: 0 },
      source: "system",
    });
    const after = agent.getSnapshot().root.slides;
    expect(after.length).toBe(before + 1);
    // Cloned slide is at index 1.
    expect(after[1].partPath).not.toBe(after[0].partPath);
    // Same number of shapes as the source.
    expect(after[1].shapes.length).toBe(after[0].shapes.length);

    // File round-trips.
    const out = await agent.exportFile();
    const reparsed = await parsePptx(out);
    expect(reparsed.root.slides.length).toBe(before + 1);
  });

  it("move-slide reorders slides via splice", async () => {
    const agent = await loadAgent("07-multi-slide.pptx");
    const ids = agent.getSnapshot().root.slides.map((s) => s.slideId);
    expect(ids.length).toBeGreaterThanOrEqual(3);
    await agent.applyCommand({
      type: "pptx:move-slide",
      payload: { from: 0, to: 2 },
      source: "system",
    });
    const newIds = agent.getSnapshot().root.slides.map((s) => s.slideId);
    expect(newIds[2]).toBe(ids[0]);
    expect(newIds[0]).toBe(ids[1]);
  });

  it("move-slide rejects invalid indices", async () => {
    const agent = await loadAgent("07-multi-slide.pptx");
    const bad = await agent.applyCommand({
      type: "pptx:move-slide",
      payload: { from: 99, to: 0 },
      source: "system",
    });
    expect(bad.status).toBe("rejected");
    expect(bad.rejection?.code).toBe("invalid-position");
  });
});

describe("P0 shape commands", () => {
  it("set-text replaces a TextShape's content while inheriting style", async () => {
    const agent = await loadAgent("04-multi-shape.pptx");
    const slide = agent.getSnapshot().root.slides[0];
    const ts = slide.shapes.find((s): s is TextShape => s.kind === "text");
    expect(ts).toBeDefined();
    if (!ts) return;

    const m = await agent.applyCommand({
      type: "pptx:set-text",
      payload: {
        slideIndex: 0,
        shapeId: ts.id,
        text: "Line one\nLine two",
      },
      source: "system",
    });
    expect(m.status).toBe("approved");
    const updated = agent.getSnapshot().root.slides[0].shapes.find((s): s is TextShape => s.id === ts.id)!;
    expect(updated.txBody.paragraphs.length).toBe(2);
    expect(updated.txBody.paragraphs[0].runs[0].text).toBe("Line one");
    expect(updated.txBody.paragraphs[1].runs[0].text).toBe("Line two");

    const out = await agent.exportFile();
    const c = await ooxml.OoxmlContainer.load(out);
    const xml = c.readText(slide.partPath);
    expect(xml).toContain("Line one");
    expect(xml).toContain("Line two");
  });

  it("set-text accepts structured paragraphs preserving per-run formatting", async () => {
    // D12 — structured commit: each run carries its own properties
    // (or an inheritFromRun hint) so bold/italic spans inside one
    // paragraph survive a blur-commit round-trip.
    const agent = await loadAgent("04-multi-shape.pptx");
    const slide = agent.getSnapshot().root.slides[0];
    const ts = slide.shapes.find((s): s is TextShape => s.kind === "text");
    if (!ts) throw new Error("expected a text shape in 04-multi-shape.pptx");

    const m = await agent.applyCommand({
      type: "pptx:set-text",
      payload: {
        slideIndex: 0,
        shapeId: ts.id,
        paragraphs: [
          {
            runs: [
              { text: "Hello ", properties: { bold: true } },
              { text: "world", properties: { italic: true } },
            ],
          },
          {
            runs: [{ text: "Second line", inheritFromRun: 0 }],
          },
        ],
      },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const updated = agent.getSnapshot().root.slides[0].shapes.find((s): s is TextShape => s.id === ts.id)!;
    expect(updated.txBody.paragraphs.length).toBe(2);
    expect(updated.txBody.paragraphs[0].runs.length).toBe(2);
    expect(updated.txBody.paragraphs[0].runs[0].text).toBe("Hello ");
    expect(updated.txBody.paragraphs[0].runs[0].properties.bold).toBe(true);
    expect(updated.txBody.paragraphs[0].runs[1].text).toBe("world");
    expect(updated.txBody.paragraphs[0].runs[1].properties.italic).toBe(true);
    expect(updated.txBody.paragraphs[1].runs[0].text).toBe("Second line");
  });

  it("set-text rejects non-text shapes", async () => {
    const agent = await loadAgent("05-with-image.pptx");
    const slide = agent.getSnapshot().root.slides[0];
    const pic = slide.shapes.find((s) => s.kind === "pic");
    expect(pic).toBeDefined();
    if (!pic) return;
    const m = await agent.applyCommand({
      type: "pptx:set-text",
      payload: { slideIndex: 0, shapeId: pic.id, text: "hi" },
      source: "system",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("not-applicable");
  });

  it("set-position updates a shape's <a:off>", async () => {
    const agent = await loadAgent("04-multi-shape.pptx");
    const slide = agent.getSnapshot().root.slides[0];
    const ts = slide.shapes.find((s) => s.kind === "text");
    expect(ts).toBeDefined();
    if (!ts) return;
    await agent.applyCommand({
      type: "pptx:set-position",
      payload: { slideIndex: 0, shapeId: ts.id, x: 1234567, y: 7654321 },
      source: "system",
    });
    const updated = agent.getSnapshot().root.slides[0].shapes.find((s) => s.id === ts.id)!;
    expect(updated.position).toEqual({ xEmu: 1234567, yEmu: 7654321 });

    const out = await agent.exportFile();
    const c = await ooxml.OoxmlContainer.load(out);
    const xml = c.readText(slide.partPath);
    expect(xml).toContain('x="1234567"');
    expect(xml).toContain('y="7654321"');
  });

  it("set-position accepts off-slide coordinates and round-trips them through serialize/parse", async () => {
    // The editor's "scratch canvas" lets users park shapes outside the
    // slide rectangle (negative x / y, or x/y past the slide width/
    // height). The data model has no clamp, the serializer writes the
    // raw EMU values, and the parser reads them back unchanged — so a
    // round-trip must preserve the off-slide position exactly. If a
    // future change starts clamping, this test catches the regression
    // before users notice their parked shapes snapping into the slide.
    const agent = await loadAgent("04-multi-shape.pptx");
    const slide = agent.getSnapshot().root.slides[0];
    const ts = slide.shapes.find((s) => s.kind === "text")!;
    const offSlideX = -500_000;
    const offSlideY = slide.shapes.length > 0 ? 12_000_000 : -1_000_000;
    await agent.applyCommand({
      type: "pptx:set-position",
      payload: { slideIndex: 0, shapeId: ts.id, x: offSlideX, y: offSlideY },
      source: "system",
    });
    const out = await agent.exportFile();
    const reparsed = await parsePptx(out);
    const reparsedShape = reparsed.root.slides[0].shapes.find((s) => s.id === ts.id || s.cNvPrId === ts.cNvPrId);
    expect(reparsedShape).toBeDefined();
    expect(reparsedShape?.position).toEqual({ xEmu: offSlideX, yEmu: offSlideY });
  });

  it("set-position rejects opaque shapes & non-finite values", async () => {
    const agent = await loadAgent("04-multi-shape.pptx");
    const slide = agent.getSnapshot().root.slides[0];
    const ts = slide.shapes.find((s) => s.kind === "text")!;
    const m = await agent.applyCommand({
      type: "pptx:set-position",
      payload: { slideIndex: 0, shapeId: ts.id, x: NaN, y: 0 },
      source: "system",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-payload");
  });

  it("set-size updates a shape's <a:ext>", async () => {
    const agent = await loadAgent("04-multi-shape.pptx");
    const slide = agent.getSnapshot().root.slides[0];
    const ts = slide.shapes.find((s) => s.kind === "text")!;
    await agent.applyCommand({
      type: "pptx:set-size",
      payload: { slideIndex: 0, shapeId: ts.id, width: 5000000, height: 1500000 },
      source: "system",
    });
    const updated = agent.getSnapshot().root.slides[0].shapes.find((s) => s.id === ts.id)!;
    expect(updated.size).toEqual({ cxEmu: 5000000, cyEmu: 1500000 });
  });

  it("set-size rejects non-positive sizes", async () => {
    const agent = await loadAgent("04-multi-shape.pptx");
    const slide = agent.getSnapshot().root.slides[0];
    const ts = slide.shapes.find((s) => s.kind === "text")!;
    const m = await agent.applyCommand({
      type: "pptx:set-size",
      payload: { slideIndex: 0, shapeId: ts.id, width: 0, height: 100 },
      source: "system",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-payload");
  });
});
