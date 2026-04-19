import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { ooxml } from "@officeai/core";
import { PptxAgent } from "../agent/agent.js";
import { parsePptx } from "../parser/parse.js";
import type { Picture, TextShape } from "../model/types.js";

const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);

async function loadAgent(name: string): Promise<PptxAgent> {
  const buf = await readFile(join(FIXTURES_DIR.pathname, name));
  return PptxAgent.fromBuffer(buf);
}

// Minimal valid 1x1 PNG (transparent).
const PNG_1X1 = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52, 0x00, 0x00,
  0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4, 0x89, 0x00, 0x00, 0x00,
  0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d,
  0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
]);

const PNG_OTHER = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52, 0x00, 0x00,
  0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4, 0x89, 0x00, 0x00, 0x00,
  0x10, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x62, 0xfc, 0xff, 0xff, 0x3f, 0x03, 0x21, 0x00, 0x00, 0x09, 0xfb,
  0x03, 0xfd, 0x18, 0xf3, 0xb8, 0xa9, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
]);

describe("P1: format-text", () => {
  it('toggles bold on a sub-range and produces XML containing b="1"', async () => {
    const agent = await loadAgent("04-multi-shape.pptx");
    const slide = agent.getSnapshot().root.slides[0];
    const ts = slide.shapes.find((s): s is TextShape => s.kind === "text" && s.txBody.paragraphs.length > 0);
    expect(ts).toBeDefined();
    if (!ts) return;
    const para = ts.txBody.paragraphs[0];
    const flatLen = para.runs.reduce((a, r) => a + (r.isLineBreak ? 0 : r.text.length), 0);
    expect(flatLen).toBeGreaterThan(2);

    const m = await agent.applyCommand({
      type: "pptx:format-text",
      payload: {
        slideIndex: 0,
        shapeId: ts.id,
        range: { paragraph: 0, start: 0, end: Math.min(2, flatLen) },
        format: { bold: true },
      },
      source: "system",
    });
    expect(m.status).toBe("approved");

    const updated = agent.getSnapshot().root.slides[0].shapes.find((s): s is TextShape => s.id === ts.id)!;
    const firstRun = updated.txBody.paragraphs[0].runs[0];
    expect(firstRun.properties.bold).toBe(true);

    const out = await agent.exportFile();
    const c = await ooxml.OoxmlContainer.load(out);
    const xml = c.readText(slide.partPath);
    expect(xml).toContain('b="1"');
  });

  it("applies color and font and re-emits a:solidFill / a:latin from model", async () => {
    const agent = await loadAgent("04-multi-shape.pptx");
    const slide = agent.getSnapshot().root.slides[0];
    const ts = slide.shapes.find((s): s is TextShape => s.kind === "text" && s.txBody.paragraphs.length > 0)!;
    const para = ts.txBody.paragraphs[0];
    const flatLen = para.runs.reduce((a, r) => a + (r.isLineBreak ? 0 : r.text.length), 0);

    await agent.applyCommand({
      type: "pptx:format-text",
      payload: {
        slideIndex: 0,
        shapeId: ts.id,
        range: { paragraph: 0, start: 0, end: flatLen },
        format: { color: "FF0000", fontFamily: "Arial" },
      },
      source: "system",
    });

    const out = await agent.exportFile();
    const c = await ooxml.OoxmlContainer.load(out);
    const xml = c.readText(slide.partPath);
    expect(xml).toContain('val="FF0000"');
    expect(xml).toContain('typeface="Arial"');
  });

  it("applies highlight and re-emits a:highlight from the model", async () => {
    const agent = await loadAgent("04-multi-shape.pptx");
    const slide = agent.getSnapshot().root.slides[0];
    const ts = slide.shapes.find((s): s is TextShape => s.kind === "text" && s.txBody.paragraphs.length > 0)!;
    const para = ts.txBody.paragraphs[0];
    const flatLen = para.runs.reduce((a, r) => a + (r.isLineBreak ? 0 : r.text.length), 0);

    await agent.applyCommand({
      type: "pptx:format-text",
      payload: {
        slideIndex: 0,
        shapeId: ts.id,
        range: { paragraph: 0, start: 0, end: flatLen },
        format: { highlight: "FFFF00" },
      },
      source: "system",
    });

    const updated = agent.getSnapshot().root.slides[0].shapes.find((s): s is TextShape => s.id === ts.id)!;
    const firstRun = updated.txBody.paragraphs[0].runs[0];
    expect(firstRun.properties.highlight).toBe("FFFF00");

    const out = await agent.exportFile();
    const c = await ooxml.OoxmlContainer.load(out);
    const xml = c.readText(slide.partPath);
    expect(xml).toContain("a:highlight");
    expect(xml).toContain('val="FFFF00"');

    // Round-trip through the parser to confirm we can read what we wrote.
    const reparsed = await PptxAgent.fromBuffer(out);
    const reShapes = reparsed.getSnapshot().root.slides[0].shapes;
    const reTs = reShapes.find((s): s is TextShape => s.kind === "text" && s.txBody.paragraphs.length > 0);
    const reHighlights = (reTs?.txBody.paragraphs[0].runs ?? []).map((r) => r.properties.highlight);
    expect(reHighlights).toContain("FFFF00");
    void parsePptx;
  });

  it("clears highlight when format passes empty string", async () => {
    const agent = await loadAgent("04-multi-shape.pptx");
    const slide = agent.getSnapshot().root.slides[0];
    const ts = slide.shapes.find((s): s is TextShape => s.kind === "text" && s.txBody.paragraphs.length > 0)!;
    const para = ts.txBody.paragraphs[0];
    const flatLen = para.runs.reduce((a, r) => a + (r.isLineBreak ? 0 : r.text.length), 0);

    await agent.applyCommand({
      type: "pptx:format-text",
      payload: {
        slideIndex: 0,
        shapeId: ts.id,
        range: { paragraph: 0, start: 0, end: flatLen },
        format: { highlight: "FFFF00" },
      },
      source: "system",
    });
    await agent.applyCommand({
      type: "pptx:format-text",
      payload: {
        slideIndex: 0,
        shapeId: ts.id,
        range: { paragraph: 0, start: 0, end: flatLen },
        format: { highlight: "" },
      },
      source: "system",
    });

    const out = await agent.exportFile();
    const c = await ooxml.OoxmlContainer.load(out);
    const xml = c.readText(slide.partPath);
    expect(xml).not.toContain("a:highlight");
  });

  it("rejects out-of-bounds range", async () => {
    const agent = await loadAgent("04-multi-shape.pptx");
    const slide = agent.getSnapshot().root.slides[0];
    const ts = slide.shapes.find((s): s is TextShape => s.kind === "text")!;
    const m = await agent.applyCommand({
      type: "pptx:format-text",
      payload: {
        slideIndex: 0,
        shapeId: ts.id,
        range: { paragraph: 0, start: 0, end: 99999 },
        format: { italic: true },
      },
      source: "system",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-payload");
  });

  it("rejects non-text shapes", async () => {
    const agent = await loadAgent("05-with-image.pptx");
    const slide = agent.getSnapshot().root.slides[0];
    const pic = slide.shapes.find((s) => s.kind === "pic");
    expect(pic).toBeDefined();
    if (!pic) return;
    const m = await agent.applyCommand({
      type: "pptx:format-text",
      payload: {
        slideIndex: 0,
        shapeId: pic.id,
        range: { paragraph: 0, start: 0, end: 0 },
        format: { bold: true },
      },
      source: "system",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("not-applicable");
  });
});

describe("P1: insert-image", () => {
  it("adds a new picture, registers media, rels and content type", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const m = await agent.applyCommand({
      type: "pptx:insert-image",
      payload: {
        slideIndex: 0,
        data: PNG_1X1,
        mimeType: "image/png",
        x: 1000000,
        y: 2000000,
        width: 3000000,
        height: 4000000,
        altText: "test",
      },
      source: "system",
    });
    expect(m.status).toBe("approved");

    const slide = agent.getSnapshot().root.slides[0];
    const pic = slide.shapes.find((s): s is Picture => s.kind === "pic");
    expect(pic).toBeDefined();
    if (!pic) return;
    expect(pic.position).toEqual({ xEmu: 1000000, yEmu: 2000000 });
    expect(pic.size).toEqual({ cxEmu: 3000000, cyEmu: 4000000 });
    expect(pic.mediaPartPath).toMatch(/^ppt\/media\/image\d+\.png$/);

    const out = await agent.exportFile();
    const c = await ooxml.OoxmlContainer.load(out);
    expect(c.has(pic.mediaPartPath)).toBe(true);
    // Re-parse and verify the picture survives a roundtrip.
    const reparsed = await parsePptx(out);
    const repic = reparsed.root.slides[0].shapes.find((s): s is Picture => s.kind === "pic");
    expect(repic).toBeDefined();
    expect(repic!.mediaPartPath).toBe(pic.mediaPartPath);
  });

  it("dedups identical bytes via SHA-256 (same media path used twice)", async () => {
    const agent = await loadAgent("01-blank.pptx");
    await agent.applyCommand({
      type: "pptx:insert-image",
      payload: {
        slideIndex: 0,
        data: PNG_1X1,
        mimeType: "image/png",
        x: 0,
        y: 0,
        width: 1000000,
        height: 1000000,
      },
      source: "system",
    });
    await agent.applyCommand({
      type: "pptx:insert-image",
      payload: {
        slideIndex: 0,
        data: PNG_1X1,
        mimeType: "image/png",
        x: 2000000,
        y: 2000000,
        width: 1000000,
        height: 1000000,
      },
      source: "system",
    });
    const pics = agent.getSnapshot().root.slides[0].shapes.filter((s): s is Picture => s.kind === "pic");
    expect(pics.length).toBe(2);
    expect(pics[0].mediaPartPath).toBe(pics[1].mediaPartPath);
    // But two distinct slide rels entries (or one if also deduped).
    expect(pics[0].mediaRelId).toBe(pics[1].mediaRelId);
  });

  it("creates distinct media parts for distinct bytes", async () => {
    const agent = await loadAgent("01-blank.pptx");
    await agent.applyCommand({
      type: "pptx:insert-image",
      payload: {
        slideIndex: 0,
        data: PNG_1X1,
        mimeType: "image/png",
        x: 0,
        y: 0,
        width: 1000000,
        height: 1000000,
      },
      source: "system",
    });
    await agent.applyCommand({
      type: "pptx:insert-image",
      payload: {
        slideIndex: 0,
        data: PNG_OTHER,
        mimeType: "image/png",
        x: 2000000,
        y: 2000000,
        width: 1000000,
        height: 1000000,
      },
      source: "system",
    });
    const pics = agent.getSnapshot().root.slides[0].shapes.filter((s): s is Picture => s.kind === "pic");
    expect(pics.length).toBe(2);
    expect(pics[0].mediaPartPath).not.toBe(pics[1].mediaPartPath);
  });

  it("rejects unsupported MIME and zero size", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const m1 = await agent.applyCommand({
      type: "pptx:insert-image",
      payload: {
        slideIndex: 0,
        data: PNG_1X1,
        mimeType: "image/heic",
        x: 0,
        y: 0,
        width: 100,
        height: 100,
      },
      source: "system",
    });
    expect(m1.status).toBe("rejected");
    expect(m1.rejection?.code).toBe("invalid-payload");
    const m2 = await agent.applyCommand({
      type: "pptx:insert-image",
      payload: {
        slideIndex: 0,
        data: PNG_1X1,
        mimeType: "image/png",
        x: 0,
        y: 0,
        width: 0,
        height: 100,
      },
      source: "system",
    });
    expect(m2.status).toBe("rejected");
    expect(m2.rejection?.code).toBe("invalid-payload");
  });
});

describe("P1: add-text-box", () => {
  it("appends a TextShape to a slide and survives roundtrip", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const before = agent.getSnapshot().root.slides[0].shapes.length;
    const m = await agent.applyCommand({
      type: "pptx:add-text-box",
      payload: {
        slideIndex: 0,
        text: "Hello, world!",
        x: 100000,
        y: 200000,
        width: 5000000,
        height: 800000,
      },
      source: "system",
    });
    expect(m.status).toBe("approved");
    const slide = agent.getSnapshot().root.slides[0];
    expect(slide.shapes.length).toBe(before + 1);
    const tb = slide.shapes[slide.shapes.length - 1] as TextShape;
    expect(tb.kind).toBe("text");
    expect(tb.txBody.paragraphs[0].runs[0].text).toBe("Hello, world!");

    const out = await agent.exportFile();
    const reparsed = await parsePptx(out);
    const reSlide = reparsed.root.slides[0];
    const reTb = reSlide.shapes[reSlide.shapes.length - 1] as TextShape;
    expect(reTb.kind).toBe("text");
    expect(reTb.txBody.paragraphs[0].runs[0].text).toBe("Hello, world!");
    expect(reTb.position).toEqual({ xEmu: 100000, yEmu: 200000 });
    expect(reTb.size).toEqual({ cxEmu: 5000000, cyEmu: 800000 });
  });

  it("rejects invalid geometry", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const m = await agent.applyCommand({
      type: "pptx:add-text-box",
      payload: {
        slideIndex: 0,
        text: "hi",
        x: 0,
        y: 0,
        width: 0,
        height: 100,
      },
      source: "system",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-payload");
  });
});
