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

function findFirstTextShape(slideShapes: ReadonlyArray<unknown>): TextShape | null {
  for (const s of slideShapes as ReadonlyArray<{ kind: string }>) {
    if (s.kind === "text") return s as unknown as TextShape;
  }
  return null;
}

describe("pptx:set-paragraph-alignment", () => {
  it("applies center alignment to a single paragraph and re-emits algn=\"ctr\"", async () => {
    const agent = await loadAgent("04-multi-shape.pptx");
    const slide = agent.getSnapshot().root.slides[0];
    const ts = findFirstTextShape(slide.shapes);
    expect(ts).not.toBeNull();
    if (!ts) return;

    const m = await agent.applyCommand({
      type: "pptx:set-paragraph-alignment",
      payload: {
        slideIndex: 0,
        shapeId: ts.id,
        alignment: "center",
        paragraphs: [0],
      },
      source: "system",
    });
    expect(m.status).toBe("approved");

    const updated = agent.getSnapshot().root.slides[0].shapes.find(
      (s): s is TextShape => s.id === ts.id
    )!;
    expect(updated.txBody.paragraphs[0].properties.alignment).toBe("center");
    expect(updated.txBody.paragraphs[0].properties.opaqueAttrs?.algn).toBe("ctr");

    const out = await agent.exportFile();
    const c = await ooxml.OoxmlContainer.load(out);
    const xml = c.readText(slide.partPath);
    expect(xml).toContain('algn="ctr"');
  });

  it("applies shape-wide alignment when paragraphs is omitted", async () => {
    const agent = await loadAgent("04-multi-shape.pptx");
    const slide = agent.getSnapshot().root.slides[0];
    const ts = findFirstTextShape(slide.shapes);
    expect(ts).not.toBeNull();
    if (!ts) return;
    expect(ts.txBody.paragraphs.length).toBeGreaterThan(0);

    await agent.applyCommand({
      type: "pptx:set-paragraph-alignment",
      payload: { slideIndex: 0, shapeId: ts.id, alignment: "right" },
      source: "system",
    });

    const updated = agent.getSnapshot().root.slides[0].shapes.find(
      (s): s is TextShape => s.id === ts.id
    )!;
    for (const p of updated.txBody.paragraphs) {
      expect(p.properties.alignment).toBe("right");
      expect(p.properties.opaqueAttrs?.algn).toBe("r");
    }
  });

  it("clears alignment when alignment is null", async () => {
    const agent = await loadAgent("04-multi-shape.pptx");
    const slide = agent.getSnapshot().root.slides[0];
    const ts = findFirstTextShape(slide.shapes);
    if (!ts) throw new Error("expected a text shape");

    await agent.applyCommand({
      type: "pptx:set-paragraph-alignment",
      payload: { slideIndex: 0, shapeId: ts.id, alignment: "justify", paragraphs: [0] },
      source: "system",
    });
    await agent.applyCommand({
      type: "pptx:set-paragraph-alignment",
      payload: { slideIndex: 0, shapeId: ts.id, alignment: null, paragraphs: [0] },
      source: "system",
    });

    const updated = agent.getSnapshot().root.slides[0].shapes.find(
      (s): s is TextShape => s.id === ts.id
    )!;
    expect(updated.txBody.paragraphs[0].properties.alignment).toBeUndefined();
    expect(updated.txBody.paragraphs[0].properties.opaqueAttrs?.algn).toBeUndefined();

    const out = await agent.exportFile();
    const c = await ooxml.OoxmlContainer.load(out);
    const xml = c.readText(slide.partPath);
    // No algn attribute should be emitted on the cleared paragraph.
    // (Other paragraphs in the slide may still carry one — we just
    // assert the value we cleared isn't there as `just`.)
    expect(xml).not.toContain('algn="just"');
  });

  it("rejects invalid alignment values", async () => {
    const agent = await loadAgent("04-multi-shape.pptx");
    const slide = agent.getSnapshot().root.slides[0];
    const ts = findFirstTextShape(slide.shapes);
    if (!ts) throw new Error("expected a text shape");

    const m = await agent.applyCommand({
      type: "pptx:set-paragraph-alignment",
      payload: {
        slideIndex: 0,
        shapeId: ts.id,
        alignment: "bogus" as unknown as "left",
      },
      source: "system",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-payload");
    expect(m.rejection?.message).toMatch(/unknown alignment/);
  });

  it("rejects out-of-range paragraph indices", async () => {
    const agent = await loadAgent("04-multi-shape.pptx");
    const slide = agent.getSnapshot().root.slides[0];
    const ts = findFirstTextShape(slide.shapes);
    if (!ts) throw new Error("expected a text shape");

    const m = await agent.applyCommand({
      type: "pptx:set-paragraph-alignment",
      payload: { slideIndex: 0, shapeId: ts.id, alignment: "left", paragraphs: [99] },
      source: "system",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-payload");
    expect(m.rejection?.message).toMatch(/out of range/);
  });

  it("survives a serialize → parse round-trip", async () => {
    const agent = await loadAgent("04-multi-shape.pptx");
    const slide = agent.getSnapshot().root.slides[0];
    const ts = findFirstTextShape(slide.shapes);
    if (!ts) throw new Error("expected a text shape");

    await agent.applyCommand({
      type: "pptx:set-paragraph-alignment",
      payload: { slideIndex: 0, shapeId: ts.id, alignment: "center", paragraphs: [0] },
      source: "system",
    });

    const buf = await agent.exportFile();
    const reloaded = await PptxAgent.fromBuffer(buf);
    const slide2 = reloaded.getSnapshot().root.slides[0];
    const ts2 = findFirstTextShape(slide2.shapes);
    expect(ts2?.txBody.paragraphs[0].properties.alignment).toBe("center");
  });
});
