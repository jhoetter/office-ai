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

describe("pptx:add-shape", () => {
  it("inserts a rectangle with a typed solidFill so the renderer/serializer see it", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const before = agent.getSnapshot().root.slides[0].shapes.length;

    const m = await agent.applyCommand({
      type: "pptx:add-shape",
      payload: {
        slideIndex: 0,
        preset: "rect",
        x: 1_000_000,
        y: 1_000_000,
        width: 2_000_000,
        height: 1_000_000,
        fill: "FF8800",
      },
      source: "human",
    });
    expect(m.status).toBe("approved");

    const slide = agent.getSnapshot().root.slides[0];
    expect(slide.shapes.length).toBe(before + 1);
    const added = slide.shapes[slide.shapes.length - 1] as TextShape;
    expect(added.kind).toBe("text");
    expect(added.position?.xEmu).toBe(1_000_000);
    expect(added.size?.cxEmu).toBe(2_000_000);

    // Confirm both prstGeom + solidFill made it into spPrTail in the right
    // order — the serializer + renderer both rely on this layout.
    const prstIdx = added.spPrTail.findIndex((c) => c.tag === "a:prstGeom");
    const fillIdx = added.spPrTail.findIndex((c) => c.tag === "a:solidFill");
    expect(prstIdx).toBeGreaterThanOrEqual(0);
    expect(fillIdx).toBeGreaterThan(prstIdx);

    const out = await agent.exportFile();
    const c = await ooxml.OoxmlContainer.load(out);
    const xml = c.readText(slide.partPath);
    expect(xml).toContain('prst="rect"');
    expect(xml.toUpperCase()).toContain("FF8800");
  });

  it("emits a stroke (a:ln) when preset is line, even without an explicit fill", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const m = await agent.applyCommand({
      type: "pptx:add-shape",
      payload: {
        slideIndex: 0,
        preset: "line",
        x: 0,
        y: 0,
        width: 5_000_000,
        height: 0,
        stroke: "112233",
      },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const added = agent.getSnapshot().root.slides[0].shapes.slice(-1)[0] as TextShape;
    const lnIdx = added.spPrTail.findIndex((c) => c.tag === "a:ln");
    expect(lnIdx).toBeGreaterThanOrEqual(0);
    // Lines must NOT carry a solidFill at the spPr level — they would draw
    // a thin filled band underneath the stroke, doubling the visual weight.
    expect(added.spPrTail.findIndex((c) => c.tag === "a:solidFill")).toBe(-1);
  });

  it("rejects unknown presets so we don't silently emit garbage XML", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const m = await agent.applyCommand({
      type: "pptx:add-shape",
      payload: {
        slideIndex: 0,
        // @ts-expect-error — testing runtime guard
        preset: "blob",
        x: 0,
        y: 0,
        width: 1_000_000,
        height: 1_000_000,
      },
      source: "human",
    });
    expect(m.status).toBe("rejected");
  });
});

describe("pptx:delete-shape", () => {
  it("removes a top-level shape and bumps revision", async () => {
    const agent = await loadAgent("04-multi-shape.pptx");
    const slide = agent.getSnapshot().root.slides[0];
    const target = slide.shapes[0]!;
    const beforeCount = slide.shapes.length;
    const beforeRev = agent.getSnapshot().revision;

    const m = await agent.applyCommand({
      type: "pptx:delete-shape",
      payload: { slideIndex: 0, shapeId: target.id },
      source: "human",
    });
    expect(m.status).toBe("approved");

    const after = agent.getSnapshot();
    expect(after.root.slides[0].shapes.length).toBe(beforeCount - 1);
    expect(after.root.slides[0].shapes.find((s) => s.id === target.id)).toBeUndefined();
    expect(after.revision).toBeGreaterThan(beforeRev);
  });

  it("returns rejected status for an unknown shape id", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const m = await agent.applyCommand({
      type: "pptx:delete-shape",
      payload: { slideIndex: 0, shapeId: "node-does-not-exist" },
      source: "human",
    });
    expect(m.status).toBe("rejected");
  });
});

describe("pptx:set-shape-fill", () => {
  it("replaces an existing solidFill with a new color and survives serialization", async () => {
    const agent = await loadAgent("01-blank.pptx");
    // Seed a shape with a known fill we can then change.
    await agent.applyCommand({
      type: "pptx:add-shape",
      payload: {
        slideIndex: 0,
        preset: "rect",
        x: 0,
        y: 0,
        width: 1_000_000,
        height: 1_000_000,
        fill: "AA0000",
      },
      source: "human",
    });
    const slide = agent.getSnapshot().root.slides[0];
    const target = slide.shapes.slice(-1)[0]!;

    const m = await agent.applyCommand({
      type: "pptx:set-shape-fill",
      payload: { slideIndex: 0, shapeId: target.id, fill: "00AA00" },
      source: "human",
    });
    expect(m.status).toBe("approved");

    const updated = agent.getSnapshot().root.slides[0].shapes.find((s) => s.id === target.id) as TextShape;
    const solidFills = updated.spPrTail.filter((c) => c.tag === "a:solidFill");
    expect(solidFills.length).toBe(1);

    const out = await agent.exportFile();
    const c = await ooxml.OoxmlContainer.load(out);
    const xml = c.readText(slide.partPath);
    expect(xml.toUpperCase()).toContain("00AA00");
    expect(xml.toUpperCase()).not.toContain("AA0000");
  });

  it("clearing the fill replaces solidFill with noFill (transparent)", async () => {
    const agent = await loadAgent("01-blank.pptx");
    await agent.applyCommand({
      type: "pptx:add-shape",
      payload: {
        slideIndex: 0,
        preset: "rect",
        x: 0,
        y: 0,
        width: 1_000_000,
        height: 1_000_000,
        fill: "123456",
      },
      source: "human",
    });
    const target = agent.getSnapshot().root.slides[0].shapes.slice(-1)[0]!;
    await agent.applyCommand({
      type: "pptx:set-shape-fill",
      payload: { slideIndex: 0, shapeId: target.id, fill: null },
      source: "human",
    });
    const updated = agent.getSnapshot().root.slides[0].shapes.find((s) => s.id === target.id) as TextShape;
    expect(updated.spPrTail.some((c) => c.tag === "a:solidFill")).toBe(false);
    expect(updated.spPrTail.some((c) => c.tag === "a:noFill")).toBe(true);
  });
});
