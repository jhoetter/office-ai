import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { ooxml } from "@officeai/core";
import { PptxAgent } from "../agent/agent.js";
import type { Picture, Shape } from "../model/types.js";

const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);

async function loadAgent(name: string): Promise<PptxAgent> {
  const buf = await readFile(join(FIXTURES_DIR.pathname, name));
  return PptxAgent.fromBuffer(buf);
}

function findFirstPicture(shapes: ReadonlyArray<Shape>): Picture | null {
  for (const s of shapes) {
    if (s.kind === "pic") return s;
  }
  return null;
}

describe("pptx:crop-picture", () => {
  it("writes srcRect on the target picture", async () => {
    const agent = await loadAgent("05-with-image.pptx");
    const slide = agent.getSnapshot().root.slides[0];
    const pic = findFirstPicture(slide.shapes);
    if (!pic) throw new Error("expected a picture in 05-with-image.pptx");

    const m = await agent.applyCommand({
      type: "pptx:crop-picture",
      payload: {
        slideIndex: 0,
        shapeId: pic.id,
        leftPct: 5,
        topPct: 10,
        rightPct: 15,
        bottomPct: 20,
      },
      source: "system",
    });
    expect(m.status).toBe("approved");

    const updated = findFirstPicture(agent.getSnapshot().root.slides[0].shapes);
    expect(updated?.srcRect).toEqual({
      leftPct: 5,
      topPct: 10,
      rightPct: 15,
      bottomPct: 20,
    });
  });

  it("survives a serialize → parse round-trip with 1000-multiplied attrs", async () => {
    const agent = await loadAgent("05-with-image.pptx");
    const slide = agent.getSnapshot().root.slides[0];
    const pic = findFirstPicture(slide.shapes);
    if (!pic) throw new Error("expected a picture");

    await agent.applyCommand({
      type: "pptx:crop-picture",
      payload: {
        slideIndex: 0,
        shapeId: pic.id,
        leftPct: 12.5,
        topPct: 7,
        rightPct: 25,
        bottomPct: 0,
      },
      source: "system",
    });

    const buf = await agent.exportFile();

    // Sniff the on-disk XML to confirm the 1000-multiplied attrs.
    const c = await ooxml.OoxmlContainer.load(buf);
    const xml = c.readText(slide.partPath);
    expect(xml).toContain("<a:srcRect");
    expect(xml).toMatch(/l="12500"/);
    expect(xml).toMatch(/t="7000"/);
    expect(xml).toMatch(/r="25000"/);
    // bottomPct is 0 so the attribute should be omitted (not `b="0"`)
    expect(xml).not.toMatch(/<a:srcRect[^>]*\bb="/);

    // Reload and confirm the typed field round-trips through the parser.
    const reloaded = await PptxAgent.fromBuffer(buf);
    const reloadedPic = findFirstPicture(reloaded.getSnapshot().root.slides[0].shapes);
    expect(reloadedPic?.srcRect).toEqual({
      leftPct: 12.5,
      topPct: 7,
      rightPct: 25,
      bottomPct: 0,
    });
  });

  it("clears the crop when all four sides are 0", async () => {
    const agent = await loadAgent("05-with-image.pptx");
    const slide = agent.getSnapshot().root.slides[0];
    const pic = findFirstPicture(slide.shapes);
    if (!pic) throw new Error("expected a picture");

    await agent.applyCommand({
      type: "pptx:crop-picture",
      payload: { slideIndex: 0, shapeId: pic.id, leftPct: 5, topPct: 5, rightPct: 5, bottomPct: 5 },
      source: "system",
    });
    expect(findFirstPicture(agent.getSnapshot().root.slides[0].shapes)?.srcRect).toBeDefined();

    await agent.applyCommand({
      type: "pptx:crop-picture",
      payload: { slideIndex: 0, shapeId: pic.id, leftPct: 0, topPct: 0, rightPct: 0, bottomPct: 0 },
      source: "system",
    });
    expect(findFirstPicture(agent.getSnapshot().root.slides[0].shapes)?.srcRect).toBeUndefined();

    // And the on-disk XML should no longer contain `<a:srcRect>`.
    const buf = await agent.exportFile();
    const c = await ooxml.OoxmlContainer.load(buf);
    const xml = c.readText(slide.partPath);
    expect(xml).not.toContain("<a:srcRect");
  });

  it("rejects out-of-range values", async () => {
    const agent = await loadAgent("05-with-image.pptx");
    const pic = findFirstPicture(agent.getSnapshot().root.slides[0].shapes);
    if (!pic) throw new Error("expected a picture");

    const negative = await agent.applyCommand({
      type: "pptx:crop-picture",
      payload: { slideIndex: 0, shapeId: pic.id, leftPct: -1, topPct: 0, rightPct: 0, bottomPct: 0 },
      source: "system",
    });
    expect(negative.status).toBe("rejected");
    expect(negative.rejection?.code).toBe("invalid-payload");
    expect(negative.rejection?.message).toMatch(/leftPct/);

    const overTen = await agent.applyCommand({
      type: "pptx:crop-picture",
      payload: { slideIndex: 0, shapeId: pic.id, leftPct: 0, topPct: 0, rightPct: 0, bottomPct: 101 },
      source: "system",
    });
    expect(overTen.status).toBe("rejected");
    expect(overTen.rejection?.message).toMatch(/bottomPct/);
  });

  it("rejects when leftPct + rightPct >= 100", async () => {
    const agent = await loadAgent("05-with-image.pptx");
    const pic = findFirstPicture(agent.getSnapshot().root.slides[0].shapes);
    if (!pic) throw new Error("expected a picture");

    const m = await agent.applyCommand({
      type: "pptx:crop-picture",
      payload: { slideIndex: 0, shapeId: pic.id, leftPct: 60, topPct: 0, rightPct: 40, bottomPct: 0 },
      source: "system",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-payload");
    expect(m.rejection?.message).toMatch(/leftPct \+ rightPct/);
  });

  it("rejects when topPct + bottomPct >= 100", async () => {
    const agent = await loadAgent("05-with-image.pptx");
    const pic = findFirstPicture(agent.getSnapshot().root.slides[0].shapes);
    if (!pic) throw new Error("expected a picture");

    const m = await agent.applyCommand({
      type: "pptx:crop-picture",
      payload: { slideIndex: 0, shapeId: pic.id, leftPct: 0, topPct: 50, rightPct: 0, bottomPct: 50 },
      source: "system",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.message).toMatch(/topPct \+ bottomPct/);
  });

  it("rejects non-picture targets", async () => {
    const agent = await loadAgent("04-multi-shape.pptx");
    const slide = agent.getSnapshot().root.slides[0];
    const nonPic = slide.shapes.find((s) => s.kind === "text");
    if (!nonPic) throw new Error("expected a text shape");

    const m = await agent.applyCommand({
      type: "pptx:crop-picture",
      payload: { slideIndex: 0, shapeId: nonPic.id, leftPct: 5, topPct: 5, rightPct: 5, bottomPct: 5 },
      source: "system",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-target");
  });
});
