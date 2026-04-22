import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { ooxml } from "@officeai/core";
import { PptxAgent } from "../agent/agent.js";

const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);

async function loadAgent(name: string): Promise<PptxAgent> {
  const buf = await readFile(join(FIXTURES_DIR.pathname, name));
  return PptxAgent.fromBuffer(buf);
}

describe("pptx:set-slide-background", () => {
  it("inserts a solid <p:bg> and round-trips through the serializer", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const slide = agent.getSnapshot().root.slides[0];

    const m = await agent.applyCommand({
      type: "pptx:set-slide-background",
      payload: { slideIndex: 0, fill: { type: "solid", color: "112233" } },
      source: "human",
    });
    expect(m.status).toBe("approved");

    const updated = agent.getSnapshot().root.slides[0];
    expect(updated.cSldHead.some((c) => c.tag === "p:bg")).toBe(true);

    const out = await agent.exportFile();
    const c = await ooxml.OoxmlContainer.load(out);
    const xml = c.readText(slide.partPath);
    expect(xml.toUpperCase()).toContain("112233");
    expect(xml).toContain("<p:bg>");
    expect(xml).toContain("<a:solidFill>");
  });

  it("inserts a linear gradient <p:bg> with both stops", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const slide = agent.getSnapshot().root.slides[0];

    await agent.applyCommand({
      type: "pptx:set-slide-background",
      payload: {
        slideIndex: 0,
        fill: {
          type: "gradient",
          kind: "linear",
          angleDeg: 90,
          stops: [
            { pos: 0, color: "ff0000" },
            { pos: 1, color: "0000ff" },
          ],
        },
      },
      source: "human",
    });

    const out = await agent.exportFile();
    const c = await ooxml.OoxmlContainer.load(out);
    const xml = c.readText(slide.partPath);
    expect(xml).toContain("<a:gradFill");
    expect(xml).toContain("<a:gsLst>");
    expect(xml.toUpperCase()).toContain("FF0000");
    expect(xml.toUpperCase()).toContain("0000FF");
    expect(xml).toContain("<a:lin");
  });

  it("clearing the background removes <p:bg>", async () => {
    const agent = await loadAgent("01-blank.pptx");
    await agent.applyCommand({
      type: "pptx:set-slide-background",
      payload: { slideIndex: 0, fill: { type: "solid", color: "abcdef" } },
      source: "human",
    });
    expect(agent.getSnapshot().root.slides[0].cSldHead.some((c) => c.tag === "p:bg")).toBe(true);

    await agent.applyCommand({
      type: "pptx:set-slide-background",
      payload: { slideIndex: 0, fill: null },
      source: "human",
    });
    expect(agent.getSnapshot().root.slides[0].cSldHead.some((c) => c.tag === "p:bg")).toBe(false);
  });

  it("rejects malformed hex with invalid-payload", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const m = await agent.applyCommand({
      type: "pptx:set-slide-background",
      payload: { slideIndex: 0, fill: { type: "solid", color: "not-a-hex" } },
      source: "human",
    });
    expect(m.status).toBe("rejected");
  });
});

describe("pptx:set-shape-fill (FillSpec accepts gradient)", () => {
  it("accepts a typed gradient FillSpec and writes <a:gradFill>", async () => {
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
        fill: "AA0000",
      },
      source: "human",
    });
    const target = agent.getSnapshot().root.slides[0].shapes.slice(-1)[0]!;

    const m = await agent.applyCommand({
      type: "pptx:set-shape-fill",
      payload: {
        slideIndex: 0,
        shapeId: target.id,
        fill: {
          type: "gradient",
          kind: "linear",
          angleDeg: 45,
          stops: [
            { pos: 0, color: "ff0000" },
            { pos: 1, color: "00ff00" },
          ],
        },
      },
      source: "human",
    });
    expect(m.status).toBe("approved");

    const slide = agent.getSnapshot().root.slides[0];
    const out = await agent.exportFile();
    const c = await ooxml.OoxmlContainer.load(out);
    const xml = c.readText(slide.partPath);
    expect(xml).toContain("<a:gradFill");
    expect(xml.toUpperCase()).toContain("FF0000");
    expect(xml.toUpperCase()).toContain("00FF00");
    expect(xml).not.toContain("AA0000");
  });
});
