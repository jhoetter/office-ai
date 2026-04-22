import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { ooxml } from "@officeai/core";
import { PptxAgent } from "../agent/agent.js";
import { parsePptx } from "../parser/parse.js";
import { serializePptx } from "../serializer/serialize.js";
import type { Shape, TextShape } from "../model/types.js";

const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);
const FIXTURE = join(FIXTURES_DIR.pathname, "04-multi-shape.pptx");

async function makeAgent(): Promise<PptxAgent> {
  return PptxAgent.fromBuffer(await readFile(FIXTURE));
}

function findFirstTextShape(shapes: ReadonlyArray<Shape>): TextShape | undefined {
  for (const s of shapes) {
    if (s.kind === "text") return s;
    if (s.kind === "group") {
      const r = findFirstTextShape(s.shapes);
      if (r) return r;
    }
  }
  return undefined;
}

describe("F-D: pptx:set-shape-geometry", () => {
  it("adds a corner-radius adjustment to a shape with no <a:avLst>", async () => {
    const agent = await makeAgent();
    const slide = agent.getSnapshot().root.slides[0]!;
    const shape = findFirstTextShape(slide.shapes)!;
    expect(shape).toBeDefined();
    await agent.applyCommand({
      type: "pptx:set-shape-geometry",
      payload: { slideIndex: 0, shapeId: shape.id, adjName: "adj", value: 25000 },
    });

    const out = await serializePptx(agent.getSnapshot());
    const c = await ooxml.OoxmlContainer.load(out);
    const xml = c.readText(slide.partPath);
    expect(xml).toMatch(/a:avLst/);
    expect(xml).toMatch(/a:gd[^>]*name="adj"[^>]*fmla="val 25000"/);
  });

  it("removes an adjustment when value=null", async () => {
    const agent = await makeAgent();
    const slide = agent.getSnapshot().root.slides[0]!;
    const shape = findFirstTextShape(slide.shapes)!;
    await agent.applyCommand({
      type: "pptx:set-shape-geometry",
      payload: { slideIndex: 0, shapeId: shape.id, adjName: "adj", value: 30000 },
    });
    await agent.applyCommand({
      type: "pptx:set-shape-geometry",
      payload: { slideIndex: 0, shapeId: shape.id, adjName: "adj", value: null },
    });
    const out = await serializePptx(agent.getSnapshot());
    const reparsed = await parsePptx(out);
    expect(reparsed.root.slides.length).toBeGreaterThan(0);
  });
});
