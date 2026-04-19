import { describe, expect, it } from "vitest";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { PptxAgent } from "../agent/agent.js";

const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);

async function loadAgent(name: string): Promise<PptxAgent> {
  const buf = await readFile(join(FIXTURES_DIR.pathname, name));
  return PptxAgent.fromBuffer(buf);
}

describe("pptx:reorder-shape", () => {
  it("brings a shape to the front and back", async () => {
    const agent = await loadAgent("04-multi-shape.pptx");
    const slide0 = agent.getSnapshot().root.slides[0];
    if (slide0.shapes.length < 2) return;
    const first = slide0.shapes[0];
    const lastIdx = slide0.shapes.length - 1;

    const m1 = await agent.applyCommand({
      type: "pptx:reorder-shape",
      payload: { slideIndex: 0, shapeId: first.id, mode: "to-front" },
      source: "human",
    });
    expect(m1.status).toBe("approved");
    const after1 = agent.getSnapshot().root.slides[0].shapes;
    expect(after1[lastIdx].id).toBe(first.id);

    const m2 = await agent.applyCommand({
      type: "pptx:reorder-shape",
      payload: { slideIndex: 0, shapeId: first.id, mode: "to-back" },
      source: "human",
    });
    expect(m2.status).toBe("approved");
    const after2 = agent.getSnapshot().root.slides[0].shapes;
    expect(after2[0].id).toBe(first.id);
  });

  it("forward swaps with the next shape; no-op at the front", async () => {
    const agent = await loadAgent("04-multi-shape.pptx");
    const slide0 = agent.getSnapshot().root.slides[0];
    if (slide0.shapes.length < 2) return;
    const first = slide0.shapes[0];
    const second = slide0.shapes[1];

    const m = await agent.applyCommand({
      type: "pptx:reorder-shape",
      payload: { slideIndex: 0, shapeId: first.id, mode: "forward" },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const after = agent.getSnapshot().root.slides[0].shapes;
    expect(after[0].id).toBe(second.id);
    expect(after[1].id).toBe(first.id);
  });

  it("rejects unknown shape ids", async () => {
    const agent = await loadAgent("04-multi-shape.pptx");
    const m = await agent.applyCommand({
      type: "pptx:reorder-shape",
      payload: { slideIndex: 0, shapeId: "does-not-exist", mode: "to-front" },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-target");
  });
});
