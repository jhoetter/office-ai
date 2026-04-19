import { describe, expect, it } from "vitest";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { PptxAgent } from "../agent/agent.js";

const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);

async function loadAgent(name: string): Promise<PptxAgent> {
  const buf = await readFile(join(FIXTURES_DIR.pathname, name));
  return PptxAgent.fromBuffer(buf);
}

describe("pptx:duplicate-shape", () => {
  it('appends a clone with a fresh cNvPrId and a ¼" diagonal nudge', async () => {
    const agent = await loadAgent("01-blank.pptx");
    await agent.applyCommand({
      type: "pptx:add-shape",
      payload: {
        slideIndex: 0,
        preset: "rect",
        x: 1_000_000,
        y: 500_000,
        width: 1_000_000,
        height: 500_000,
      },
      source: "human",
    });
    const before = agent.getSnapshot().root.slides[0].shapes;
    const source = before[before.length - 1];

    const m = await agent.applyCommand({
      type: "pptx:duplicate-shape",
      payload: { slideIndex: 0, shapeId: source.id },
      source: "human",
    });
    expect(m.status).toBe("approved");

    const after = agent.getSnapshot().root.slides[0].shapes;
    expect(after.length).toBe(before.length + 1);
    const clone = after[after.length - 1];
    expect(clone.id).not.toBe(source.id);
    expect(clone.cNvPrId).toBeGreaterThan(source.cNvPrId);
    expect(clone.kind).toBe(source.kind);
    expect(clone.position?.xEmu).toBe((source.position?.xEmu ?? 0) + 228_600);
    expect(clone.position?.yEmu).toBe((source.position?.yEmu ?? 0) + 228_600);
  });

  it("honours custom dx/dy offsets", async () => {
    const agent = await loadAgent("01-blank.pptx");
    await agent.applyCommand({
      type: "pptx:add-shape",
      payload: {
        slideIndex: 0,
        preset: "ellipse",
        x: 0,
        y: 0,
        width: 1_000_000,
        height: 500_000,
      },
      source: "human",
    });
    const source = agent.getSnapshot().root.slides[0].shapes[0];
    await agent.applyCommand({
      type: "pptx:duplicate-shape",
      payload: { slideIndex: 0, shapeId: source.id, dxEmu: 0, dyEmu: 1_000_000 },
      source: "human",
    });
    const clone = agent.getSnapshot().root.slides[0].shapes[1];
    expect(clone.position?.xEmu).toBe(0);
    expect(clone.position?.yEmu).toBe(1_000_000);
  });

  it("rejects unknown shape ids", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const m = await agent.applyCommand({
      type: "pptx:duplicate-shape",
      payload: { slideIndex: 0, shapeId: "missing" },
      source: "human",
    });
    expect(m.status).toBe("rejected");
  });
});
