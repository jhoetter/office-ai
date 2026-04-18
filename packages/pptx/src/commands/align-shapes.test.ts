import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { PptxAgent } from "../agent/agent.js";

const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);

async function loadAgent(name: string): Promise<PptxAgent> {
  const buf = await readFile(join(FIXTURES_DIR.pathname, name));
  return PptxAgent.fromBuffer(buf);
}

/**
 * Helper to seed N rectangles at known positions/sizes so we can assert
 * the post-alignment positions exactly. We always insert into the first
 * slide and return the created shape ids in insertion order.
 */
async function seedRects(
  agent: PptxAgent,
  rects: ReadonlyArray<{ x: number; y: number; cx: number; cy: number }>
): Promise<string[]> {
  const ids: string[] = [];
  for (const r of rects) {
    await agent.applyCommand({
      type: "pptx:add-shape",
      payload: {
        slideIndex: 0,
        preset: "rect",
        x: r.x,
        y: r.y,
        width: r.cx,
        height: r.cy,
      },
      source: "human",
    });
    const slide = agent.getSnapshot().root.slides[0];
    ids.push(slide.shapes[slide.shapes.length - 1]!.id);
  }
  return ids;
}

function boxOf(agent: PptxAgent, shapeId: string): { x: number; y: number; cx: number; cy: number } {
  const slide = agent.getSnapshot().root.slides[0];
  const s = slide.shapes.find((sh) => sh.id === shapeId)!;
  return { x: s.position!.xEmu, y: s.position!.yEmu, cx: s.size!.cxEmu, cy: s.size!.cyEmu };
}

describe("pptx:align-shapes", () => {
  it("align left snaps every shape's x to the union's leftmost x and leaves y untouched", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const ids = await seedRects(agent, [
      { x: 1_000_000, y: 500_000, cx: 1_000_000, cy: 500_000 },
      { x: 3_000_000, y: 1_500_000, cx: 1_000_000, cy: 500_000 },
      { x: 2_000_000, y: 2_500_000, cx: 1_000_000, cy: 500_000 },
    ]);

    const m = await agent.applyCommand({
      type: "pptx:align-shapes",
      payload: { slideIndex: 0, shapeIds: ids, mode: "left" },
      source: "human",
    });
    expect(m.status).toBe("approved");
    expect(boxOf(agent, ids[0]).x).toBe(1_000_000);
    expect(boxOf(agent, ids[1]).x).toBe(1_000_000);
    expect(boxOf(agent, ids[2]).x).toBe(1_000_000);
    expect(boxOf(agent, ids[1]).y).toBe(1_500_000);
  });

  it("align right snaps every shape's right edge to the union's right edge", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const ids = await seedRects(agent, [
      { x: 1_000_000, y: 500_000, cx: 1_000_000, cy: 500_000 },
      { x: 3_000_000, y: 1_500_000, cx: 2_000_000, cy: 500_000 }, // right = 5M
      { x: 2_000_000, y: 2_500_000, cx: 1_000_000, cy: 500_000 },
    ]);
    await agent.applyCommand({
      type: "pptx:align-shapes",
      payload: { slideIndex: 0, shapeIds: ids, mode: "right" },
      source: "human",
    });
    const a = boxOf(agent, ids[0]);
    const b = boxOf(agent, ids[1]);
    const c = boxOf(agent, ids[2]);
    expect(a.x + a.cx).toBe(5_000_000);
    expect(b.x + b.cx).toBe(5_000_000);
    expect(c.x + c.cx).toBe(5_000_000);
  });

  it("align center-h centres every shape on the union's x-center", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const ids = await seedRects(agent, [
      { x: 0, y: 0, cx: 1_000_000, cy: 500_000 }, // x range [0,1M], cx=1M
      { x: 4_000_000, y: 1_000_000, cx: 2_000_000, cy: 500_000 }, // x range [4M,6M]
    ]);
    await agent.applyCommand({
      type: "pptx:align-shapes",
      payload: { slideIndex: 0, shapeIds: ids, mode: "center-h" },
      source: "human",
    });
    // Union: [0, 6M], centre = 3M.
    // Shape A (cx=1M) -> x = 3M - 0.5M = 2.5M
    // Shape B (cx=2M) -> x = 3M - 1M = 2M
    expect(boxOf(agent, ids[0]).x).toBe(2_500_000);
    expect(boxOf(agent, ids[1]).x).toBe(2_000_000);
  });

  it("align middle-v matches centre-h on the y axis only", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const ids = await seedRects(agent, [
      { x: 0, y: 0, cx: 1_000_000, cy: 1_000_000 },
      { x: 2_000_000, y: 4_000_000, cx: 1_000_000, cy: 2_000_000 },
    ]);
    await agent.applyCommand({
      type: "pptx:align-shapes",
      payload: { slideIndex: 0, shapeIds: ids, mode: "middle-v" },
      source: "human",
    });
    // Union y range [0, 6M], centre = 3M.
    // Shape A (cy=1M) -> y = 3M - 0.5M = 2.5M
    // Shape B (cy=2M) -> y = 3M - 1M = 2M
    expect(boxOf(agent, ids[0]).y).toBe(2_500_000);
    expect(boxOf(agent, ids[1]).y).toBe(2_000_000);
  });

  it("rejects when fewer than two valid targets are supplied", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const ids = await seedRects(agent, [{ x: 0, y: 0, cx: 1_000_000, cy: 1_000_000 }]);
    const m = await agent.applyCommand({
      type: "pptx:align-shapes",
      payload: { slideIndex: 0, shapeIds: ids, mode: "left" },
      source: "human",
    });
    expect(m.status).toBe("rejected");
  });
});

describe("pptx:distribute-shapes", () => {
  it("evenly spaces shape centres along the horizontal axis, anchoring the extremes", async () => {
    const agent = await loadAgent("01-blank.pptx");
    // Three rectangles placed at x = 0, 1M, 6M (so middle is far too close to A).
    const ids = await seedRects(agent, [
      { x: 0, y: 0, cx: 1_000_000, cy: 500_000 }, // centre = 0.5M
      { x: 1_000_000, y: 0, cx: 1_000_000, cy: 500_000 }, // centre = 1.5M
      { x: 6_000_000, y: 0, cx: 1_000_000, cy: 500_000 }, // centre = 6.5M
    ]);
    await agent.applyCommand({
      type: "pptx:distribute-shapes",
      payload: { slideIndex: 0, shapeIds: ids, axis: "horizontal" },
      source: "human",
    });
    // First and last centre stay at 0.5M and 6.5M; mid centre = 3.5M
    // -> x = 3.5M - 0.5M = 3M.
    expect(boxOf(agent, ids[0]).x).toBe(0);
    expect(boxOf(agent, ids[2]).x).toBe(6_000_000);
    expect(boxOf(agent, ids[1]).x).toBe(3_000_000);
  });

  it("evenly spaces vertically, sorting by current centre even if input order differs", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const ids = await seedRects(agent, [
      { x: 0, y: 8_000_000, cx: 500_000, cy: 1_000_000 }, // centre = 8.5M
      { x: 0, y: 0, cx: 500_000, cy: 1_000_000 }, // centre = 0.5M
      { x: 0, y: 1_000_000, cx: 500_000, cy: 1_000_000 }, // centre = 1.5M
    ]);
    await agent.applyCommand({
      type: "pptx:distribute-shapes",
      payload: { slideIndex: 0, shapeIds: ids, axis: "vertical" },
      source: "human",
    });
    // Sorted by centre: ids[1] (0.5M), ids[2] (1.5M), ids[0] (8.5M).
    // Span = 8M, step = 4M. New centre for the middle one = 4.5M -> y = 4M.
    expect(boxOf(agent, ids[1]).y).toBe(0);
    expect(boxOf(agent, ids[0]).y).toBe(8_000_000);
    expect(boxOf(agent, ids[2]).y).toBe(4_000_000);
  });

  it("rejects when fewer than three valid targets are supplied", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const ids = await seedRects(agent, [
      { x: 0, y: 0, cx: 1_000_000, cy: 500_000 },
      { x: 2_000_000, y: 0, cx: 1_000_000, cy: 500_000 },
    ]);
    const m = await agent.applyCommand({
      type: "pptx:distribute-shapes",
      payload: { slideIndex: 0, shapeIds: ids, axis: "horizontal" },
      source: "human",
    });
    expect(m.status).toBe("rejected");
  });
});
