import { describe, expect, it } from "vitest";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { PptxAgent } from "../agent/agent.js";

const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);

async function loadAgent(name: string): Promise<PptxAgent> {
  const buf = await readFile(join(FIXTURES_DIR.pathname, name));
  return PptxAgent.fromBuffer(buf);
}

async function seedRect(
  agent: PptxAgent,
  rect: { x: number; y: number; cx: number; cy: number }
): Promise<string> {
  await agent.applyCommand({
    type: "pptx:add-shape",
    payload: {
      slideIndex: 0,
      preset: "rect",
      x: rect.x,
      y: rect.y,
      width: rect.cx,
      height: rect.cy,
    },
    source: "human",
  });
  const slide = agent.getSnapshot().root.slides[0];
  return slide.shapes[slide.shapes.length - 1]!.id;
}

describe("pptx:group-shapes / pptx:ungroup-shape", () => {
  it("groups two top-level shapes into a group with the union bounding box", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const a = await seedRect(agent, { x: 1_000_000, y: 500_000, cx: 1_000_000, cy: 500_000 });
    const b = await seedRect(agent, { x: 4_000_000, y: 1_500_000, cx: 2_000_000, cy: 1_000_000 });
    const beforeCount = agent.getSnapshot().root.slides[0].shapes.length;

    const m = await agent.applyCommand({
      type: "pptx:group-shapes",
      payload: { slideIndex: 0, shapeIds: [a, b] },
      source: "human",
    });
    expect(m.status).toBe("approved");

    const shapes = agent.getSnapshot().root.slides[0].shapes;
    expect(shapes.length).toBe(beforeCount - 1);
    const group = shapes[shapes.length - 1];
    expect(group.kind).toBe("group");
    if (group.kind !== "group") return;
    expect(group.children.length).toBe(2);
    expect(group.position?.xEmu).toBe(1_000_000);
    expect(group.position?.yEmu).toBe(500_000);
    expect(group.size?.cxEmu).toBe(5_000_000);
    expect(group.size?.cyEmu).toBe(2_000_000);
    expect(group.chOffExtRaw.length).toBe(2);
    const off = group.chOffExtRaw[0];
    const ext = group.chOffExtRaw[1];
    expect(off.tag).toBe("a:chOff");
    expect(off.attrs.x).toBe("1000000");
    expect(off.attrs.y).toBe("500000");
    expect(ext.tag).toBe("a:chExt");
    expect(ext.attrs.cx).toBe("5000000");
    expect(ext.attrs.cy).toBe("2000000");
  });

  it("ungroup re-inserts children at the group's slot, preserving order", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const a = await seedRect(agent, { x: 0, y: 0, cx: 1_000_000, cy: 1_000_000 });
    const b = await seedRect(agent, { x: 2_000_000, y: 0, cx: 1_000_000, cy: 1_000_000 });
    await agent.applyCommand({
      type: "pptx:group-shapes",
      payload: { slideIndex: 0, shapeIds: [a, b] },
      source: "human",
    });
    const groupId = agent.getSnapshot().root.slides[0].shapes.find((s) => s.kind === "group")!.id;

    const m = await agent.applyCommand({
      type: "pptx:ungroup-shape",
      payload: { slideIndex: 0, shapeId: groupId },
      source: "human",
    });
    expect(m.status).toBe("approved");

    const shapes = agent.getSnapshot().root.slides[0].shapes;
    expect(shapes.find((s) => s.id === groupId)).toBeUndefined();
    expect(shapes.find((s) => s.id === a)).toBeDefined();
    expect(shapes.find((s) => s.id === b)).toBeDefined();
  });

  it("rejects when fewer than 2 shape ids are provided", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const a = await seedRect(agent, { x: 0, y: 0, cx: 1_000_000, cy: 1_000_000 });
    const m = await agent.applyCommand({
      type: "pptx:group-shapes",
      payload: { slideIndex: 0, shapeIds: [a] },
      source: "human",
    });
    expect(m.status).toBe("rejected");
  });

  it("ungroup rejects non-group targets", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const a = await seedRect(agent, { x: 0, y: 0, cx: 1_000_000, cy: 1_000_000 });
    const m = await agent.applyCommand({
      type: "pptx:ungroup-shape",
      payload: { slideIndex: 0, shapeId: a },
      source: "human",
    });
    expect(m.status).toBe("rejected");
  });
});
