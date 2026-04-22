import { describe, expect, it } from "vitest";
import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { PptxAgent } from "../agent/agent.js";
import type { Shape } from "../model/types.js";

const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);

async function loadAgent(name: string): Promise<PptxAgent> {
  const buf = await readFile(join(FIXTURES_DIR.pathname, name));
  return PptxAgent.fromBuffer(buf);
}

describe("pptx:paste-shapes", () => {
  it('appends pasted shapes with fresh ids and a default ¼" nudge', async () => {
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
      type: "pptx:paste-shapes",
      payload: { slideIndex: 0, shapes: [source] },
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

  it("appends multiple shapes in order with allocated cNvPrIds", async () => {
    const agent = await loadAgent("01-blank.pptx");
    await agent.applyCommand({
      type: "pptx:add-shape",
      payload: { slideIndex: 0, preset: "rect", x: 0, y: 0, width: 100_000, height: 100_000 },
      source: "human",
    });
    await agent.applyCommand({
      type: "pptx:add-shape",
      payload: { slideIndex: 0, preset: "ellipse", x: 0, y: 0, width: 100_000, height: 100_000 },
      source: "human",
    });
    const before = agent.getSnapshot().root.slides[0].shapes;
    const sources: Shape[] = [before[0], before[1]];

    const m = await agent.applyCommand({
      type: "pptx:paste-shapes",
      payload: { slideIndex: 0, shapes: sources, dxEmu: 500_000, dyEmu: 0 },
      source: "human",
    });
    expect(m.status).toBe("approved");

    const after = agent.getSnapshot().root.slides[0].shapes;
    expect(after.length).toBe(before.length + 2);
    expect(after[after.length - 2].kind).toBe("text"); // simple shapes are modelled as text shapes
    expect(after[after.length - 1].kind).toBe("text");
    expect(after[after.length - 2].cNvPrId).toBeLessThan(after[after.length - 1].cNvPrId);
    // Both clones offset by the supplied dx, retaining each source's y.
    expect(after[after.length - 2].position?.xEmu).toBe(500_000);
    expect(after[after.length - 1].position?.xEmu).toBe(500_000);
  });

  it("rejects an empty shape list", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const m = await agent.applyCommand({
      type: "pptx:paste-shapes",
      payload: { slideIndex: 0, shapes: [] },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    if (m.status !== "rejected") return;
    expect(m.rejection.code).toBe("not-applicable");
  });

  it("refuses to paste shapes whose kind references container parts", async () => {
    const agent = await loadAgent("01-blank.pptx");
    // Synthesise an opaque shape — `pic`/`chart`/etc. are also
    // refused but `opaque` is the easiest to construct without
    // wiring in a media part fixture.
    const fakeOpaque = {
      id: "op-1",
      cNvPrId: 99,
      name: "Opaque",
      kind: "opaque",
      raw: { tag: "p:sp", attrs: {}, children: [] },
    } as unknown as Shape;
    const m = await agent.applyCommand({
      type: "pptx:paste-shapes",
      payload: { slideIndex: 0, shapes: [fakeOpaque] },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    if (m.status !== "rejected") return;
    expect(m.rejection.code).toBe("not-applicable");
  });
});
