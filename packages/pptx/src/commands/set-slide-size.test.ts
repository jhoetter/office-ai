import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { PptxAgent } from "../agent/agent.js";
import { SLIDE_SIZE_PRESETS } from "./set-slide-size.js";

const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);
const SIMPLE_FIXTURE = join(FIXTURES_DIR.pathname, "04-multi-shape.pptx");

async function loadAgent(): Promise<PptxAgent> {
  const buf = await readFile(SIMPLE_FIXTURE);
  return PptxAgent.fromBuffer(buf);
}

describe("pptx:set-slide-size", () => {
  it("applies the widescreen preset", async () => {
    const agent = await loadAgent();
    await agent.applyCommand({
      type: "pptx:set-slide-size",
      payload: { preset: "widescreen" },
    });

    const size = agent.getSnapshot().root.slideSize;
    expect(size.cxEmu).toBe(SLIDE_SIZE_PRESETS.widescreen.cxEmu);
    expect(size.cyEmu).toBe(SLIDE_SIZE_PRESETS.widescreen.cyEmu);
    expect(size.type).toBe("screen16x9");
  });

  it("applies the standard 4:3 preset", async () => {
    const agent = await loadAgent();
    await agent.applyCommand({
      type: "pptx:set-slide-size",
      payload: { preset: "standard" },
    });

    const size = agent.getSnapshot().root.slideSize;
    expect(size.cxEmu).toBe(SLIDE_SIZE_PRESETS.standard.cxEmu);
    expect(size.cyEmu).toBe(SLIDE_SIZE_PRESETS.standard.cyEmu);
    expect(size.type).toBe("screen4x3");
  });

  it("accepts custom EMU values", async () => {
    const agent = await loadAgent();
    await agent.applyCommand({
      type: "pptx:set-slide-size",
      payload: { preset: "custom", cxEmu: 12_000_000, cyEmu: 6_750_000 },
    });

    const size = agent.getSnapshot().root.slideSize;
    expect(size.cxEmu).toBe(12_000_000);
    expect(size.cyEmu).toBe(6_750_000);
  });

  it("is a no-op when the size is unchanged", async () => {
    const agent = await loadAgent();
    const before = agent.getSnapshot();
    await agent.applyCommand({
      type: "pptx:set-slide-size",
      payload: {
        preset: "custom",
        cxEmu: before.root.slideSize.cxEmu,
        cyEmu: before.root.slideSize.cyEmu,
        sizeType: before.root.slideSize.type,
      },
    });

    expect(agent.getSnapshot().revision).toBe(before.revision);
  });

  it("rejects custom preset without explicit cxEmu/cyEmu", async () => {
    const agent = await loadAgent();
    const result = await agent.applyCommand({
      type: "pptx:set-slide-size",
      payload: { preset: "custom" },
    });
    expect(result.status).toBe("rejected");
    if (result.status !== "rejected") return;
    expect(result.rejection.message).toMatch(/cxEmu and cyEmu/);
  });

  it("survives a save/reload round-trip", async () => {
    const agent = await loadAgent();
    await agent.applyCommand({
      type: "pptx:set-slide-size",
      payload: { preset: "standard" },
    });
    const bytes = await agent.exportFile();
    const reopened = await PptxAgent.fromBuffer(bytes);
    const size = reopened.getSnapshot().root.slideSize;
    expect(size.cxEmu).toBe(SLIDE_SIZE_PRESETS.standard.cxEmu);
    expect(size.cyEmu).toBe(SLIDE_SIZE_PRESETS.standard.cyEmu);
    expect(size.type).toBe("screen4x3");
  });
});
