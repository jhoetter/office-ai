import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { PptxAgent } from "../agent/agent.js";

const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);
const SIMPLE_FIXTURE = join(FIXTURES_DIR.pathname, "04-multi-shape.pptx");

async function loadAgent(): Promise<PptxAgent> {
  const buf = await readFile(SIMPLE_FIXTURE);
  return PptxAgent.fromBuffer(buf);
}

describe("pptx:set-slide-hidden", () => {
  it("sets show=\"0\" on the targeted slide and round-trips", async () => {
    const agent = await loadAgent();
    expect(agent.getSnapshot().root.slides[0]!.slideRootAttrs.show).toBeUndefined();

    await agent.applyCommand({
      type: "pptx:set-slide-hidden",
      payload: { slideIndex: 0, hidden: true },
    });

    expect(agent.getSnapshot().root.slides[0]!.slideRootAttrs.show).toBe("0");
  });

  it("clears show=\"0\" when toggled back to visible", async () => {
    const agent = await loadAgent();
    await agent.applyCommand({
      type: "pptx:set-slide-hidden",
      payload: { slideIndex: 0, hidden: true },
    });
    await agent.applyCommand({
      type: "pptx:set-slide-hidden",
      payload: { slideIndex: 0, hidden: false },
    });

    expect(agent.getSnapshot().root.slides[0]!.slideRootAttrs.show).toBeUndefined();
  });

  it("is a no-op when toggling to the existing state", async () => {
    const agent = await loadAgent();
    const beforeRev = agent.getSnapshot().revision;
    await agent.applyCommand({
      type: "pptx:set-slide-hidden",
      payload: { slideIndex: 0, hidden: false },
    });
    expect(agent.getSnapshot().revision).toBe(beforeRev);
  });
});
