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

function findShowPr(agent: PptxAgent) {
  return agent
    .getSnapshot()
    .root.presentationOpaqueTail.find((o) => o.tag === "p:showPr");
}

describe("pptx:set-show-options", () => {
  it("writes a fresh <p:showPr> with attribute toggles", async () => {
    const agent = await loadAgent();
    expect(findShowPr(agent)).toBeUndefined();

    await agent.applyCommand({
      type: "pptx:set-show-options",
      payload: { loop: true, showNarration: true, showAnimation: false },
    });

    const showPr = findShowPr(agent);
    expect(showPr).toBeDefined();
    expect(showPr?.attrs.loop).toBe("1");
    expect(showPr?.attrs.showNarration).toBe("1");
    expect(showPr?.attrs.showAnimation).toBeUndefined();
  });

  it("emits <p:browse/> for showType=browse and replaces it for showType=kiosk", async () => {
    const agent = await loadAgent();
    await agent.applyCommand({
      type: "pptx:set-show-options",
      payload: { showType: "browse", loop: true },
    });
    let showPr = findShowPr(agent);
    expect(showPr?.subtree.some((c) => isElement(c, "p:browse"))).toBe(true);
    expect(showPr?.subtree.some((c) => isElement(c, "p:kiosk"))).toBe(false);

    await agent.applyCommand({
      type: "pptx:set-show-options",
      payload: { showType: "kiosk" },
    });
    showPr = findShowPr(agent);
    expect(showPr?.subtree.some((c) => isElement(c, "p:browse"))).toBe(false);
    expect(showPr?.subtree.some((c) => isElement(c, "p:kiosk"))).toBe(true);
    // Loop attribute persists across edits.
    expect(showPr?.attrs.loop).toBe("1");

    await agent.applyCommand({
      type: "pptx:set-show-options",
      payload: { showType: "presenter" },
    });
    showPr = findShowPr(agent);
    expect(showPr?.subtree.some((c) => isElement(c, "p:browse"))).toBe(false);
    expect(showPr?.subtree.some((c) => isElement(c, "p:kiosk"))).toBe(false);
  });

  it("clears the <p:showPr> element when clear:true", async () => {
    const agent = await loadAgent();
    await agent.applyCommand({
      type: "pptx:set-show-options",
      payload: { loop: true, showNarration: true },
    });
    expect(findShowPr(agent)).toBeDefined();

    await agent.applyCommand({
      type: "pptx:set-show-options",
      payload: { clear: true },
    });
    expect(findShowPr(agent)).toBeUndefined();
  });

  it("rejects an empty payload", async () => {
    const agent = await loadAgent();
    const result = await agent.applyCommand({
      type: "pptx:set-show-options",
      payload: {},
    });
    expect(result.status).toBe("rejected");
    if (result.status !== "rejected") return;
    expect(result.rejection.message).toMatch(/must supply/i);
  });

  it("survives a save/reload round-trip", async () => {
    const agent = await loadAgent();
    await agent.applyCommand({
      type: "pptx:set-show-options",
      payload: { showType: "kiosk", loop: true, useTimings: false },
    });

    const bytes = await agent.exportFile();
    const reopened = await PptxAgent.fromBuffer(bytes);
    const showPr = reopened
      .getSnapshot()
      .root.presentationOpaqueTail.find((o) => o.tag === "p:showPr");
    expect(showPr).toBeDefined();
    expect(showPr?.attrs.loop).toBe("1");
    expect(showPr?.attrs.useTimings).toBeUndefined();
    expect(showPr?.subtree.some((c) => isElement(c, "p:kiosk"))).toBe(true);
  });
});

function isElement(node: unknown, tag: string): boolean {
  if (!node || typeof node !== "object" || Array.isArray(node)) return false;
  const obj = node as Record<string, unknown>;
  const keys = Object.keys(obj).filter((k) => k !== ":@");
  return keys[0] === tag;
}
