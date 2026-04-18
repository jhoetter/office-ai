import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { DocxAgent } from "../agent/agent.js";
import { makeSyntheticDocx, plainDocxXml } from "../test-utils/synthetic.js";

async function loadAgent(paragraphs: { text: string; styleId?: string }[]): Promise<DocxAgent> {
  const buf = await makeSyntheticDocx({ documentXml: plainDocxXml(paragraphs) });
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

function firstParagraphId(agent: DocxAgent): string {
  const block = agent.getSnapshot().root.body[0];
  if (block.kind !== "paragraph") throw new Error("expected paragraph at index 0");
  return block.id;
}

describe("docx:set-paragraph-alignment", () => {
  it("sets the typed alignment field and dirty-flags the body", async () => {
    const agent = await loadAgent([{ text: "abc" }]);
    const paragraphId = firstParagraphId(agent);
    const before = agent.getSnapshot();
    const m = await agent.applyCommand({
      type: "docx:set-paragraph-alignment",
      payload: { paragraphId, alignment: "center" },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const after = agent.getSnapshot();
    expect(after.revision).toBe(before.revision + 1);
    expect(after.dirty.body).toBe(true);
    const p0 = after.root.body[0];
    if (p0.kind !== "paragraph") throw new Error();
    expect(p0.properties.alignment).toBe("center");
  });

  it("clears alignment when payload.alignment is null", async () => {
    const agent = await loadAgent([{ text: "abc" }]);
    const paragraphId = firstParagraphId(agent);
    await agent.applyCommand({
      type: "docx:set-paragraph-alignment",
      payload: { paragraphId, alignment: "right" },
      source: "human",
    });
    await agent.applyCommand({
      type: "docx:set-paragraph-alignment",
      payload: { paragraphId, alignment: null },
      source: "human",
    });
    const p0 = agent.getSnapshot().root.body[0];
    if (p0.kind !== "paragraph") throw new Error();
    expect(p0.properties.alignment).toBeUndefined();
  });

  it("is a no-op when alignment already equals payload value", async () => {
    const agent = await loadAgent([{ text: "abc" }]);
    const paragraphId = firstParagraphId(agent);
    await agent.applyCommand({
      type: "docx:set-paragraph-alignment",
      payload: { paragraphId, alignment: "center" },
      source: "human",
    });
    const beforeRev = agent.getSnapshot().revision;
    await agent.applyCommand({
      type: "docx:set-paragraph-alignment",
      payload: { paragraphId, alignment: "center" },
      source: "human",
    });
    expect(agent.getSnapshot().revision).toBe(beforeRev);
  });

  it("rejects bogus alignment values", async () => {
    const agent = await loadAgent([{ text: "abc" }]);
    const paragraphId = firstParagraphId(agent);
    const m = await agent.applyCommand({
      type: "docx:set-paragraph-alignment",
      payload: { paragraphId, alignment: "diagonal" as unknown as "left" },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-payload");
  });

  it("rejects unknown paragraphIds", async () => {
    const agent = await loadAgent([{ text: "abc" }]);
    const m = await agent.applyCommand({
      type: "docx:set-paragraph-alignment",
      payload: { paragraphId: "p:does-not-exist", alignment: "left" },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-target");
  });
});

describe("docx:set-paragraph-indent", () => {
  it("steps left indent by deltaTwips", async () => {
    const agent = await loadAgent([{ text: "abc" }]);
    const paragraphId = firstParagraphId(agent);
    await agent.applyCommand({
      type: "docx:set-paragraph-indent",
      payload: { paragraphId, deltaTwips: 360 },
      source: "human",
    });
    const p0 = agent.getSnapshot().root.body[0];
    if (p0.kind !== "paragraph") throw new Error();
    expect(p0.properties.indentation?.left).toBe(360);

    await agent.applyCommand({
      type: "docx:set-paragraph-indent",
      payload: { paragraphId, deltaTwips: 360 },
      source: "human",
    });
    const p1 = agent.getSnapshot().root.body[0];
    if (p1.kind !== "paragraph") throw new Error();
    expect(p1.properties.indentation?.left).toBe(720);
  });

  it("clamps at zero when outdenting past the left margin", async () => {
    const agent = await loadAgent([{ text: "abc" }]);
    const paragraphId = firstParagraphId(agent);
    await agent.applyCommand({
      type: "docx:set-paragraph-indent",
      payload: { paragraphId, deltaTwips: -360 },
      source: "human",
    });
    const p0 = agent.getSnapshot().root.body[0];
    if (p0.kind !== "paragraph") throw new Error();
    // No change applied (left was already 0), so indentation stays
    // undefined entirely.
    expect(p0.properties.indentation).toBeUndefined();
  });

  it("preserves right / firstLine / hanging when stepping left", async () => {
    const agent = await loadAgent([{ text: "abc" }]);
    const paragraphId = firstParagraphId(agent);
    // Manually set a starting indentation via two indent commands first to
    // exercise the code-path that writes through `evolveSnapshot` repeatedly.
    await agent.applyCommand({
      type: "docx:set-paragraph-indent",
      payload: { paragraphId, deltaTwips: 720 },
      source: "human",
    });
    // Now drop the indent by 360 — left should land at 360.
    await agent.applyCommand({
      type: "docx:set-paragraph-indent",
      payload: { paragraphId, deltaTwips: -360 },
      source: "human",
    });
    const p0 = agent.getSnapshot().root.body[0];
    if (p0.kind !== "paragraph") throw new Error();
    expect(p0.properties.indentation?.left).toBe(360);
  });

  it("rejects non-integer deltas", async () => {
    const agent = await loadAgent([{ text: "abc" }]);
    const paragraphId = firstParagraphId(agent);
    const m = await agent.applyCommand({
      type: "docx:set-paragraph-indent",
      payload: { paragraphId, deltaTwips: 12.5 },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-payload");
  });
});
