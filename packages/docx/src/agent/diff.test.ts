import { deterministicIdMinter } from "@officeai/core";
import { describe, expect, it } from "vitest";
import { DocxAgent } from "./agent.js";
import { diffDocxSnapshots } from "./diff.js";
import { makeSyntheticDocx, plainDocxXml } from "../test-utils/synthetic.js";

async function loadAgent(parts: ReadonlyArray<{ text: string; styleId?: string }>): Promise<DocxAgent> {
  const buf = await makeSyntheticDocx({ documentXml: plainDocxXml(parts) });
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

describe("DocxAgent.getDiff (structural snapshot diff)", () => {
  it("reports a paragraph text change", async () => {
    const a = await loadAgent([{ text: "Hello" }, { text: "World" }]);
    const before = a.getSnapshot();
    await a.applyCommand({
      type: "docx:insert-text",
      payload: { at: { paragraph: 1, run: 0, offset: 0 }, text: "Big " },
      source: "human",
    });
    const after = a.getSnapshot();
    const diff = a.getDiff(before, after);
    expect(diff.format).toBe("docx");
    expect(diff.changes.length).toBeGreaterThanOrEqual(1);
    const updates = diff.changes.filter((c) => c.kind === "node-updated" && c.field === "text");
    expect(updates).toHaveLength(1);
    expect(updates[0].summary).toMatch(/Big/);
  });

  it("reports a paragraph style change", async () => {
    const a = await loadAgent([{ text: "Heading-y" }, { text: "Body" }]);
    const before = a.getSnapshot();
    await a.applyCommand({
      type: "docx:set-paragraph-style",
      payload: { at: { paragraph: 0 }, style: "Heading2" },
      source: "human",
    });
    const after = a.getSnapshot();
    const diff = a.getDiff(before, after);
    const styleChanges = diff.changes.filter((c) => c.kind === "node-updated" && c.field === "styleId");
    expect(styleChanges).toHaveLength(1);
  });

  it("reports a new comment", async () => {
    const a = await loadAgent([{ text: "Please review this." }]);
    const before = a.getSnapshot();
    await a.applyCommand({
      type: "docx:add-comment",
      payload: {
        range: { start: { paragraph: 0, run: 0, offset: 0 }, end: { paragraph: 0, run: 0, offset: 6 } },
        text: "rephrase?",
        author: "AI",
        initials: "AI",
      },
      source: "human",
    });
    const after = a.getSnapshot();
    const diff = a.getDiff(before, after);
    const inserted = diff.changes.filter((c) => c.kind === "node-inserted" && c.path[0] === "comments");
    expect(inserted).toHaveLength(1);
    expect(inserted[0].summary).toContain("AI");
  });

  it("reports a resolved comment as a comment update", async () => {
    const a = await loadAgent([{ text: "Sentence." }]);
    await a.applyCommand({
      type: "docx:add-comment",
      payload: {
        range: { start: { paragraph: 0, run: 0, offset: 0 }, end: { paragraph: 0, run: 0, offset: 5 } },
        text: "?",
        author: "A",
      },
      source: "human",
    });
    const before = a.getSnapshot();
    const commentId = before.root.comments[0].id;
    await a.applyCommand({
      type: "docx:resolve-comment",
      payload: { commentId, resolved: true },
      source: "human",
    });
    const diff = a.getDiff(before, a.getSnapshot());
    const resolved = diff.changes.filter((c) => c.kind === "node-updated" && c.field === "resolved");
    expect(resolved).toHaveLength(1);
    expect(resolved[0].summary).toMatch(/resolved/);
  });

  it("returns an empty change list when nothing changed", async () => {
    const a = await loadAgent([{ text: "stable" }]);
    const snap = a.getSnapshot();
    const diff = diffDocxSnapshots(snap, snap);
    expect(diff.changes).toEqual([]);
  });
});

describe("DocxAgent.importFile", () => {
  it("replaces the in-memory document", async () => {
    const buf1 = await makeSyntheticDocx({ documentXml: plainDocxXml([{ text: "first doc" }]) });
    const buf2 = await makeSyntheticDocx({
      documentXml: plainDocxXml([{ text: "second doc" }, { text: "with two paragraphs" }]),
    });
    const a = await DocxAgent.fromBuffer(buf1, { idMinter: deterministicIdMinter() });
    expect(a.toMarkdown()).toContain("first doc");
    await a.importFile(buf2);
    expect(a.toMarkdown()).toContain("second doc");
    expect(a.toMarkdown()).toContain("with two paragraphs");
    expect(a.toMarkdown()).not.toContain("first doc");
  });

  it("drops pending mutations from the previous document", async () => {
    const buf1 = await makeSyntheticDocx({ documentXml: plainDocxXml([{ text: "alpha" }]) });
    const buf2 = await makeSyntheticDocx({ documentXml: plainDocxXml([{ text: "beta" }]) });
    const a = await DocxAgent.fromBuffer(buf1, { idMinter: deterministicIdMinter() });
    await a.applyCommand({
      type: "docx:insert-text",
      payload: { at: { paragraph: 0, run: 0, offset: 0 }, text: "X" },
      source: "agent",
      agentId: "test",
    });
    expect(a.getPendingMutations()).toHaveLength(1);
    await a.importFile(buf2);
    expect(a.getPendingMutations()).toHaveLength(0);
  });
});
