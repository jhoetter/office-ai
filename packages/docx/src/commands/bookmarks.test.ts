import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { DocxAgent } from "../agent/agent.js";
import { makeSyntheticDocx, plainDocxXml } from "../test-utils/synthetic.js";
import { listBookmarks } from "./bookmarks.js";

async function load(paragraphs: { text: string; styleId?: string }[]): Promise<DocxAgent> {
  const buf = await makeSyntheticDocx({ documentXml: plainDocxXml(paragraphs) });
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

describe("bookmarks", () => {
  it("inserts a named bookmark anchor pair around a range", async () => {
    const agent = await load([{ text: "Hello world" }]);
    const p0 = agent.getSnapshot().root.body[0];
    if (p0.kind !== "paragraph") throw new Error();
    const m = await agent.applyCommand({
      type: "docx:insert-bookmark",
      payload: { name: "Anchor1", paragraphId: p0.id, startOffset: 6, endOffset: 11 },
      source: "human",
    });
    expect(m.status).toBe("approved");

    const list = listBookmarks(agent.getSnapshot().root);
    expect(list).toHaveLength(1);
    expect(list[0]).toMatchObject({
      name: "Anchor1",
      paragraphId: p0.id,
      startOffset: 6,
      endOffset: 11,
    });
  });

  it("inserts a zero-length anchor at offset 0 when no selection", async () => {
    const agent = await load([{ text: "Hello world" }]);
    const p0 = agent.getSnapshot().root.body[0];
    if (p0.kind !== "paragraph") throw new Error();
    await agent.applyCommand({
      type: "docx:insert-bookmark",
      payload: { name: "Top", paragraphId: p0.id, startOffset: 0, endOffset: 0 },
      source: "human",
    });
    const list = listBookmarks(agent.getSnapshot().root);
    expect(list).toMatchObject([{ name: "Top", startOffset: 0, endOffset: 0 }]);
  });

  it("rejects invalid bookmark names", async () => {
    const agent = await load([{ text: "Hello world" }]);
    const p0 = agent.getSnapshot().root.body[0];
    if (p0.kind !== "paragraph") throw new Error();
    const m = await agent.applyCommand({
      type: "docx:insert-bookmark",
      payload: { name: "1bad name", paragraphId: p0.id, startOffset: 0, endOffset: 0 },
      source: "human",
    });
    expect(m.status).toBe("rejected");
  });

  it("delete-bookmark removes both anchors", async () => {
    const agent = await load([{ text: "Hello world" }]);
    const p0 = agent.getSnapshot().root.body[0];
    if (p0.kind !== "paragraph") throw new Error();
    await agent.applyCommand({
      type: "docx:insert-bookmark",
      payload: { name: "X", paragraphId: p0.id, startOffset: 0, endOffset: 5 },
      source: "human",
    });
    expect(listBookmarks(agent.getSnapshot().root)).toHaveLength(1);
    await agent.applyCommand({ type: "docx:delete-bookmark", payload: { name: "X" }, source: "human" });
    expect(listBookmarks(agent.getSnapshot().root)).toHaveLength(0);
  });

  it("re-inserting a bookmark with the same name moves it (overwrite)", async () => {
    const agent = await load([{ text: "Hello world" }]);
    const p0 = agent.getSnapshot().root.body[0];
    if (p0.kind !== "paragraph") throw new Error();
    await agent.applyCommand({
      type: "docx:insert-bookmark",
      payload: { name: "X", paragraphId: p0.id, startOffset: 0, endOffset: 5 },
      source: "human",
    });
    await agent.applyCommand({
      type: "docx:insert-bookmark",
      payload: { name: "X", paragraphId: p0.id, startOffset: 6, endOffset: 11 },
      source: "human",
    });
    const list = listBookmarks(agent.getSnapshot().root);
    expect(list).toHaveLength(1);
    expect(list[0]).toMatchObject({ name: "X", startOffset: 6, endOffset: 11 });
  });

  it("survives a serialize round-trip via opaque-inline", async () => {
    const agent = await load([{ text: "Hello world" }]);
    const p0 = agent.getSnapshot().root.body[0];
    if (p0.kind !== "paragraph") throw new Error();
    await agent.applyCommand({
      type: "docx:insert-bookmark",
      payload: { name: "Mid", paragraphId: p0.id, startOffset: 5, endOffset: 11 },
      source: "human",
    });
    const buf = await agent.exportFile();
    const reloaded = await DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    const list = listBookmarks(reloaded.getSnapshot().root);
    expect(list).toMatchObject([{ name: "Mid", startOffset: 5, endOffset: 11 }]);
  });
});
