import { readFile } from "node:fs/promises";
import { join } from "node:path";
import { describe, expect, it } from "vitest";
import { ooxml } from "@officeai/core";
import { PptxAgent } from "../agent/agent.js";

const FIXTURES_DIR = new URL("../../../../fixtures/pptx/synthetic/", import.meta.url);

async function loadAgent(name: string): Promise<PptxAgent> {
  const buf = await readFile(join(FIXTURES_DIR.pathname, name));
  return PptxAgent.fromBuffer(buf);
}

describe("pptx comments lifecycle", () => {
  it("adds the first comment, synthesising the comments part + author registry", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const before = agent.getSnapshot();
    expect(before.root.commentAuthors).toBeNull();

    const m = await agent.applyCommand({
      type: "pptx:add-comment",
      payload: { slideIndex: 0, author: "Alice", text: "First note" },
      source: "human",
    });
    expect(m.status).toBe("approved");

    const after = agent.getSnapshot();
    expect(after.root.commentAuthors).not.toBeNull();
    expect(after.root.commentAuthors!.authors[0]!.name).toBe("Alice");

    const slide = after.root.slides[0];
    expect(slide.commentsPartPath).toBeTruthy();
    const part = after.root.commentsByPart.get(slide.commentsPartPath!)!;
    expect(part.comments).toHaveLength(1);
    expect(part.comments[0]!.text).toBe("First note");
    expect(part.comments[0]!.id).toBe("0:1");
  });

  it("threads replies under their parent and preserves them across resolve", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const m1 = await agent.applyCommand({
      type: "pptx:add-comment",
      payload: { slideIndex: 0, author: "Alice", text: "Top" },
      source: "human",
    });
    expect(m1.status).toBe("approved");
    const parentId = agent
      .getSnapshot()
      .root.commentsByPart.values()
      .next().value!.comments[0]!.id;

    const m2 = await agent.applyCommand({
      type: "pptx:reply-comment",
      payload: { slideIndex: 0, parentId, author: "Bob", text: "Reply" },
      source: "human",
    });
    expect(m2.status).toBe("approved");

    const after = agent.getSnapshot();
    const part = after.root.commentsByPart.values().next().value!;
    expect(part.comments).toHaveLength(2);
    const reply = part.comments.find((c: { parentId?: string }) => c.parentId === parentId)!;
    expect(reply.text).toBe("Reply");

    const m3 = await agent.applyCommand({
      type: "pptx:resolve-comment",
      payload: { slideIndex: 0, commentId: parentId, resolved: true },
      source: "human",
    });
    expect(m3.status).toBe("approved");
    const finalPart = agent.getSnapshot().root.commentsByPart.values().next().value!;
    expect(finalPart.comments.find((c: { id: string }) => c.id === parentId)!.resolved).toBe(true);
  });

  it("cascades delete: removing the parent drops its replies", async () => {
    const agent = await loadAgent("01-blank.pptx");
    await agent.applyCommand({
      type: "pptx:add-comment",
      payload: { slideIndex: 0, author: "Alice", text: "Top" },
      source: "human",
    });
    const parentId = agent
      .getSnapshot()
      .root.commentsByPart.values()
      .next().value!.comments[0]!.id;
    await agent.applyCommand({
      type: "pptx:reply-comment",
      payload: { slideIndex: 0, parentId, author: "Bob", text: "Reply" },
      source: "human",
    });

    await agent.applyCommand({
      type: "pptx:delete-comment",
      payload: { slideIndex: 0, commentId: parentId },
      source: "human",
    });
    const part = agent.getSnapshot().root.commentsByPart.values().next().value!;
    expect(part.comments).toHaveLength(0);
  });

  it("survives serialize → re-parse with comment + author + resolved flag intact", async () => {
    const agent = await loadAgent("01-blank.pptx");
    await agent.applyCommand({
      type: "pptx:add-comment",
      payload: { slideIndex: 0, author: "Alice", text: "Round-trip me" },
      source: "human",
    });
    const id = agent
      .getSnapshot()
      .root.commentsByPart.values()
      .next().value!.comments[0]!.id;
    await agent.applyCommand({
      type: "pptx:resolve-comment",
      payload: { slideIndex: 0, commentId: id, resolved: true },
      source: "human",
    });
    const out = await agent.exportFile();

    const c = await ooxml.OoxmlContainer.load(out);
    expect(c.has("ppt/commentAuthors.xml")).toBe(true);
    expect([...c.parts.keys()].some((p) => p.startsWith("ppt/comments/comment"))).toBe(true);

    const reparsed = await PptxAgent.fromBuffer(out);
    const part = reparsed
      .getSnapshot()
      .root.commentsByPart.values()
      .next().value!;
    expect(part.comments).toHaveLength(1);
    expect(part.comments[0]!.text).toBe("Round-trip me");
    expect(part.comments[0]!.resolved).toBe(true);
    expect(reparsed.getSnapshot().root.commentAuthors!.authors[0]!.name).toBe("Alice");
  });

  it("rejects replies to a non-existent parent", async () => {
    const agent = await loadAgent("01-blank.pptx");
    const m = await agent.applyCommand({
      type: "pptx:reply-comment",
      payload: { slideIndex: 0, parentId: "nope", author: "Alice", text: "x" },
      source: "human",
    });
    expect(m.status).toBe("rejected");
  });
});
