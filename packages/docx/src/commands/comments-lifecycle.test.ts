import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { DocxAgent } from "../agent/agent.js";
import { parseDocx } from "../parser/parse.js";
import { makeSyntheticDocx, plainDocxXml } from "../test-utils/synthetic.js";

async function loadAgent(paragraphs: { text: string; styleId?: string }[]): Promise<DocxAgent> {
  const buf = await makeSyntheticDocx({ documentXml: plainDocxXml(paragraphs) });
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

async function addOneComment(
  agent: DocxAgent,
  text: string,
  author = "Reviewer",
  initials = "R"
): Promise<string> {
  await agent.applyCommand({
    type: "docx:add-comment",
    payload: {
      range: {
        start: { paragraph: 0, run: 0, offset: 0 },
        end: { paragraph: 0, run: 0, offset: 5 },
      },
      text,
      author,
      initials,
    },
    source: "human",
  });
  const snap = agent.getSnapshot();
  return snap.root.comments[snap.root.comments.length - 1].id;
}

describe("comment lifecycle commands", () => {
  describe("docx:resolve-comment", () => {
    it("marks the comment as resolved and dirties commentsExtended", async () => {
      const agent = await loadAgent([{ text: "Please review this." }]);
      const id = await addOneComment(agent, "rephrase?");
      const before = agent.getSnapshot();
      expect(before.root.comments[0].resolved).toBeUndefined();

      const m = await agent.applyCommand({
        type: "docx:resolve-comment",
        payload: { commentId: id },
        source: "human",
      });
      expect(m.status).toBe("approved");

      const after = agent.getSnapshot();
      expect(after.root.comments[0].resolved).toBe(true);
      expect(after.dirty.commentsExtended).toBe(true);
      expect(after.revision).toBe(before.revision + 1);
    });

    it("re-opens a resolved comment when resolved:false", async () => {
      const agent = await loadAgent([{ text: "Please review this." }]);
      const id = await addOneComment(agent, "rephrase?");
      await agent.applyCommand({
        type: "docx:resolve-comment",
        payload: { commentId: id },
        source: "human",
      });
      expect(agent.getSnapshot().root.comments[0].resolved).toBe(true);

      const m = await agent.applyCommand({
        type: "docx:resolve-comment",
        payload: { commentId: id, resolved: false },
        source: "human",
      });
      expect(m.status).toBe("approved");
      const c = agent.getSnapshot().root.comments[0];
      expect(c.resolved).toBeUndefined();
    });

    it("is a no-op when already in the requested state but still bumps revision", async () => {
      const agent = await loadAgent([{ text: "Please review this." }]);
      const id = await addOneComment(agent, "rephrase?");
      const before = agent.getSnapshot();
      const m = await agent.applyCommand({
        type: "docx:resolve-comment",
        payload: { commentId: id, resolved: false },
        source: "human",
      });
      expect(m.status).toBe("approved");
      const after = agent.getSnapshot();
      expect(after.root.comments[0].resolved).toBeUndefined();
      expect(after.revision).toBe(before.revision + 1);
    });

    it("rejects unknown comment ids", async () => {
      const agent = await loadAgent([{ text: "Please review this." }]);
      const m = await agent.applyCommand({
        type: "docx:resolve-comment",
        payload: { commentId: "does-not-exist" },
        source: "human",
      });
      expect(m.status).toBe("rejected");
      expect(m.rejection?.code).toBe("unknown-comment");
    });

    it("survives a serializer/parser round-trip with commentsExtended.xml", async () => {
      const agent = await loadAgent([{ text: "Please review this." }]);
      const id = await addOneComment(agent, "rephrase?");
      await agent.applyCommand({
        type: "docx:resolve-comment",
        payload: { commentId: id },
        source: "human",
      });
      const buf = await agent.exportFile();
      const reparsed = await parseDocx(buf);
      const reC = reparsed.root.comments.find((c) => c.id === id);
      expect(reC).toBeTruthy();
      expect(reC?.resolved).toBe(true);
    });
  });

  describe("docx:reply-comment", () => {
    it("appends a reply with parentId and a fresh id", async () => {
      const agent = await loadAgent([{ text: "Please review this." }]);
      const parentId = await addOneComment(agent, "rephrase?", "Alice", "A");
      const m = await agent.applyCommand({
        type: "docx:reply-comment",
        payload: {
          parentId,
          text: "I will rewrite the intro.",
          author: "Bob",
          initials: "B",
        },
        source: "human",
      });
      expect(m.status).toBe("approved");

      const snap = agent.getSnapshot();
      expect(snap.root.comments).toHaveLength(2);
      const reply = snap.root.comments[1];
      expect(reply.id).not.toBe(parentId);
      expect(reply.parentId).toBe(parentId);
      expect(reply.author).toBe("Bob");
      expect(reply.body).toHaveLength(1);
      expect(snap.dirty.commentsExtended).toBe(true);
      expect(snap.dirty.comments).toBe(true);
    });

    it("does not add new commentRange markers in the body", async () => {
      const agent = await loadAgent([{ text: "Please review this." }]);
      const parentId = await addOneComment(agent, "rephrase?");
      const beforeBody = agent.getSnapshot().root.body[0];
      if (beforeBody.kind !== "paragraph") throw new Error();
      const beforeMarkers = beforeBody.children.filter(
        (c) =>
          c.kind === "comment-range-start" || c.kind === "comment-range-end" || c.kind === "comment-reference"
      ).length;

      await agent.applyCommand({
        type: "docx:reply-comment",
        payload: { parentId, text: "ack", author: "Bob" },
        source: "human",
      });

      const afterBody = agent.getSnapshot().root.body[0];
      if (afterBody.kind !== "paragraph") throw new Error();
      const afterMarkers = afterBody.children.filter(
        (c) =>
          c.kind === "comment-range-start" || c.kind === "comment-range-end" || c.kind === "comment-reference"
      ).length;
      expect(afterMarkers).toBe(beforeMarkers);
    });

    it("rejects when parent is unknown", async () => {
      const agent = await loadAgent([{ text: "Please review this." }]);
      const m = await agent.applyCommand({
        type: "docx:reply-comment",
        payload: { parentId: "999", text: "hi", author: "Bob" },
        source: "human",
      });
      expect(m.status).toBe("rejected");
      expect(m.rejection?.code).toBe("unknown-comment");
    });

    it("rejects empty replies", async () => {
      const agent = await loadAgent([{ text: "Please review this." }]);
      const parentId = await addOneComment(agent, "rephrase?");
      const m = await agent.applyCommand({
        type: "docx:reply-comment",
        payload: { parentId, text: "", author: "Bob" },
        source: "human",
      });
      expect(m.status).toBe("rejected");
      expect(m.rejection?.code).toBe("empty-reply");
    });

    it("survives a round-trip and preserves parentId via commentsExtended.xml", async () => {
      const agent = await loadAgent([{ text: "Please review this." }]);
      const parentId = await addOneComment(agent, "rephrase?");
      await agent.applyCommand({
        type: "docx:reply-comment",
        payload: { parentId, text: "ack", author: "Bob" },
        source: "human",
      });
      const buf = await agent.exportFile();
      const reparsed = await parseDocx(buf);
      expect(reparsed.root.comments).toHaveLength(2);
      const reply = reparsed.root.comments.find((c) => c.id !== parentId);
      expect(reply?.parentId).toBe(parentId);
    });
  });

  describe("docx:delete-comment", () => {
    it("removes the comment and its inline range markers", async () => {
      const agent = await loadAgent([{ text: "Please review this." }]);
      const id = await addOneComment(agent, "rephrase?");
      const m = await agent.applyCommand({
        type: "docx:delete-comment",
        payload: { commentId: id },
        source: "human",
      });
      expect(m.status).toBe("approved");
      const snap = agent.getSnapshot();
      expect(snap.root.comments).toHaveLength(0);
      const p0 = snap.root.body[0];
      if (p0.kind !== "paragraph") throw new Error();
      const markerKinds = new Set(["comment-range-start", "comment-range-end", "comment-reference"]);
      expect(p0.children.every((c) => !markerKinds.has(c.kind))).toBe(true);
      expect(snap.dirty.body).toBe(true);
      expect(snap.dirty.comments).toBe(true);
      expect(snap.dirty.rels).toBe(true);
      expect(snap.dirty.contentTypes).toBe(true);
    });

    it("deleting the parent removes the entire reply thread", async () => {
      const agent = await loadAgent([{ text: "Please review this." }]);
      const parentId = await addOneComment(agent, "rephrase?", "Alice", "A");
      await agent.applyCommand({
        type: "docx:reply-comment",
        payload: { parentId, text: "ack", author: "Bob" },
        source: "human",
      });
      await agent.applyCommand({
        type: "docx:reply-comment",
        payload: { parentId, text: "also", author: "Carol" },
        source: "human",
      });
      expect(agent.getSnapshot().root.comments).toHaveLength(3);

      await agent.applyCommand({
        type: "docx:delete-comment",
        payload: { commentId: parentId },
        source: "human",
      });
      expect(agent.getSnapshot().root.comments).toHaveLength(0);
    });

    it("deleting a reply leaves the parent intact", async () => {
      const agent = await loadAgent([{ text: "Please review this." }]);
      const parentId = await addOneComment(agent, "rephrase?", "Alice", "A");
      await agent.applyCommand({
        type: "docx:reply-comment",
        payload: { parentId, text: "ack", author: "Bob" },
        source: "human",
      });
      const replyId = agent.getSnapshot().root.comments[1].id;
      await agent.applyCommand({
        type: "docx:delete-comment",
        payload: { commentId: replyId },
        source: "human",
      });
      const snap = agent.getSnapshot();
      expect(snap.root.comments).toHaveLength(1);
      expect(snap.root.comments[0].id).toBe(parentId);
    });

    it("rejects unknown ids", async () => {
      const agent = await loadAgent([{ text: "Please review this." }]);
      const m = await agent.applyCommand({
        type: "docx:delete-comment",
        payload: { commentId: "nope" },
        source: "human",
      });
      expect(m.status).toBe("rejected");
      expect(m.rejection?.code).toBe("unknown-comment");
    });

    it("survives a round-trip — parts are gone after save+reparse", async () => {
      const agent = await loadAgent([{ text: "Please review this." }]);
      const id = await addOneComment(agent, "rephrase?");
      await agent.applyCommand({
        type: "docx:delete-comment",
        payload: { commentId: id },
        source: "human",
      });
      const buf = await agent.exportFile();
      const reparsed = await parseDocx(buf);
      expect(reparsed.root.comments).toHaveLength(0);
    });
  });
});
