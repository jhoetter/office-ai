import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { DocxAgent } from "../agent/agent.js";
import { makeSyntheticDocx, plainDocxXml } from "../test-utils/synthetic.js";
import { paragraphPlainText } from "./helpers.js";

async function loadAgent(paragraphs: { text: string; styleId?: string }[]): Promise<DocxAgent> {
  const buf = await makeSyntheticDocx({ documentXml: plainDocxXml(paragraphs) });
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

describe("docx command handlers", () => {
  it("insert-text inserts at the start of a paragraph (no run targeting)", async () => {
    const agent = await loadAgent([{ text: "world" }]);
    const m = await agent.applyCommand({
      type: "docx:insert-text",
      payload: { at: { paragraph: 0 }, text: "hello " },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const p0 = agent.getSnapshot().root.body[0];
    if (p0.kind !== "paragraph") throw new Error();
    expect(paragraphPlainText(p0)).toBe("hello world");
  });

  it("insert-text splits a run at offset", async () => {
    const agent = await loadAgent([{ text: "abcdef" }]);
    await agent.applyCommand({
      type: "docx:insert-text",
      payload: { at: { paragraph: 0, run: 0, offset: 3 }, text: "X" },
      source: "human",
    });
    const p0 = agent.getSnapshot().root.body[0];
    if (p0.kind !== "paragraph") throw new Error();
    expect(paragraphPlainText(p0)).toBe("abcXdef");
  });

  it("delete-range removes characters within one paragraph", async () => {
    const agent = await loadAgent([{ text: "abcdefghij" }]);
    await agent.applyCommand({
      type: "docx:delete-range",
      payload: {
        range: {
          start: { paragraph: 0, run: 0, offset: 2 },
          end: { paragraph: 0, run: 0, offset: 6 },
        },
      },
      source: "human",
    });
    const p0 = agent.getSnapshot().root.body[0];
    if (p0.kind !== "paragraph") throw new Error();
    expect(paragraphPlainText(p0)).toBe("abghij");
  });

  it("format-range applies bold across a run subrange", async () => {
    const agent = await loadAgent([{ text: "alpha beta gamma" }]);
    await agent.applyCommand({
      type: "docx:format-range",
      payload: {
        range: {
          start: { paragraph: 0, run: 0, offset: 6 },
          end: { paragraph: 0, run: 0, offset: 10 },
        },
        format: { bold: true },
      },
      source: "human",
    });
    const p0 = agent.getSnapshot().root.body[0];
    if (p0.kind !== "paragraph") throw new Error();
    const runs = p0.children.filter((c) => c.kind === "run");
    const boldRun = runs.find((r) => r.kind === "run" && r.properties.bold === true);
    expect(boldRun).toBeTruthy();
    if (boldRun?.kind === "run") {
      const txt = boldRun.children
        .filter((c) => c.kind === "text")
        .map((c) => (c.kind === "text" ? c.text : ""))
        .join("");
      expect(txt).toBe("beta");
    }
  });

  it("insert-paragraph (offset 0) adds an empty paragraph BEFORE the index", async () => {
    const agent = await loadAgent([{ text: "first" }, { text: "second" }]);
    const before = agent.getSnapshot().root.body.length;
    await agent.applyCommand({
      type: "docx:insert-paragraph",
      payload: { at: { paragraph: 1, run: 0, offset: 0 }, style: "Heading2" },
      source: "human",
    });
    const body = agent.getSnapshot().root.body;
    expect(body.length).toBe(before + 1);
    const p1 = body[1];
    if (p1.kind !== "paragraph") throw new Error();
    expect(p1.properties.styleId).toBe("Heading2");
    expect(paragraphPlainText(p1)).toBe("");
    const p2 = body[2];
    if (p2.kind !== "paragraph") throw new Error();
    expect(paragraphPlainText(p2)).toBe("second");
  });

  it("insert-paragraph (mid-paragraph) splits the paragraph at offset (Enter semantics)", async () => {
    const agent = await loadAgent([{ text: "Hello world", styleId: "Heading1" }]);
    await agent.applyCommand({
      type: "docx:insert-paragraph",
      payload: { at: { paragraph: 0, run: 0, offset: 5 } },
      source: "human",
    });
    const paragraphs = agent.getSnapshot().root.body.filter((b) => b.kind === "paragraph");
    expect(paragraphs.length).toBe(2);
    const [p0, p1] = paragraphs;
    if (p0.kind !== "paragraph" || p1.kind !== "paragraph") throw new Error();
    expect(paragraphPlainText(p0)).toBe("Hello");
    expect(paragraphPlainText(p1)).toBe(" world");
    expect(p0.properties.styleId).toBe("Heading1");
    expect(p1.properties.styleId).toBe("Heading1");
  });

  it("insert-paragraph (offset >= text length) appends an empty paragraph after", async () => {
    const agent = await loadAgent([{ text: "tail" }]);
    await agent.applyCommand({
      type: "docx:insert-paragraph",
      payload: { at: { paragraph: 0, run: 0, offset: 99 } },
      source: "human",
    });
    const paragraphs = agent.getSnapshot().root.body.filter((b) => b.kind === "paragraph");
    expect(paragraphs.length).toBe(2);
    const [p0, p1] = paragraphs;
    if (p0.kind !== "paragraph" || p1.kind !== "paragraph") throw new Error();
    expect(paragraphPlainText(p0)).toBe("tail");
    expect(paragraphPlainText(p1)).toBe("");
  });

  it("set-paragraph-style updates styleId", async () => {
    const agent = await loadAgent([{ text: "Hello" }]);
    await agent.applyCommand({
      type: "docx:set-paragraph-style",
      payload: { at: { paragraph: 0 }, style: "Heading1" },
      source: "human",
    });
    const p0 = agent.getSnapshot().root.body[0];
    if (p0.kind !== "paragraph") throw new Error();
    expect(p0.properties.styleId).toBe("Heading1");
  });

  it("add-comment inserts comment markers and a comment record", async () => {
    const agent = await loadAgent([{ text: "Please review this." }]);
    const m = await agent.applyCommand({
      type: "docx:add-comment",
      payload: {
        range: {
          start: { paragraph: 0, run: 0, offset: 0 },
          end: { paragraph: 0, run: 0, offset: 5 },
        },
        text: "rephrase?",
        author: "AI",
        initials: "AI",
      },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const snap = agent.getSnapshot();
    expect(snap.root.comments).toHaveLength(1);
    expect(snap.root.comments[0].author).toBe("AI");
    const p0 = snap.root.body[0];
    if (p0.kind !== "paragraph") throw new Error();
    const kinds = p0.children.map((c) => c.kind);
    expect(kinds).toContain("comment-range-start");
    expect(kinds).toContain("comment-range-end");
    expect(kinds).toContain("comment-reference");
  });

  it("stub commands return a rejected mutation with not-implemented", async () => {
    const agent = await loadAgent([{ text: "x" }]);
    const m = await agent.applyCommand({
      type: "docx:insert-table",
      payload: { at: { paragraph: 0 }, rows: 2, cols: 2 },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("not-implemented");
  });

  it("agent command source produces pending mutation", async () => {
    const agent = await loadAgent([{ text: "draft" }]);
    const m = await agent.applyCommand({
      type: "docx:insert-text",
      payload: { at: { paragraph: 0 }, text: "[AI] " },
      source: "agent",
      agentId: "claude-1",
    });
    expect(m.status).toBe("pending");
    expect(agent.getPendingMutations()).toHaveLength(1);
    const p0 = agent.getSnapshot().root.body[0];
    if (p0.kind !== "paragraph") throw new Error();
    expect(paragraphPlainText(p0).startsWith("[AI] ")).toBe(true);
    agent.approveMutation(m.id);
    expect(agent.getPendingMutations()).toHaveLength(0);
    expect(agent.getApprovedSnapshot().root.body[0].kind).toBe("paragraph");
  });
});
