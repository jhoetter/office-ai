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

  it("format-range across three paragraphs leaves untouched runs untouched and dirty-flags the body", async () => {
    const agent = await loadAgent([{ text: "alpha" }, { text: "beta" }, { text: "gamma" }]);
    const m = await agent.applyCommand({
      type: "docx:format-range",
      payload: {
        range: {
          start: { paragraph: 0, offset: 2 },
          end: { paragraph: 2, offset: 3 },
        },
        format: { bold: true },
      },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const snap = agent.getSnapshot();
    expect(snap.dirty.body).toBe(true);

    const collectBoldText = (idx: number): string => {
      const p = snap.root.body[idx];
      if (p.kind !== "paragraph") throw new Error();
      let out = "";
      for (const c of p.children) {
        if (c.kind !== "run") continue;
        if (c.properties.bold !== true) continue;
        for (const child of c.children) {
          if (child.kind === "text") out += child.text;
        }
      }
      return out;
    };
    expect(collectBoldText(0)).toBe("pha"); // tail of "alpha" from offset 2
    expect(collectBoldText(1)).toBe("beta"); // entire intermediate paragraph
    expect(collectBoldText(2)).toBe("gam"); // head of "gamma" up to offset 3

    // Untouched plain text still exists.
    const p0 = snap.root.body[0];
    if (p0.kind !== "paragraph") throw new Error();
    expect(paragraphPlainText(p0)).toBe("alpha");
  });

  it("delete-range across three paragraphs merges start+end and drops intermediates", async () => {
    const agent = await loadAgent([{ text: "alpha" }, { text: "middle" }, { text: "omega tail" }]);
    const before = agent.getSnapshot().root.body.length;
    const startId = (() => {
      const b = agent.getSnapshot().root.body[0];
      if (b.kind !== "paragraph") throw new Error();
      return b.id;
    })();
    await agent.applyCommand({
      type: "docx:delete-range",
      payload: {
        range: {
          start: { paragraph: 0, offset: 3 },
          end: { paragraph: 2, offset: 6 },
        },
      },
      source: "human",
    });
    const body = agent.getSnapshot().root.body;
    expect(body.length).toBe(before - 2);
    const merged = body[0];
    if (merged.kind !== "paragraph") throw new Error();
    expect(merged.id).toBe(startId);
    // "alpha" trimmed at offset 3 → "alp"; "omega tail" trimmed before
    // offset 6 → "tail"; merged = "alp" + "tail" = "alptail".
    expect(paragraphPlainText(merged)).toBe("alptail");
  });

  it("delete-range that empties the merged paragraph still leaves a well-formed paragraph", async () => {
    const agent = await loadAgent([{ text: "alpha" }, { text: "beta" }]);
    await agent.applyCommand({
      type: "docx:delete-range",
      payload: {
        range: {
          start: { paragraph: 0, offset: 0 },
          end: { paragraph: 1, offset: 4 },
        },
      },
      source: "human",
    });
    const body = agent.getSnapshot().root.body;
    const paragraphs = body.filter((b) => b.kind === "paragraph");
    expect(paragraphs.length).toBe(1);
    const merged = paragraphs[0];
    if (merged.kind !== "paragraph") throw new Error();
    expect(paragraphPlainText(merged)).toBe("");
    // Must still contain at least one (empty) run so the renderer/serializer
    // produce a well-formed <w:p>.
    const runs = merged.children.filter((c) => c.kind === "run");
    expect(runs.length).toBeGreaterThanOrEqual(1);
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
    // `docx:insert-image` is the last remaining stub after P1.3 / W7 shipped
    // the four typed-table commands. Repointed by W7.
    const m = await agent.applyCommand({
      type: "docx:insert-image",
      payload: {
        at: { paragraph: 0 },
        data: new ArrayBuffer(0),
        mimeType: "image/png",
        width: 1,
        height: 1,
      },
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
