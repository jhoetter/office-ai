import { ooxml, sha256Hex, deterministicIdMinter } from "@officeai/core";
import { describe, expect, it } from "vitest";
import { DocxAgent } from "./agent.js";
import { parseDocx } from "../parser/parse.js";
import { makeSyntheticDocx, plainDocxXml } from "../test-utils/synthetic.js";

describe("DocxAgent end-to-end", () => {
  it("trivial-edit roundtrip preserves untouched parts byte-for-byte", async () => {
    const buf = await makeSyntheticDocx({
      documentXml: plainDocxXml([{ text: "Original heading", styleId: "Heading1" }, { text: "Body." }]),
    });
    const original = await ooxml.OoxmlContainer.load(buf);
    const originalHashes: Record<string, string> = {};
    for (const path of original.parts.keys()) {
      originalHashes[path] = sha256Hex(original.readBytes(path));
    }

    const agent = await DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    await agent.applyCommand({
      type: "docx:insert-text",
      payload: { at: { paragraph: 0, run: 0, offset: 0 }, text: "DRAFT — " },
      source: "human",
    });
    const out = await agent.exportFile();
    const reloaded = await ooxml.OoxmlContainer.load(out);

    for (const path of reloaded.parts.keys()) {
      if (path === "word/document.xml") continue; // expected to differ
      const before = originalHashes[path];
      const after = sha256Hex(reloaded.readBytes(path));
      expect(after, `untouched part ${path} should be byte-identical`).toBe(before);
    }
  });

  it("re-parses an exported edited file with the new text in place", async () => {
    const buf = await makeSyntheticDocx({
      documentXml: plainDocxXml([{ text: "old" }]),
    });
    const agent = await DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    await agent.applyCommand({
      type: "docx:insert-text",
      payload: { at: { paragraph: 0, run: 0, offset: 0 }, text: "very " },
      source: "human",
    });
    const out = await agent.exportFile();
    const reparsed = await parseDocx(out, { idMinter: deterministicIdMinter("x") });
    const p0 = reparsed.root.body[0];
    if (p0.kind !== "paragraph") throw new Error();
    let text = "";
    for (const c of p0.children) {
      if (c.kind === "run") {
        for (const ch of c.children) if (ch.kind === "text") text += ch.text;
      }
    }
    expect(text).toBe("very old");
  });

  it("toMarkdown projects headings + paragraphs", async () => {
    const buf = await makeSyntheticDocx({
      documentXml: plainDocxXml([
        { text: "Title here", styleId: "Heading1" },
        { text: "Some body content." },
        { text: "Bullet item", styleId: "ListParagraph" },
      ]),
    });
    const agent = await DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    const md = agent.toMarkdown();
    expect(md).toContain("# Title here");
    expect(md).toContain("Some body content.");
    expect(md).toContain("- Bullet item");
  });

  it("search finds substrings and returns paragraph indexes", async () => {
    const buf = await makeSyntheticDocx({
      documentXml: plainDocxXml([
        { text: "alpha beta" },
        { text: "gamma delta" },
        { text: "alpha epsilon" },
      ]),
    });
    const agent = await DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    const results = agent.search({ query: "alpha" });
    expect(results.map((r) => r.paragraphIndex)).toEqual([0, 2]);
  });

  it("export after add-comment includes the new comment in word/comments.xml", async () => {
    const buf = await makeSyntheticDocx({
      documentXml: plainDocxXml([{ text: "Please review this paragraph." }]),
    });
    const agent = await DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    await agent.applyCommand({
      type: "docx:add-comment",
      payload: {
        range: {
          start: { paragraph: 0, run: 0, offset: 0 },
          end: { paragraph: 0, run: 0, offset: 6 },
        },
        text: "rephrase?",
        author: "AI Reviewer",
        initials: "AI",
      },
      source: "human",
    });
    const out = await agent.exportFile();
    const c = await ooxml.OoxmlContainer.load(out);
    expect(c.has("word/comments.xml")).toBe(true);
    const xml = c.readText("word/comments.xml");
    expect(xml).toContain("rephrase?");
    expect(xml).toContain('w:author="AI Reviewer"');
    const reparsed = await parseDocx(out, { idMinter: deterministicIdMinter("y") });
    expect(reparsed.root.comments).toHaveLength(1);
    expect(reparsed.root.comments[0].author).toBe("AI Reviewer");
  });

  it("is fully headless (works without DOM globals)", async () => {
    expect(typeof globalThis.window).toBe("undefined");
    expect(typeof globalThis.document).toBe("undefined");
    const buf = await makeSyntheticDocx({ documentXml: plainDocxXml([{ text: "h" }]) });
    const agent = await DocxAgent.fromBuffer(buf);
    expect(agent.getSnapshot().format).toBe("docx");
  });
});
