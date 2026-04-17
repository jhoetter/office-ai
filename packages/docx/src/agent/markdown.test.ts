import { describe, expect, it, vi } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { DocxAgent } from "./agent.js";
import { snapshotToMarkdown } from "./markdown.js";
import { DEFAULT_DOC_ROOT_ATTRS, makeSyntheticDocx, plainDocxXml } from "../test-utils/synthetic.js";

const HEADER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>`;
const FOOTER = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;

async function buildAgent(bodyXml: string): Promise<DocxAgent> {
  const documentXml = `${HEADER}${bodyXml}${FOOTER}`;
  const buf = await makeSyntheticDocx({ documentXml });
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

describe("snapshotToMarkdown — extended projection", () => {
  it("projects every Heading[1-6] / Title style to the matching ATX heading", async () => {
    const buf = await makeSyntheticDocx({
      documentXml: plainDocxXml([
        { text: "Top", styleId: "Title" },
        { text: "Big", styleId: "Heading1" },
        { text: "Mid", styleId: "Heading2" },
        { text: "Smaller", styleId: "Heading3" },
        { text: "Even smaller", styleId: "Heading4" },
        { text: "Smallest visible", styleId: "Heading5" },
        { text: "Tiny", styleId: "Heading6" },
        { text: "Body text" },
      ]),
    });
    const agent = await DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    const md = snapshotToMarkdown(agent.getSnapshot());
    expect(md).toContain("# Top");
    expect(md).toContain("# Big");
    expect(md).toContain("## Mid");
    expect(md).toContain("### Smaller");
    expect(md).toContain("#### Even smaller");
    expect(md).toContain("##### Smallest visible");
    expect(md).toContain("###### Tiny");
    expect(md).toContain("Body text");
    expect(md).not.toContain("####### "); // we never go past h6
  });

  it("renders bullet vs numbered list paragraphs with ilvl indentation", async () => {
    const agent = await buildAgent(`
      <w:p><w:pPr><w:pStyle w:val="ListParagraph"/></w:pPr><w:r><w:t xml:space="preserve">A bare bullet</w:t></w:r></w:p>
      <w:p><w:pPr>
        <w:pStyle w:val="ListParagraph"/>
        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
      </w:pPr><w:r><w:t xml:space="preserve">First numbered</w:t></w:r></w:p>
      <w:p><w:pPr>
        <w:pStyle w:val="ListParagraph"/>
        <w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr>
      </w:pPr><w:r><w:t xml:space="preserve">Indented numbered</w:t></w:r></w:p>
      <w:p><w:pPr>
        <w:pStyle w:val="ListParagraph"/>
        <w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>
      </w:pPr><w:r><w:t xml:space="preserve">Back to top level</w:t></w:r></w:p>
    `);
    const md = snapshotToMarkdown(agent.getSnapshot());
    expect(md).toContain("- A bare bullet");
    expect(md).toContain("1. First numbered");
    expect(md).toContain("  1. Indented numbered");
    expect(md).toContain("1. Back to top level");
  });

  it("turns a w:tbl into a GFM pipe table with header row", async () => {
    const agent = await buildAgent(`
      <w:tbl>
        <w:tblPr><w:tblW w:type="pct" w:w="100%"/></w:tblPr>
        <w:tblGrid><w:gridCol w:w="100"/><w:gridCol w:w="100"/></w:tblGrid>
        <w:tr>
          <w:tc><w:p><w:r><w:t xml:space="preserve">Week</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t xml:space="preserve">Owner</w:t></w:r></w:p></w:tc>
        </w:tr>
        <w:tr>
          <w:tc><w:p><w:r><w:t xml:space="preserve">W1</w:t></w:r></w:p></w:tc>
          <w:tc><w:p><w:r><w:t xml:space="preserve">Alex</w:t></w:r></w:p></w:tc>
        </w:tr>
      </w:tbl>
    `);
    const md = snapshotToMarkdown(agent.getSnapshot());
    expect(md).toContain("| Week | Owner |");
    expect(md).toContain("| --- | --- |");
    expect(md).toContain("| W1 | Alex |");
  });

  it("falls back to the placeholder + warns when a table has no rows", async () => {
    const warn = vi.spyOn(console, "warn").mockImplementation(() => {});
    try {
      const agent = await buildAgent(`
        <w:tbl><w:tblPr><w:tblW w:type="pct" w:w="100%"/></w:tblPr></w:tbl>
      `);
      const md = snapshotToMarkdown(agent.getSnapshot());
      expect(md).toContain("> [table preserved]");
      expect(warn).toHaveBeenCalled();
    } finally {
      warn.mockRestore();
    }
  });

  it("emits --- between sections (one per SectionBreak)", async () => {
    // The synthetic helper appends the trailing sectPr inside body; manually
    // wedge a mid-document sectPr so we get a real SectionBreak block in the
    // model, not just the body's terminal sectPr.
    const agent = await buildAgent(`
      <w:p><w:r><w:t xml:space="preserve">First section</w:t></w:r></w:p>
      <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
      <w:p><w:r><w:t xml:space="preserve">Second section</w:t></w:r></w:p>
    `);
    const md = snapshotToMarkdown(agent.getSnapshot());
    expect(md).toContain("First section");
    expect(md).toContain("Second section");
    expect(md.split("---").length).toBeGreaterThanOrEqual(2);
  });

  it("appends a ## Comments section that lists thread heads with the parent text snippet", async () => {
    const buf = await makeSyntheticDocx({
      documentXml: plainDocxXml([{ text: "Please review this clause carefully." }, { text: "Boilerplate." }]),
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
        author: "Reviewer",
        initials: "R",
      },
      source: "human",
    });
    const md = snapshotToMarkdown(agent.getSnapshot());
    expect(md).toContain("## Comments");
    expect(md).toContain("**Reviewer**");
    expect(md).toContain("rephrase?");
    expect(md).toContain('on "Please review this clause carefully."');
  });

  it("does NOT append a Comments section when the document has no comments", async () => {
    const buf = await makeSyntheticDocx({
      documentXml: plainDocxXml([{ text: "Just a body paragraph." }]),
    });
    const agent = await DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    const md = snapshotToMarkdown(agent.getSnapshot());
    expect(md).not.toContain("## Comments");
  });

  it("renders replies indented under the thread head", async () => {
    const buf = await makeSyntheticDocx({
      documentXml: plainDocxXml([{ text: "Original sentence here." }]),
    });
    const agent = await DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    await agent.applyCommand({
      type: "docx:add-comment",
      payload: {
        range: {
          start: { paragraph: 0, run: 0, offset: 0 },
          end: { paragraph: 0, run: 0, offset: 8 },
        },
        text: "Tighten this?",
        author: "Reviewer",
      },
      source: "human",
    });
    const headId = agent.getSnapshot().root.comments[0].id;
    await agent.applyCommand({
      type: "docx:reply-comment",
      payload: { parentId: headId, text: "Sure, will do.", author: "Author" },
      source: "human",
    });
    const md = snapshotToMarkdown(agent.getSnapshot());
    expect(md).toContain("**Reviewer**");
    expect(md).toContain("  - **Author**: Sure, will do.");
  });
});
