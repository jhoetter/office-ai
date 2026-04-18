import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { DocxAgent } from "../agent/agent.js";
import { parseDocx } from "../parser/parse.js";
import {
  DEFAULT_DOC_ROOT_ATTRS,
  escapeXml,
  makeSyntheticDocx,
} from "../test-utils/synthetic.js";
import type { Paragraph, Run } from "../model/types.js";
import { chunkIntoPages } from "../renderer/page-chunker.js";

function syntheticDocXml(paraCount = 2): string {
  const paras = Array.from({ length: paraCount }, (_, i) =>
    `<w:p><w:r><w:t xml:space="preserve">${escapeXml(`Paragraph ${i + 1} body text`)}</w:t></w:r></w:p>`
  ).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>${paras}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`;
}

async function makeAgent(paraCount = 2): Promise<DocxAgent> {
  const buf = await makeSyntheticDocx({ documentXml: syntheticDocXml(paraCount) });
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

describe("docx:insert-page-break (P3.5 / W18)", () => {
  it("inserts a typed page-break leaf inside the targeted paragraph", async () => {
    const agent = await makeAgent();
    const para = agent.getSnapshot().root.body[0] as Paragraph;
    const m = await agent.applyCommand({
      type: "docx:insert-page-break",
      payload: { paragraphId: para.id, offset: 0 },
      source: "human",
    });
    expect(m.status).toBe("approved");

    const snap = agent.getSnapshot();
    const updated = snap.root.body[0] as Paragraph;
    const breakRun = updated.children.find(
      (c): c is Run => c.kind === "run" && c.children.some((rc) => rc.kind === "page-break")
    );
    expect(breakRun).toBeTruthy();
    expect(snap.dirty.body).toBe(true);
  });

  it("splits a run when inserting in the middle", async () => {
    const agent = await makeAgent();
    const para = agent.getSnapshot().root.body[0] as Paragraph;
    // Original text is "Paragraph 1 body text". Insert page break after
    // "Paragraph " (10 chars).
    await agent.applyCommand({
      type: "docx:insert-page-break",
      payload: { paragraphId: para.id, offset: 10 },
      source: "human",
    });
    const snap = agent.getSnapshot();
    const updated = snap.root.body[0] as Paragraph;
    const breakIdx = updated.children.findIndex(
      (c) => c.kind === "run" && c.children.some((rc) => rc.kind === "page-break")
    );
    expect(breakIdx).toBeGreaterThan(0);
    expect(breakIdx).toBeLessThan(updated.children.length - 1);
  });

  it("advances the page-chunker output by one page", async () => {
    const agent = await makeAgent(3);
    expect(chunkIntoPages(agent.getSnapshot()).length).toBe(1);
    const para = agent.getSnapshot().root.body[1] as Paragraph;
    await agent.applyCommand({
      type: "docx:insert-page-break",
      payload: { paragraphId: para.id, offset: 0 },
      source: "human",
    });
    const chunks = chunkIntoPages(agent.getSnapshot());
    expect(chunks.length).toBe(2);
  });

  it("rejects unknown paragraph id with unknown-target", async () => {
    const agent = await makeAgent();
    const m = await agent.applyCommand({
      type: "docx:insert-page-break",
      payload: { paragraphId: "ghost-para", offset: 0 },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-target");
  });

  it("rejects negative offset with invalid-payload", async () => {
    const agent = await makeAgent();
    const para = agent.getSnapshot().root.body[0] as Paragraph;
    const m = await agent.applyCommand({
      type: "docx:insert-page-break",
      payload: { paragraphId: para.id, offset: -5 },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-payload");
  });

  it("round-trip: page break survives parse → export → parse as <w:br w:type=\"page\"/>", async () => {
    const agent = await makeAgent();
    const para = agent.getSnapshot().root.body[0] as Paragraph;
    await agent.applyCommand({
      type: "docx:insert-page-break",
      payload: { paragraphId: para.id, offset: 5 },
      source: "human",
    });
    const buf = await agent.exportFile();
    const reparsed = await parseDocx(buf);
    const updated = reparsed.root.body[0] as Paragraph;
    const breakRun = updated.children.find(
      (c): c is Run => c.kind === "run" && c.children.some((rc) => rc.kind === "page-break")
    );
    expect(breakRun).toBeTruthy();
  });
});
