import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { DocxAgent } from "../agent/agent.js";
import { paragraphPlainText } from "./helpers.js";
import { makeSyntheticDocx, DEFAULT_DOC_ROOT_ATTRS } from "../test-utils/synthetic.js";
import type { BlockNode, InlineNode, Paragraph } from "../model/types.js";

function trackedDocXml(): string {
  const ins1 = `<w:ins w:id="100" w:author="Alice" w:date="2026-04-17T10:00:00Z"><w:r><w:t xml:space="preserve">INS-A </w:t></w:r></w:ins>`;
  const del1 = `<w:del w:id="200" w:author="Alice" w:date="2026-04-17T10:01:00Z"><w:r><w:delText xml:space="preserve">DEL-A </w:delText></w:r></w:del>`;
  const ins2 = `<w:ins w:id="300" w:author="Bob" w:date="2026-04-17T10:02:00Z"><w:r><w:t xml:space="preserve">INS-B </w:t></w:r></w:ins>`;
  const tail = `<w:r><w:t xml:space="preserve">end.</w:t></w:r>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body><w:p>${ins1}${del1}${ins2}${tail}</w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
}

async function loadAgent(): Promise<DocxAgent> {
  const buf = await makeSyntheticDocx({ documentXml: trackedDocXml() });
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

function paraOf(block: BlockNode): Paragraph {
  if (block.kind !== "paragraph") throw new Error("expected paragraph");
  return block;
}

function revisionCount(snapshot: { root: { body: ReadonlyArray<BlockNode> } }): number {
  let n = 0;
  for (const block of snapshot.root.body) {
    if (block.kind !== "paragraph") continue;
    for (const child of block.children) n += countInline(child);
  }
  return n;
}

function countInline(node: InlineNode): number {
  if (node.kind !== "revision") return 0;
  let n = 1;
  for (const c of node.children) n += countInline(c);
  return n;
}

describe("docx:accept-all-changes / docx:reject-all-changes (B8)", () => {
  it("accept-all keeps insertions, drops deletions, leaves no revisions behind", async () => {
    const agent = await loadAgent();
    const m = await agent.applyCommand({
      type: "docx:accept-all-changes",
      payload: {},
      source: "human",
    });
    expect(m.status).toBe("approved");
    const snap = agent.getSnapshot();
    expect(revisionCount(snap)).toBe(0);
    expect(paragraphPlainText(paraOf(snap.root.body[0]))).toBe("INS-A INS-B end.");
  });

  it("reject-all drops insertions, restores deletions, leaves no revisions behind", async () => {
    const agent = await loadAgent();
    const m = await agent.applyCommand({
      type: "docx:reject-all-changes",
      payload: {},
      source: "human",
    });
    expect(m.status).toBe("approved");
    const snap = agent.getSnapshot();
    expect(revisionCount(snap)).toBe(0);
    expect(paragraphPlainText(paraOf(snap.root.body[0]))).toBe("DEL-A end.");
  });

  it("is a no-op when there are no tracked changes", async () => {
    const buf = await makeSyntheticDocx({
      documentXml: `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body><w:p><w:r><w:t>plain</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
    });
    const agent = await DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    const before = agent.getSnapshot().revision;
    const m = await agent.applyCommand({
      type: "docx:accept-all-changes",
      payload: {},
      source: "human",
    });
    expect(m.status).toBe("approved");
    expect(agent.getSnapshot().revision).toBe(before);
  });
});
