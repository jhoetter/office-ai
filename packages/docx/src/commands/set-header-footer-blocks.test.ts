import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { DocxAgent } from "../agent/agent.js";
import { parseDocx } from "../parser/parse.js";
import { DEFAULT_DOC_ROOT_ATTRS, escapeXml, makeSyntheticDocx } from "../test-utils/synthetic.js";
import type { BlockNode, PageNumberFieldLeaf, Paragraph } from "../model/types.js";

const HDR_NS_ATTRS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;

const HEADER_PART = "word/header1.xml";
const FOOTER_PART = "word/footer1.xml";

function syntheticDocXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body><w:p><w:r><w:t xml:space="preserve">Body</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`;
}

function syntheticHeaderXml(text: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr ${HDR_NS_ATTRS}><w:p><w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p></w:hdr>`;
}

function syntheticFooterXml(text: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr ${HDR_NS_ATTRS}><w:p><w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p></w:ftr>`;
}

function syntheticDocRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
<Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>`;
}

async function makeAgent(): Promise<DocxAgent> {
  const buf = await makeSyntheticDocx({
    documentXml: syntheticDocXml(),
    extra: {
      "word/_rels/document.xml.rels": syntheticDocRels(),
      [HEADER_PART]: syntheticHeaderXml("Header"),
      [FOOTER_PART]: syntheticFooterXml("Footer"),
    },
  });
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

function makeRichBody(mintId: () => string): BlockNode[] {
  // Two paragraphs: one with a literal text run plus a PAGE field,
  // one with just text. This is the canonical "Word-style header"
  // shape we need to round-trip without flattening.
  const field: PageNumberFieldLeaf = {
    kind: "page-number-field",
    id: mintId(),
    field: "PAGE",
    instr: " PAGE \\* MERGEFORMAT ",
  };
  const p1: Paragraph = {
    kind: "paragraph",
    id: mintId(),
    properties: {},
    children: [
      {
        kind: "run",
        id: mintId(),
        properties: {},
        children: [{ kind: "text", id: mintId(), text: "Page ", xmlSpacePreserve: true }],
      },
      {
        kind: "run",
        id: mintId(),
        properties: {},
        children: [field],
      },
    ],
  };
  const p2: Paragraph = {
    kind: "paragraph",
    id: mintId(),
    properties: {},
    children: [
      {
        kind: "run",
        id: mintId(),
        properties: {},
        children: [{ kind: "text", id: mintId(), text: "Confidential", xmlSpacePreserve: false }],
      },
    ],
  };
  return [p1, p2];
}

describe("docx:set-header-footer-blocks", () => {
  it("replaces the part body with a multi-paragraph + field body", async () => {
    const agent = await makeAgent();
    const minter = deterministicIdMinter("blocks");
    const body = makeRichBody(minter);

    const m = await agent.applyCommand({
      type: "docx:set-header-footer-blocks",
      payload: { partPath: HEADER_PART, body },
      source: "human",
    });
    expect(m.status).toBe("approved");

    const snap = agent.getSnapshot();
    const header = snap.root.headersAndFooters.find((p) => p.partPath === HEADER_PART);
    expect(header?.body.length).toBe(2);
    const firstPara = header!.body[0] as Paragraph;
    const fieldRun = firstPara.children.find(
      (c) => c.kind === "run" && c.children.some((rc) => rc.kind === "page-number-field")
    );
    expect(fieldRun).toBeTruthy();
    expect(snap.dirty.headersAndFooters.has(HEADER_PART)).toBe(true);
  });

  it("rejects an empty body with invalid-payload", async () => {
    const agent = await makeAgent();
    const m = await agent.applyCommand({
      type: "docx:set-header-footer-blocks",
      payload: { partPath: HEADER_PART, body: [] },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-payload");
  });

  it("rejects an unknown partPath with unknown-target", async () => {
    const agent = await makeAgent();
    const minter = deterministicIdMinter("blocks");
    const m = await agent.applyCommand({
      type: "docx:set-header-footer-blocks",
      payload: { partPath: "word/header99.xml", body: makeRichBody(minter) },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-target");
  });

  it("survives parse → set-blocks → export → parse", async () => {
    const agent = await makeAgent();
    const minter = deterministicIdMinter("blocks");
    await agent.applyCommand({
      type: "docx:set-header-footer-blocks",
      payload: { partPath: HEADER_PART, body: makeRichBody(minter) },
      source: "human",
    });
    const buf = await agent.exportFile();
    const reparsed = await parseDocx(buf);
    const header = reparsed.root.headersAndFooters.find((p) => p.partPath === HEADER_PART);
    expect(header).toBeTruthy();
    // Two paragraphs survived. (The PAGE field round-trips as either
    // a typed leaf or an opaque-inline; what we care about here is
    // that the multi-paragraph body shape was preserved end-to-end.)
    expect(header!.body.length).toBe(2);
    const second = header!.body[1] as Paragraph;
    expect(second.kind).toBe("paragraph");
    const flat = second.children
      .flatMap((c) => (c.kind === "run" ? c.children : []))
      .filter((c) => c.kind === "text")
      .map((c) => (c as { text: string }).text)
      .join("");
    expect(flat).toBe("Confidential");
  });
});
