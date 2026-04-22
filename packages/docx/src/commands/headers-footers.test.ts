import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { deterministicIdMinter, ooxml, sha256Hex } from "@officeai/core";
import { DocxAgent } from "../agent/agent.js";
import { parseDocx } from "../parser/parse.js";
import { paragraphPlainText } from "./helpers.js";
import { makeSyntheticDocx, DEFAULT_DOC_ROOT_ATTRS, escapeXml } from "../test-utils/synthetic.js";

const FIXTURE_PATH = resolve(
  __dirname,
  "../../../../fixtures/docx/real-world/02-report-headers-footers.docx"
);

const HEADER_PART = "word/header1.xml";
const FOOTER_PART = "word/footer1.xml";

const HDR_NS_ATTRS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;

function syntheticDocXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body><w:p><w:r><w:t xml:space="preserve">Body para</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:footerReference w:type="default" r:id="rIdF"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
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

async function makeAgentWithHeaderFooter(
  headerText = "Acme Corp — Confidential",
  footerText = "Page footer · office-ai"
): Promise<DocxAgent> {
  const buf = await makeSyntheticDocx({
    documentXml: syntheticDocXml(),
    extra: {
      "word/_rels/document.xml.rels": syntheticDocRels(),
      [HEADER_PART]: syntheticHeaderXml(headerText),
      [FOOTER_PART]: syntheticFooterXml(footerText),
    },
  });
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

describe("docx headers/footers — parser + commands", () => {
  it("parses header and footer parts from a synthetic fixture", async () => {
    const agent = await makeAgentWithHeaderFooter("Hello header", "Hello footer");
    const snap = agent.getSnapshot();
    expect(snap.root.headersAndFooters).toHaveLength(2);
    const header = snap.root.headersAndFooters.find((p) => p.kind === "header");
    const footer = snap.root.headersAndFooters.find((p) => p.kind === "footer");
    expect(header).toBeTruthy();
    expect(footer).toBeTruthy();
    expect(header?.partPath).toBe(HEADER_PART);
    expect(footer?.partPath).toBe(FOOTER_PART);
    expect(header?.target).toBe("default");
    if (!header || header.body[0].kind !== "paragraph") throw new Error("header shape");
    expect(paragraphPlainText(header.body[0])).toBe("Hello header");
    if (!footer || footer.body[0].kind !== "paragraph") throw new Error("footer shape");
    expect(paragraphPlainText(footer.body[0])).toBe("Hello footer");
    expect(snap.dirty.headersAndFooters.size).toBe(0);
  });

  it("promotes a <w:tbl> inside a header part to a typed Table block", async () => {
    const headerWithTable = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr ${HDR_NS_ATTRS}><w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2880"/><w:gridCol w:w="2880"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Left</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2880" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Right</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:hdr>`;
    const buf = await makeSyntheticDocx({
      documentXml: syntheticDocXml(),
      extra: {
        "word/_rels/document.xml.rels": syntheticDocRels(),
        [HEADER_PART]: headerWithTable,
        [FOOTER_PART]: syntheticFooterXml("plain footer"),
      },
    });
    const snap = await parseDocx(buf);
    const header = snap.root.headersAndFooters.find((p) => p.kind === "header");
    expect(header).toBeTruthy();
    const block = header!.body[0];
    expect(block.kind).toBe("table");
    if (block.kind !== "table") throw new Error();
    expect(block.rows).toHaveLength(1);
    expect(block.rows[0]!.cells).toHaveLength(2);
  });

  it("parses a real-world fixture (02-report-headers-footers.docx) and reads header/footer text", async () => {
    const buf = await readFile(FIXTURE_PATH);
    const snap = await parseDocx(buf);
    expect(snap.root.headersAndFooters.length).toBeGreaterThanOrEqual(2);
    const header = snap.root.headersAndFooters.find((p) => p.kind === "header");
    const footer = snap.root.headersAndFooters.find((p) => p.kind === "footer");
    expect(header).toBeTruthy();
    expect(footer).toBeTruthy();
    if (!header || header.body[0].kind !== "paragraph") throw new Error("header shape");
    if (!footer || footer.body[0].kind !== "paragraph") throw new Error("footer shape");
    expect(paragraphPlainText(header.body[0])).toBe("Acme Corp — Confidential");
    expect(paragraphPlainText(footer.body[0])).toBe("Page footer · office-ai fixture");
  });

  it("set-header-text replaces the targeted paragraph and dirties only that part", async () => {
    const agent = await makeAgentWithHeaderFooter();
    const m = await agent.applyCommand({
      type: "docx:set-header-text",
      payload: { partId: HEADER_PART, paragraphIndex: 0, text: "New header text" },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const snap = agent.getSnapshot();
    const header = snap.root.headersAndFooters.find((p) => p.kind === "header");
    if (!header || header.body[0].kind !== "paragraph") throw new Error();
    expect(paragraphPlainText(header.body[0])).toBe("New header text");
    expect(snap.dirty.headersAndFooters.has(HEADER_PART)).toBe(true);
    expect(snap.dirty.headersAndFooters.has(FOOTER_PART)).toBe(false);
    expect(snap.dirty.body).toBe(false);
  });

  it("set-footer-text + DocxAgent round-trip preserves the new text", async () => {
    const agent = await makeAgentWithHeaderFooter();
    await agent.applyCommand({
      type: "docx:set-footer-text",
      payload: { partId: FOOTER_PART, paragraphIndex: 0, text: "Footer v2" },
      source: "human",
    });
    const out = await agent.exportFile();
    const reparsed = await parseDocx(out);
    const footer = reparsed.root.headersAndFooters.find((p) => p.kind === "footer");
    if (!footer || footer.body[0].kind !== "paragraph") throw new Error();
    expect(paragraphPlainText(footer.body[0])).toBe("Footer v2");
  });

  it("set-header-text is idempotent — re-applying with the same text bumps the revision but is otherwise equivalent", async () => {
    const agent = await makeAgentWithHeaderFooter();
    await agent.applyCommand({
      type: "docx:set-header-text",
      payload: { partId: HEADER_PART, paragraphIndex: 0, text: "Stable" },
      source: "human",
    });
    const r1 = agent.getSnapshot().revision;
    await agent.applyCommand({
      type: "docx:set-header-text",
      payload: { partId: HEADER_PART, paragraphIndex: 0, text: "Stable" },
      source: "human",
    });
    const snap = agent.getSnapshot();
    expect(snap.revision).toBe(r1 + 1);
    const header = snap.root.headersAndFooters.find((p) => p.kind === "header");
    if (!header || header.body[0].kind !== "paragraph") throw new Error();
    expect(paragraphPlainText(header.body[0])).toBe("Stable");
  });

  it("set-header-text rejects a missing part with code unknown-target", async () => {
    const agent = await makeAgentWithHeaderFooter();
    const m = await agent.applyCommand({
      type: "docx:set-header-text",
      payload: { partId: "word/header9.xml", paragraphIndex: 0, text: "x" },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-target");
  });

  it("set-header-text rejects an out-of-range paragraph index with code unknown-target", async () => {
    const agent = await makeAgentWithHeaderFooter();
    const m = await agent.applyCommand({
      type: "docx:set-header-text",
      payload: { partId: HEADER_PART, paragraphIndex: 99, text: "x" },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-target");
  });

  it("byte-preservation invariant: untouched header/footer parts re-emit identical SHA-256 on a round-trip with no mutations", async () => {
    const buf = await readFile(FIXTURE_PATH);
    const original = await ooxml.OoxmlContainer.load(buf);
    const originalHeaderHash = sha256Hex(original.readBytes(HEADER_PART));
    const originalFooterHash = sha256Hex(original.readBytes(FOOTER_PART));

    const agent = await DocxAgent.fromBuffer(buf);
    const out = await agent.exportFile();
    const reloaded = await ooxml.OoxmlContainer.load(out);
    expect(sha256Hex(reloaded.readBytes(HEADER_PART))).toBe(originalHeaderHash);
    expect(sha256Hex(reloaded.readBytes(FOOTER_PART))).toBe(originalFooterHash);
  });

  it("byte-preservation invariant: footer stays byte-identical when only the header is mutated", async () => {
    const buf = await readFile(FIXTURE_PATH);
    const original = await ooxml.OoxmlContainer.load(buf);
    const originalFooterHash = sha256Hex(original.readBytes(FOOTER_PART));

    const agent = await DocxAgent.fromBuffer(buf);
    await agent.applyCommand({
      type: "docx:set-header-text",
      payload: { partId: HEADER_PART, paragraphIndex: 0, text: "MUTATED" },
      source: "human",
    });
    const out = await agent.exportFile();
    const reloaded = await ooxml.OoxmlContainer.load(out);
    expect(sha256Hex(reloaded.readBytes(FOOTER_PART))).toBe(originalFooterHash);
    // Header bytes did change.
    expect(sha256Hex(reloaded.readBytes(HEADER_PART))).not.toBe(sha256Hex(original.readBytes(HEADER_PART)));
    // And the new content survives a full re-parse.
    const reparsed = await parseDocx(out);
    const header = reparsed.root.headersAndFooters.find((p) => p.kind === "header");
    if (!header || header.body[0].kind !== "paragraph") throw new Error();
    expect(paragraphPlainText(header.body[0])).toBe("MUTATED");
  });
});
