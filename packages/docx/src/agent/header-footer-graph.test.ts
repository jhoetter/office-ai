import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { DocxAgent } from "./agent.js";
import { resolveHeaderFooterParts } from "./header-footer-graph.js";
import { DEFAULT_DOC_ROOT_ATTRS, escapeXml, makeSyntheticDocx } from "../test-utils/synthetic.js";

const HDR_NS_ATTRS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function headerXml(text: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr ${HDR_NS_ATTRS}><w:p><w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p></w:hdr>`;
}

function footerXml(text: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr ${HDR_NS_ATTRS}><w:p><w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p></w:ftr>`;
}

describe("resolveHeaderFooterParts (P3.2 / W8)", () => {
  it("returns the typed parts referenced by the section's headerRefs/footerRefs", async () => {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>
  <w:p><w:r><w:t>body</w:t></w:r></w:p>
  <w:sectPr>
    <w:headerReference w:type="default" r:id="rIdH"/>
    <w:headerReference w:type="first" r:id="rIdHFirst"/>
    <w:footerReference w:type="default" r:id="rIdF"/>
    <w:titlePg/>
    <w:pgSz w:w="12240" w:h="15840"/>
  </w:sectPr>
</w:body></w:document>`;

    const docRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
<Relationship Id="rIdHFirst" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header2.xml"/>
<Relationship Id="rIdF" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>
</Relationships>`;

    const buf = await makeSyntheticDocx({
      documentXml,
      extra: {
        "word/_rels/document.xml.rels": docRels,
        "word/header1.xml": headerXml("Default header"),
        "word/header2.xml": headerXml("First-page header"),
        "word/footer1.xml": footerXml("Default footer"),
      },
    });
    const agent = await DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    const snap = agent.getSnapshot();
    const sectionIndex = snap.root.body.findIndex((b) => b.kind === "section-break");
    expect(sectionIndex).toBeGreaterThanOrEqual(0);

    const resolved = resolveHeaderFooterParts(snap, sectionIndex);
    expect(resolved.headers.default?.partPath).toBe("word/header1.xml");
    expect(resolved.headers.first?.partPath).toBe("word/header2.xml");
    expect(resolved.headers.even).toBeUndefined();
    expect(resolved.footers.default?.partPath).toBe("word/footer1.xml");
    expect(resolved.footers.first).toBeUndefined();
  });

  it("returns empty slots for an out-of-range index or non-section block", async () => {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>
  <w:p><w:r><w:t>only paragraph</w:t></w:r></w:p>
  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
</w:body></w:document>`;
    const buf = await makeSyntheticDocx({ documentXml });
    const agent = await DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    const snap = agent.getSnapshot();

    const fromOutOfRange = resolveHeaderFooterParts(snap, 999);
    expect(fromOutOfRange).toEqual({ headers: {}, footers: {} });

    const paragraphIndex = snap.root.body.findIndex((b) => b.kind === "paragraph");
    const fromParagraph = resolveHeaderFooterParts(snap, paragraphIndex);
    expect(fromParagraph).toEqual({ headers: {}, footers: {} });
  });

  it("ignores references to relationships that do not resolve to a part", async () => {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>
  <w:p><w:r><w:t>body</w:t></w:r></w:p>
  <w:sectPr>
    <w:headerReference w:type="default" r:id="rIdMissing"/>
    <w:pgSz w:w="12240" w:h="15840"/>
  </w:sectPr>
</w:body></w:document>`;
    const buf = await makeSyntheticDocx({ documentXml });
    const agent = await DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    const snap = agent.getSnapshot();
    const sectionIndex = snap.root.body.findIndex((b) => b.kind === "section-break");
    const resolved = resolveHeaderFooterParts(snap, sectionIndex);
    expect(resolved.headers).toEqual({});
    expect(resolved.footers).toEqual({});
  });
});
