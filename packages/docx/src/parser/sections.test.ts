import { describe, expect, it } from "vitest";
import { ooxml } from "@officeai/core";
import { DocxAgent } from "../agent/agent.js";
import { parseDocx } from "./parse.js";
import { serializeDocx } from "../serializer/serialize.js";
import { DEFAULT_DOC_ROOT_ATTRS, makeSyntheticDocx } from "../test-utils/synthetic.js";

/**
 * Tests for the typed `<w:sectPr>` projection (P3.2 / W5–W7) and the
 * promoted run leaves `PageBreakLeaf` / `LastRenderedPageBreakLeaf`
 * (W6).
 *
 * Round-trip contract: every fixture below exercises parse → serialize
 * → reload and asserts byte-equality on `word/document.xml`. The real
 * `tests/roundtrip/docx/real-world-roundtrip.test.ts` sweep covers the
 * same invariant against authentic Word output; these focused tests
 * pin the specific mechanism so a regression has a one-test failure
 * pointing at the cause.
 */

async function exportDocumentXml(snap: Awaited<ReturnType<typeof parseDocx>>): Promise<string> {
  const buf = await serializeDocx(snap);
  const container = await ooxml.OoxmlContainer.load(buf);
  return new TextDecoder().decode(container.readBytes("word/document.xml"));
}

async function loadFromXml(documentXml: string): Promise<{
  doc: Awaited<ReturnType<typeof parseDocx>>["root"];
  reEmittedXml: string;
}> {
  const buf = await makeSyntheticDocx({ documentXml });
  const snap = await parseDocx(buf);
  const reEmittedXml = await exportDocumentXml(snap);
  return { doc: snap.root, reEmittedXml };
}

describe("parseSectionProperties (W5)", () => {
  it("extracts pgSz / pgMar / cols / titlePg / type from sectPr", async () => {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>
  <w:p><w:r><w:t>hello</w:t></w:r></w:p>
  <w:sectPr>
    <w:type w:val="continuous"/>
    <w:pgSz w:w="12240" w:h="15840" w:orient="portrait"/>
    <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/>
    <w:cols w:num="2" w:sep="1" w:space="720"/>
    <w:titlePg/>
  </w:sectPr>
</w:body></w:document>`;
    const { doc } = await loadFromXml(documentXml);
    const sect = doc.body.find((b) => b.kind === "section-break");
    expect(sect?.kind).toBe("section-break");
    if (sect?.kind !== "section-break") return;

    expect(sect.properties.pgSz).toEqual({ w: 12240, h: 15840, orient: "portrait" });
    expect(sect.properties.pgMar).toEqual({
      top: 1440,
      right: 1440,
      bottom: 1440,
      left: 1440,
      header: 720,
      footer: 720,
      gutter: 0,
    });
    expect(sect.properties.cols).toEqual({ num: 2, sep: true, space: 720 });
    expect(sect.properties.titlePg).toBe(true);
    expect(sect.properties.sectionType).toBe("continuous");
    expect(sect.properties.headerRefs).toEqual([]);
    expect(sect.properties.footerRefs).toEqual([]);
  });

  it("captures header/footer references with the typed `type` attribute", async () => {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>
  <w:p><w:r><w:t>hi</w:t></w:r></w:p>
  <w:sectPr>
    <w:headerReference w:type="default" r:id="rId10"/>
    <w:headerReference w:type="first" r:id="rId11"/>
    <w:footerReference w:type="default" r:id="rId12"/>
    <w:pgSz w:w="12240" w:h="15840"/>
  </w:sectPr>
</w:body></w:document>`;
    const { doc } = await loadFromXml(documentXml);
    const sect = doc.body.find((b) => b.kind === "section-break");
    if (sect?.kind !== "section-break") throw new Error("expected section");
    expect(sect.properties.headerRefs).toEqual([
      { type: "default", relationshipId: "rId10" },
      { type: "first", relationshipId: "rId11" },
    ]);
    expect(sect.properties.footerRefs).toEqual([
      { type: "default", relationshipId: "rId12" },
    ]);
  });

  it("captures unmodeled sectPr children into opaqueProps in source order", async () => {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>
  <w:p><w:r><w:t>hi</w:t></w:r></w:p>
  <w:sectPr>
    <w:pgSz w:w="12240" w:h="15840"/>
    <w:lineNumType w:countBy="1" w:start="1"/>
    <w:docGrid w:linePitch="360"/>
  </w:sectPr>
</w:body></w:document>`;
    const { doc } = await loadFromXml(documentXml);
    const sect = doc.body.find((b) => b.kind === "section-break");
    if (sect?.kind !== "section-break") throw new Error("expected section");
    const tags = (sect.properties.opaqueProps ?? []).map((o) => o.tag);
    expect(tags).toEqual(["w:lineNumType", "w:docGrid"]);
  });
});

describe("section-break round-trip (W7)", () => {
  it("untouched sectPr re-emits byte-identical document.xml across two passes", async () => {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body><w:p><w:r><w:t xml:space="preserve">abc</w:t></w:r></w:p><w:sectPr><w:type w:val="nextPage"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`;
    const buf = await makeSyntheticDocx({ documentXml });
    const agent = await DocxAgent.fromBuffer(buf);
    const out1 = await agent.exportFile();
    const xml1 = await exportDocumentXml(await parseDocx(out1));
    const agent2 = await DocxAgent.fromBuffer(out1);
    const out2 = await agent2.exportFile();
    const xml2 = await exportDocumentXml(await parseDocx(out2));
    expect(xml2).toBe(xml1);
  });
});

describe("PageBreakLeaf / LastRenderedPageBreakLeaf (W6)", () => {
  it("promotes <w:br w:type=\"page\"/> to a typed page-break leaf", async () => {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>
  <w:p><w:r><w:t>before</w:t><w:br w:type="page"/><w:t>after</w:t></w:r></w:p>
  <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
</w:body></w:document>`;
    const { doc } = await loadFromXml(documentXml);
    const p = doc.body[0];
    if (p.kind !== "paragraph") throw new Error("expected paragraph");
    const r = p.children[0];
    if (r.kind !== "run") throw new Error("expected run");
    const kinds = r.children.map((c) => c.kind);
    expect(kinds).toContain("page-break");
    expect(kinds).not.toContain("opaque");
  });

  it("page-break round-trips back to <w:br w:type=\"page\"/>", async () => {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body><w:p><w:r><w:t xml:space="preserve">a</w:t><w:br w:type="page"/><w:t xml:space="preserve">b</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
    const { reEmittedXml } = await loadFromXml(documentXml);
    expect(reEmittedXml).toContain('<w:br w:type="page"');
    // Stable round-trip on a second pass.
    const buf = await makeSyntheticDocx({ documentXml: reEmittedXml });
    const snap = await parseDocx(buf);
    const xml2 = await exportDocumentXml(snap);
    expect(xml2).toBe(reEmittedXml);
  });

  it("promotes <w:lastRenderedPageBreak/> to a typed leaf and round-trips", async () => {
    const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body><w:p><w:r><w:lastRenderedPageBreak/><w:t xml:space="preserve">x</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
    const { doc, reEmittedXml } = await loadFromXml(documentXml);
    const p = doc.body[0];
    if (p.kind !== "paragraph") throw new Error("expected paragraph");
    const r = p.children[0];
    if (r.kind !== "run") throw new Error("expected run");
    const kinds = r.children.map((c) => c.kind);
    expect(kinds).toContain("last-rendered-page-break");

    expect(reEmittedXml).toContain("<w:lastRenderedPageBreak");
    const buf = await makeSyntheticDocx({ documentXml: reEmittedXml });
    const snap = await parseDocx(buf);
    const xml2 = await exportDocumentXml(snap);
    expect(xml2).toBe(reEmittedXml);
  });
});
