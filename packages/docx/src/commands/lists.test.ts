import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { deterministicIdMinter, ooxml, sha256Hex } from "@officeai/core";
import { DocxAgent } from "../agent/agent.js";
import { parseDocx } from "../parser/parse.js";
import { serializeDocx } from "../serializer/serialize.js";
import type { BlockNode, Paragraph, Table } from "../model/types.js";
import { DEFAULT_DOC_ROOT_ATTRS, escapeXml, makeSyntheticDocx } from "../test-utils/synthetic.js";

const FIXTURE_PATH = resolve(__dirname, "../../../../fixtures/docx/real-world/03-numbered-list.docx");

const NUMBERING_XML_TWO_ABSTRACTS_THREE_NUMS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:abstractNum w:abstractNumId="1">
    <w:multiLevelType w:val="hybridMultilevel"/>
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/></w:lvl>
    <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="○"/></w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="2">
    <w:multiLevelType w:val="multilevel"/>
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num>
  <w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num>
</w:numbering>`;

const NUMBERING_CONTENT_TYPE_OVERRIDE = `<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>`;

const NUMBERING_REL = `<Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>`;

function paragraphsXml(paragraphs: ReadonlyArray<{ text: string; numId?: number; ilvl?: number }>): string {
  return paragraphs
    .map((p) => {
      const numPr =
        p.numId !== undefined && p.ilvl !== undefined
          ? `<w:pPr><w:numPr><w:ilvl w:val="${p.ilvl}"/><w:numId w:val="${p.numId}"/></w:numPr></w:pPr>`
          : "";
      return `<w:p>${numPr}<w:r><w:t xml:space="preserve">${escapeXml(p.text)}</w:t></w:r></w:p>`;
    })
    .join("");
}

function bodyXml(inner: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>${inner}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
}

const CONTENT_TYPES_WITH_NUMBERING = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  ${NUMBERING_CONTENT_TYPE_OVERRIDE}
</Types>`;

const DOC_RELS_WITH_NUMBERING = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  ${NUMBERING_REL}
</Relationships>`;

async function loadAgentWithNumbering(
  paragraphs: ReadonlyArray<{ text: string; numId?: number; ilvl?: number }>
): Promise<DocxAgent> {
  const buf = await makeSyntheticDocx({
    documentXml: bodyXml(paragraphsXml(paragraphs)),
    extra: {
      "[Content_Types].xml": CONTENT_TYPES_WITH_NUMBERING,
      "word/_rels/document.xml.rels": DOC_RELS_WITH_NUMBERING,
      "word/numbering.xml": NUMBERING_XML_TWO_ABSTRACTS_THREE_NUMS,
    },
  });
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

function paraOf(block: BlockNode): Paragraph {
  if (block.kind !== "paragraph") throw new Error("expected paragraph");
  return block;
}

describe("docx lists — parser + commands (P1.4 / W10)", () => {
  it("parses a numbering.xml with two abstractNums + three nums into the typed carrier", async () => {
    const agent = await loadAgentWithNumbering([{ text: "x" }]);
    const snap = agent.getSnapshot();
    expect(snap.root.numbering).toBeTruthy();
    if (!snap.root.numbering) return;
    expect(snap.root.numbering.abstractNums.size).toBe(2);
    expect(snap.root.numbering.nums.size).toBe(3);
    const a1 = snap.root.numbering.abstractNums.get("1");
    expect(a1?.multiLevelType).toBe("hybridMultilevel");
    expect(a1?.levels.length).toBe(2);
    const n2 = snap.root.numbering.nums.get(2);
    expect(n2?.abstractNumId).toBe("2");
  });

  it("returns root.numbering === undefined when the doc has no numbering.xml", async () => {
    const buf = await makeSyntheticDocx({ documentXml: bodyXml(paragraphsXml([{ text: "no list" }])) });
    const snap = await parseDocx(buf, { idMinter: deterministicIdMinter() });
    expect(snap.root.numbering).toBeUndefined();
  });

  it("round-trips a fixture containing numbering.xml byte-identically when no list command runs", async () => {
    const buf = await readFile(FIXTURE_PATH);
    const snap = await parseDocx(buf, { idMinter: deterministicIdMinter() });
    const out = await serializeDocx(snap);
    const reloaded = await ooxml.OoxmlContainer.load(out);
    for (const path of ["word/numbering.xml", "word/document.xml", "word/_rels/document.xml.rels"]) {
      const before = sha256Hex(snap.container.readBytes(path));
      const after = sha256Hex(reloaded.readBytes(path));
      expect(after, `part ${path} should be byte-identical`).toBe(before);
    }
  });

  it("set-paragraph-list happy path on a body paragraph; dirty.body set, dirty.numbering not set", async () => {
    const agent = await loadAgentWithNumbering([{ text: "Item one" }]);
    const id = paraOf(agent.getSnapshot().root.body[0]).id;
    const m = await agent.applyCommand({
      type: "docx:set-paragraph-list",
      payload: { paragraphId: id, numId: 1, ilvl: 0 },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const snap = agent.getSnapshot();
    expect(snap.dirty.body).toBe(true);
    expect(snap.dirty.numbering).toBe(false);
    const p = paraOf(snap.root.body[0]);
    expect(p.properties.numbering).toEqual({ numId: 1, ilvl: 0 });
  });

  it("set-paragraph-list works on a paragraph nested inside a table cell", async () => {
    const tableXml = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t xml:space="preserve">cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:p/>`;
    const buf = await makeSyntheticDocx({
      documentXml: bodyXml(tableXml),
      extra: {
        "[Content_Types].xml": CONTENT_TYPES_WITH_NUMBERING,
        "word/_rels/document.xml.rels": DOC_RELS_WITH_NUMBERING,
        "word/numbering.xml": NUMBERING_XML_TWO_ABSTRACTS_THREE_NUMS,
      },
    });
    const agent = await DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    const tbl = agent.getSnapshot().root.body[0] as Table;
    expect(tbl.kind).toBe("table");
    const cellPara = tbl.rows[0].cells[0].body[0] as Paragraph;
    const m = await agent.applyCommand({
      type: "docx:set-paragraph-list",
      payload: { paragraphId: cellPara.id, numId: 2, ilvl: 0 },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const snap2 = agent.getSnapshot();
    const tbl2 = snap2.root.body[0] as Table;
    const cellPara2 = tbl2.rows[0].cells[0].body[0] as Paragraph;
    expect(cellPara2.properties.numbering).toEqual({ numId: 2, ilvl: 0 });
    // Round-trips through parse(serialize(s)).
    const out = await serializeDocx(snap2);
    const snap3 = await parseDocx(out, { idMinter: deterministicIdMinter("z") });
    const tbl3 = snap3.root.body[0] as Table;
    const cellPara3 = tbl3.rows[0].cells[0].body[0] as Paragraph;
    expect(cellPara3.properties.numbering).toEqual({ numId: 2, ilvl: 0 });
  });

  it("set-paragraph-list replaces an existing list reference on a paragraph", async () => {
    const agent = await loadAgentWithNumbering([{ text: "already a list", numId: 1, ilvl: 0 }]);
    const id = paraOf(agent.getSnapshot().root.body[0]).id;
    await agent.applyCommand({
      type: "docx:set-paragraph-list",
      payload: { paragraphId: id, numId: 2, ilvl: 0 },
      source: "human",
    });
    const p = paraOf(agent.getSnapshot().root.body[0]);
    expect(p.properties.numbering).toEqual({ numId: 2, ilvl: 0 });
    // The opaque numPr carrier is gone — only one numPr ends up in the
    // serialized output regardless of typed/opaque interleave.
    const out = await serializeDocx(agent.getSnapshot());
    const reloaded = await ooxml.OoxmlContainer.load(out);
    const xml = reloaded.readText("word/document.xml");
    const numPrCount = (xml.match(/<w:numPr/g) ?? []).length;
    expect(numPrCount).toBe(1);
    expect(xml).toContain('<w:numId w:val="2"');
  });

  it("set-paragraph-list rejects unknown numId with unknown-target", async () => {
    const agent = await loadAgentWithNumbering([{ text: "x" }]);
    const id = paraOf(agent.getSnapshot().root.body[0]).id;
    const m = await agent.applyCommand({
      type: "docx:set-paragraph-list",
      payload: { paragraphId: id, numId: 99, ilvl: 0 },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-target");
  });

  it("set-paragraph-list rejects when document has no numbering.xml", async () => {
    const buf = await makeSyntheticDocx({ documentXml: bodyXml(paragraphsXml([{ text: "no list" }])) });
    const agent = await DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    const id = paraOf(agent.getSnapshot().root.body[0]).id;
    const m = await agent.applyCommand({
      type: "docx:set-paragraph-list",
      payload: { paragraphId: id, numId: 1, ilvl: 0 },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-target");
  });

  it("set-paragraph-list rejects ilvl < 0 with invalid-payload", async () => {
    const agent = await loadAgentWithNumbering([{ text: "x" }]);
    const id = paraOf(agent.getSnapshot().root.body[0]).id;
    const m = await agent.applyCommand({
      type: "docx:set-paragraph-list",
      payload: { paragraphId: id, numId: 1, ilvl: -1 },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-payload");
  });

  it("remove-paragraph-list happy path; dirty.body set", async () => {
    const agent = await loadAgentWithNumbering([{ text: "list", numId: 1, ilvl: 0 }]);
    const id = paraOf(agent.getSnapshot().root.body[0]).id;
    expect(paraOf(agent.getSnapshot().root.body[0]).properties.numbering).toEqual({ numId: 1, ilvl: 0 });
    const m = await agent.applyCommand({
      type: "docx:remove-paragraph-list",
      payload: { paragraphId: id },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const snap = agent.getSnapshot();
    expect(snap.dirty.body).toBe(true);
    expect(paraOf(snap.root.body[0]).properties.numbering).toBeUndefined();
    // Serialized doc no longer contains a w:numPr in the affected paragraph.
    const out = await serializeDocx(snap);
    const reloaded = await ooxml.OoxmlContainer.load(out);
    const xml = reloaded.readText("word/document.xml");
    expect(xml).not.toContain("<w:numPr");
  });

  it("remove-paragraph-list rejects with not-applicable when paragraph has no numbering set", async () => {
    const agent = await loadAgentWithNumbering([{ text: "plain" }]);
    const id = paraOf(agent.getSnapshot().root.body[0]).id;
    const m = await agent.applyCommand({
      type: "docx:remove-paragraph-list",
      payload: { paragraphId: id },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("not-applicable");
  });

  it("set-paragraph-list round-trips through parse(serialize(snapshot))", async () => {
    const agent = await loadAgentWithNumbering([{ text: "Item" }]);
    const id = paraOf(agent.getSnapshot().root.body[0]).id;
    await agent.applyCommand({
      type: "docx:set-paragraph-list",
      payload: { paragraphId: id, numId: 2, ilvl: 0 },
      source: "human",
    });
    const out = await serializeDocx(agent.getSnapshot());
    const snap2 = await parseDocx(out, { idMinter: deterministicIdMinter("z") });
    const p = paraOf(snap2.root.body[0]);
    expect(p.properties.numbering).toEqual({ numId: 2, ilvl: 0 });
  });
});
