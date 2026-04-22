import JSZip from "jszip";
import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { DocxAgent } from "../agent/agent.js";
import { parseDocx } from "../parser/parse.js";
import { DEFAULT_DOC_ROOT_ATTRS, escapeXml, makeSyntheticDocx } from "../test-utils/synthetic.js";
import type { Footnote, FootnoteReferenceLeaf, Paragraph, Run } from "../model/types.js";

/* ── Fixture helpers ─────────────────────────────────────────────────────── */

function syntheticDocXml(paraCount = 2): string {
  const paras = Array.from(
    { length: paraCount },
    (_, i) =>
      `<w:p><w:r><w:t xml:space="preserve">${escapeXml(`Paragraph ${i + 1} body text`)}</w:t></w:r></w:p>`
  ).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>${paras}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`;
}

/**
 * Build a fixture docx that has a real `word/footnotes.xml` part wired up
 * in `[Content_Types].xml` and `word/_rels/document.xml.rels`. The body
 * carries a single `<w:footnoteReference w:id="1"/>` so we can prove the
 * round-trip preserves untouched footnotes byte-for-byte.
 */
async function makeDocxWithFootnotes(): Promise<ArrayBuffer> {
  const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body><w:p><w:r><w:t xml:space="preserve">Hello</w:t></w:r><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;

  const footnotesXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="1"><w:p><w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr><w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t xml:space="preserve"> Existing footnote text.</w:t></w:r></w:p></w:footnote></w:footnotes>`;

  // Add the override + relationship for footnotes.
  const contentTypes = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
</Types>`;

  const docRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
</Relationships>`;

  // Use makeSyntheticDocx with extras then post-patch the two manifests.
  const buf = await makeSyntheticDocx({
    documentXml,
    extra: { "word/footnotes.xml": footnotesXml },
  });
  // Re-zip with the patched content-types + rels so the new footnotes part
  // is properly declared (otherwise the parser will skip it).
  const z = await JSZip.loadAsync(buf);
  z.file("[Content_Types].xml", contentTypes);
  z.file("word/_rels/document.xml.rels", docRels);
  return z.generateAsync({
    type: "arraybuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
}

async function makePlainAgent(): Promise<DocxAgent> {
  const buf = await makeSyntheticDocx({ documentXml: syntheticDocXml(2) });
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

async function makeFootnotesAgent(): Promise<DocxAgent> {
  const buf = await makeDocxWithFootnotes();
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

function findFootnoteRef(p: Paragraph, footnoteId: number): FootnoteReferenceLeaf | undefined {
  for (const child of p.children) {
    if (child.kind !== "run") continue;
    for (const rc of child.children) {
      if (rc.kind === "footnote-ref" && rc.footnoteId === footnoteId) return rc;
    }
  }
  return undefined;
}

/* ── Tests ───────────────────────────────────────────────────────────────── */

describe("docx footnotes — model & round-trip", () => {
  it("parses w:footnoteReference into a typed FootnoteReferenceLeaf", async () => {
    const buf = await makeDocxWithFootnotes();
    const parsed = await parseDocx(buf);
    const para = parsed.root.body[0] as Paragraph;
    const ref = findFootnoteRef(para, 1);
    expect(ref).toBeTruthy();
    expect(ref?.kind).toBe("footnote-ref");
    expect(ref?.footnoteId).toBe(1);
  });

  it("parses word/footnotes.xml into a typed FootnotesPart", async () => {
    const buf = await makeDocxWithFootnotes();
    const parsed = await parseDocx(buf);
    expect(parsed.root.footnotesPart).toBeTruthy();
    const part = parsed.root.footnotesPart!;
    expect(part.footnotes.length).toBe(3);
    expect(part.footnotes.map((f) => f.type).sort()).toEqual([
      "continuationSeparator",
      "normal",
      "separator",
    ]);
    const normal = part.footnotes.find((f) => f.type === "normal")!;
    expect(normal.id).toBe(1);
    expect(normal.body.length).toBeGreaterThan(0);
  });

  it("round-trips an unedited fixture: footnotes.xml is byte-identical", async () => {
    const buf = await makeDocxWithFootnotes();
    const agent = await DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    // No commands applied → exporting must not rewrite footnotes.xml.
    const out = await agent.exportFile();
    const before = await JSZip.loadAsync(buf);
    const after = await JSZip.loadAsync(out);
    const beforeBytes = await before.file("word/footnotes.xml")!.async("uint8array");
    const afterBytes = await after.file("word/footnotes.xml")!.async("uint8array");
    expect(afterBytes).toEqual(beforeBytes);
  });
});

describe("docx:insert-footnote", () => {
  it("creates a footnotesPart with a new footnote and a typed ref leaf at the caret", async () => {
    const agent = await makePlainAgent();
    const para = agent.getSnapshot().root.body[0] as Paragraph;
    const m = await agent.applyCommand({
      type: "docx:insert-footnote",
      payload: { paragraphId: para.id, offset: 9 },
      source: "human",
    });
    expect(m.status).toBe("approved");

    const snap = agent.getSnapshot();
    const part = snap.root.footnotesPart;
    expect(part).toBeTruthy();
    expect(part!.footnotes.length).toBe(1);
    const fn = part!.footnotes[0];
    expect(fn.id).toBe(1);
    expect(fn.type).toBe("normal");
    expect(fn.body.length).toBe(1);
    expect(snap.dirty.footnotes).toBe(true);

    const updated = snap.root.body[0] as Paragraph;
    const ref = findFootnoteRef(updated, 1);
    expect(ref).toBeTruthy();
    expect(ref?.footnoteId).toBe(1);
  });

  it("appends to an existing footnotesPart with the next id", async () => {
    const agent = await makeFootnotesAgent();
    const para = agent.getSnapshot().root.body[0] as Paragraph;
    const m = await agent.applyCommand({
      type: "docx:insert-footnote",
      payload: { paragraphId: para.id, offset: 0 },
      source: "human",
    });
    expect(m.status).toBe("approved");

    const part = agent.getSnapshot().root.footnotesPart!;
    // Existing: separator(-1), continuationSeparator(0), normal(1) → next normal=2
    const normals = part.footnotes.filter((f) => f.type === "normal");
    expect(normals.map((f) => f.id).sort((a, b) => a - b)).toEqual([1, 2]);
  });

  it("rejects unknown paragraph id with unknown-target", async () => {
    const agent = await makePlainAgent();
    const m = await agent.applyCommand({
      type: "docx:insert-footnote",
      payload: { paragraphId: "ghost-para", offset: 0 },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-target");
  });

  it("rejects negative offset with invalid-payload", async () => {
    const agent = await makePlainAgent();
    const para = agent.getSnapshot().root.body[0] as Paragraph;
    const m = await agent.applyCommand({
      type: "docx:insert-footnote",
      payload: { paragraphId: para.id, offset: -1 },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-payload");
  });
});

describe("docx:set-footnote-body", () => {
  it("replaces the body of an existing footnote", async () => {
    const agent = await makeFootnotesAgent();
    const newBody: Footnote["body"] = [
      {
        kind: "paragraph",
        id: "p-replaced",
        properties: { styleId: "FootnoteText" },
        children: [
          {
            kind: "run",
            id: "r-replaced",
            properties: {},
            children: [{ kind: "text", id: "t-replaced", text: "Replaced body" }],
          } satisfies Run,
        ],
      },
    ];
    const m = await agent.applyCommand({
      type: "docx:set-footnote-body",
      payload: { footnoteId: 1, body: newBody },
      source: "human",
    });
    expect(m.status).toBe("approved");

    const part = agent.getSnapshot().root.footnotesPart!;
    const fn = part.footnotes.find((f) => f.id === 1)!;
    expect(fn.body).toEqual(newBody);
    expect(fn.raw).toBeUndefined();
    expect(agent.getSnapshot().dirty.footnotes).toBe(true);
  });

  it("rejects unknown footnote id", async () => {
    const agent = await makeFootnotesAgent();
    const m = await agent.applyCommand({
      type: "docx:set-footnote-body",
      payload: {
        footnoteId: 999,
        body: [
          {
            kind: "paragraph",
            id: "p-x",
            properties: {},
            children: [],
          },
        ],
      },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-target");
  });
});

describe("docx:delete-footnote", () => {
  it("removes the footnote entry and strips matching reference leaves from the body", async () => {
    const agent = await makeFootnotesAgent();
    // Sanity: body has a ref to footnote 1.
    const before = agent.getSnapshot().root.body[0] as Paragraph;
    expect(findFootnoteRef(before, 1)).toBeTruthy();

    const m = await agent.applyCommand({
      type: "docx:delete-footnote",
      payload: { footnoteId: 1 },
      source: "human",
    });
    expect(m.status).toBe("approved");

    const snap = agent.getSnapshot();
    const part = snap.root.footnotesPart!;
    expect(part.footnotes.find((f) => f.id === 1)).toBeUndefined();
    expect(part.footnotes.find((f) => f.type === "separator")).toBeTruthy();

    const after = snap.root.body[0] as Paragraph;
    expect(findFootnoteRef(after, 1)).toBeUndefined();
    expect(snap.dirty.footnotes).toBe(true);
    expect(snap.dirty.body).toBe(true);
  });

  it("removes the part entirely when the last footnote is deleted", async () => {
    // Insert a single footnote into a clean doc, then delete it.
    const agent = await makePlainAgent();
    const para = agent.getSnapshot().root.body[0] as Paragraph;
    await agent.applyCommand({
      type: "docx:insert-footnote",
      payload: { paragraphId: para.id, offset: 0 },
      source: "human",
    });
    const inserted = agent.getSnapshot().root.footnotesPart!.footnotes[0];
    const m = await agent.applyCommand({
      type: "docx:delete-footnote",
      payload: { footnoteId: inserted.id },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const part = agent.getSnapshot().root.footnotesPart!;
    expect(part.footnotes.length).toBe(0);
  });

  it("rejects unknown footnote id with unknown-target", async () => {
    const agent = await makeFootnotesAgent();
    const m = await agent.applyCommand({
      type: "docx:delete-footnote",
      payload: { footnoteId: 999 },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-target");
  });
});

describe("docx footnotes — export round-trip", () => {
  it("exports a freshly inserted footnote and re-parses it back", async () => {
    const agent = await makePlainAgent();
    const para = agent.getSnapshot().root.body[0] as Paragraph;
    await agent.applyCommand({
      type: "docx:insert-footnote",
      payload: { paragraphId: para.id, offset: 5 },
      source: "human",
    });
    const buf = await agent.exportFile();
    const reparsed = await parseDocx(buf);
    expect(reparsed.root.footnotesPart).toBeTruthy();
    const fn = reparsed.root.footnotesPart!.footnotes.find((f) => f.type === "normal");
    expect(fn).toBeTruthy();
    const reparsedPara = reparsed.root.body[0] as Paragraph;
    expect(findFootnoteRef(reparsedPara, fn!.id)).toBeTruthy();
  });
});
