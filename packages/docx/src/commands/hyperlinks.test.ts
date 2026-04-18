import { describe, expect, it } from "vitest";
import { deterministicIdMinter, ooxml, sha256Hex } from "@officeai/core";
import { DocxAgent } from "../agent/agent.js";
import { parseDocx } from "../parser/parse.js";
import { serializeDocx } from "../serializer/serialize.js";
import type { Hyperlink, Paragraph } from "../model/types.js";
import { DEFAULT_DOC_ROOT_ATTRS, makeSyntheticDocx } from "../test-utils/synthetic.js";

const HYPERLINK_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

function bodyXml(inner: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>${inner}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
}

function plainParaXml(text: string): string {
  return `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
}

async function plainAgent(text: string): Promise<DocxAgent> {
  const buf = await makeSyntheticDocx({
    documentXml: bodyXml(plainParaXml(text)),
    extra: { "[Content_Types].xml": CONTENT_TYPES },
  });
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

async function existingHyperlinkBuf(): Promise<ArrayBuffer> {
  const docXml = bodyXml(
    `<w:p><w:r><w:t xml:space="preserve">visit </w:t></w:r><w:hyperlink r:id="rId5"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t xml:space="preserve">our site</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve"> today</w:t></w:r></w:p>`
  );
  const docRels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId5" Type="${HYPERLINK_REL_TYPE}" Target="https://example.com/" TargetMode="External"/>
</Relationships>`;
  return makeSyntheticDocx({
    documentXml: docXml,
    extra: {
      "[Content_Types].xml": CONTENT_TYPES,
      "word/_rels/document.xml.rels": docRels,
    },
  });
}

function paraOf(b: { kind: string }): Paragraph {
  if (b.kind !== "paragraph") throw new Error("expected paragraph");
  return b as Paragraph;
}

describe("docx hyperlinks — parser + commands (P1.4 / W11)", () => {
  it("round-trips a fixture with an existing hyperlink structurally through parse(serialize(s))", async () => {
    const buf = await existingHyperlinkBuf();
    const snap = await parseDocx(buf, { idMinter: deterministicIdMinter("a") });
    const out = await serializeDocx(snap);
    const snap2 = await parseDocx(out, { idMinter: deterministicIdMinter("b") });
    const p1 = paraOf(snap.root.body[0]);
    const p2 = paraOf(snap2.root.body[0]);
    expect(p2.children.length).toBe(p1.children.length);
    const hl1 = p1.children.find((c) => c.kind === "hyperlink") as Hyperlink;
    const hl2 = p2.children.find((c) => c.kind === "hyperlink") as Hyperlink;
    expect(hl2.relationshipId).toBe(hl1.relationshipId);
    expect(hl2.children.length).toBe(hl1.children.length);
  });

  it("byte-preserves document.xml + document.xml.rels when no hyperlink command runs", async () => {
    const buf = await existingHyperlinkBuf();
    const snap = await parseDocx(buf, { idMinter: deterministicIdMinter() });
    const out = await serializeDocx(snap);
    const reloaded = await ooxml.OoxmlContainer.load(out);
    for (const path of ["word/document.xml", "word/_rels/document.xml.rels"]) {
      const before = sha256Hex(snap.container.readBytes(path));
      const after = sha256Hex(reloaded.readBytes(path));
      expect(after, `${path} should be byte-identical`).toBe(before);
    }
  });

  it("insert-hyperlink: external URL happy path mints a rel and sets dirty.relationships", async () => {
    const agent = await plainAgent("Click the link here");
    const id = paraOf(agent.getSnapshot().root.body[0]).id;
    const m = await agent.applyCommand({
      type: "docx:insert-hyperlink",
      payload: { paragraphId: id, range: { start: 10, end: 14 }, url: "https://example.com/" },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const snap = agent.getSnapshot();
    expect(snap.dirty.body).toBe(true);
    expect(snap.dirty.relationships.has("word/document.xml")).toBe(true);
    const p = paraOf(snap.root.body[0]);
    const hl = p.children.find((c) => c.kind === "hyperlink") as Hyperlink;
    expect(hl).toBeTruthy();
    expect(hl.relationshipId).toBeTruthy();
    expect(hl.anchor).toBeUndefined();
    const docRels = snap.root.relationships.get("word/document.xml") ?? [];
    const rel = docRels.find((r) => r.id === hl.relationshipId);
    expect(rel?.type).toBe(HYPERLINK_REL_TYPE);
    expect(rel?.target).toBe("https://example.com/");
    expect(rel?.targetMode).toBe("External");
  });

  it("insert-hyperlink: anchor (internal) does NOT mint a rel", async () => {
    const agent = await plainAgent("Jump to chapter two");
    const id = paraOf(agent.getSnapshot().root.body[0]).id;
    const beforeRels = agent.getSnapshot().root.relationships.get("word/document.xml") ?? [];
    const m = await agent.applyCommand({
      type: "docx:insert-hyperlink",
      payload: { paragraphId: id, range: { start: 8, end: 15 }, anchor: "chapter-two" },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const snap = agent.getSnapshot();
    expect(snap.dirty.relationships.has("word/document.xml")).toBe(false);
    const afterRels = snap.root.relationships.get("word/document.xml") ?? [];
    expect(afterRels.length).toBe(beforeRels.length);
    const p = paraOf(snap.root.body[0]);
    const hl = p.children.find((c) => c.kind === "hyperlink") as Hyperlink;
    expect(hl.anchor).toBe("chapter-two");
    expect(hl.relationshipId).toBeUndefined();
  });

  it("insert-hyperlink: same URL twice reuses the existing relationship (no duplicate)", async () => {
    const agent = await plainAgent("alpha beta gamma delta");
    const id = paraOf(agent.getSnapshot().root.body[0]).id;
    await agent.applyCommand({
      type: "docx:insert-hyperlink",
      payload: { paragraphId: id, range: { start: 0, end: 5 }, url: "https://example.com/x" },
      source: "human",
    });
    const id2 = paraOf(agent.getSnapshot().root.body[0]).id;
    await agent.applyCommand({
      type: "docx:insert-hyperlink",
      payload: { paragraphId: id2, range: { start: 17, end: 22 }, url: "https://example.com/x" },
      source: "human",
    });
    const snap = agent.getSnapshot();
    const docRels = snap.root.relationships.get("word/document.xml") ?? [];
    const matching = docRels.filter(
      (r) => r.target === "https://example.com/x" && r.type === HYPERLINK_REL_TYPE
    );
    expect(matching.length).toBe(1);
    const p = paraOf(snap.root.body[0]);
    const hyperlinks = p.children.filter((c) => c.kind === "hyperlink") as Hyperlink[];
    expect(hyperlinks.length).toBe(2);
    expect(hyperlinks[0].relationshipId).toBe(hyperlinks[1].relationshipId);
  });

  it("insert-hyperlink: range straddling an existing hyperlink rejects with invalid-position", async () => {
    const buf = await existingHyperlinkBuf();
    const agent = await DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    const id = paraOf(agent.getSnapshot().root.body[0]).id;
    // Paragraph text is "visit our site today" (20 chars).
    // The existing hyperlink covers "our site" (offsets 6..14).
    const m = await agent.applyCommand({
      type: "docx:insert-hyperlink",
      payload: { paragraphId: id, range: { start: 0, end: 10 }, url: "https://other.example/" },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-position");
  });

  it("insert-hyperlink: missing both url and anchor → invalid-payload", async () => {
    const agent = await plainAgent("hello world");
    const id = paraOf(agent.getSnapshot().root.body[0]).id;
    const m = await agent.applyCommand({
      type: "docx:insert-hyperlink",
      payload: { paragraphId: id, range: { start: 0, end: 5 } },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-payload");
  });

  it("insert-hyperlink: both url and anchor → invalid-payload", async () => {
    const agent = await plainAgent("hello world");
    const id = paraOf(agent.getSnapshot().root.body[0]).id;
    const m = await agent.applyCommand({
      type: "docx:insert-hyperlink",
      payload: { paragraphId: id, range: { start: 0, end: 5 }, url: "https://x.test/", anchor: "y" },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-payload");
  });

  it("insert-hyperlink: range outside paragraph length → invalid-position", async () => {
    const agent = await plainAgent("short"); // length 5
    const id = paraOf(agent.getSnapshot().root.body[0]).id;
    const m = await agent.applyCommand({
      type: "docx:insert-hyperlink",
      payload: { paragraphId: id, range: { start: 0, end: 99 }, url: "https://x.test/" },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-position");
  });

  it("insert-hyperlink: url is not well-formed → invalid-payload", async () => {
    const agent = await plainAgent("hello world");
    const id = paraOf(agent.getSnapshot().root.body[0]).id;
    const m = await agent.applyCommand({
      type: "docx:insert-hyperlink",
      payload: { paragraphId: id, range: { start: 0, end: 5 }, url: "not a url" },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-payload");
  });

  it("remove-hyperlink: happy path; rel removed when sole reference", async () => {
    const buf = await existingHyperlinkBuf();
    const agent = await DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    const p = paraOf(agent.getSnapshot().root.body[0]);
    const hl = p.children.find((c) => c.kind === "hyperlink") as Hyperlink;
    expect(hl).toBeTruthy();
    const m = await agent.applyCommand({
      type: "docx:remove-hyperlink",
      payload: { hyperlinkId: hl.id },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const snap = agent.getSnapshot();
    expect(snap.dirty.body).toBe(true);
    expect(snap.dirty.relationships.has("word/document.xml")).toBe(true);
    const p2 = paraOf(snap.root.body[0]);
    expect(p2.children.find((c) => c.kind === "hyperlink")).toBeUndefined();
    const docRels = snap.root.relationships.get("word/document.xml") ?? [];
    expect(docRels.find((r) => r.id === hl.relationshipId)).toBeUndefined();
  });

  it("remove-hyperlink: rel preserved when another hyperlink still references the same target", async () => {
    const agent = await plainAgent("alpha beta gamma delta");
    const id = paraOf(agent.getSnapshot().root.body[0]).id;
    await agent.applyCommand({
      type: "docx:insert-hyperlink",
      payload: { paragraphId: id, range: { start: 0, end: 5 }, url: "https://shared.test/" },
      source: "human",
    });
    const idB = paraOf(agent.getSnapshot().root.body[0]).id;
    await agent.applyCommand({
      type: "docx:insert-hyperlink",
      payload: { paragraphId: idB, range: { start: 17, end: 22 }, url: "https://shared.test/" },
      source: "human",
    });
    const p = paraOf(agent.getSnapshot().root.body[0]);
    const hyperlinks = p.children.filter((c) => c.kind === "hyperlink") as Hyperlink[];
    expect(hyperlinks.length).toBe(2);
    const sharedRelId = hyperlinks[0].relationshipId;
    expect(hyperlinks[1].relationshipId).toBe(sharedRelId);

    const beforeDirty = agent.getSnapshot().dirty.relationships.has("word/document.xml");
    expect(beforeDirty).toBe(true); // already dirty from inserts

    // Remove the first one — the rel should remain because the second still uses it.
    await agent.applyCommand({
      type: "docx:remove-hyperlink",
      payload: { hyperlinkId: hyperlinks[0].id },
      source: "human",
    });
    const snap = agent.getSnapshot();
    const docRels = snap.root.relationships.get("word/document.xml") ?? [];
    expect(docRels.find((r) => r.id === sharedRelId)).toBeTruthy();
    const p2 = paraOf(snap.root.body[0]);
    const hyperlinks2 = p2.children.filter((c) => c.kind === "hyperlink") as Hyperlink[];
    expect(hyperlinks2.length).toBe(1);
    expect(hyperlinks2[0].relationshipId).toBe(sharedRelId);
  });

  it("remove-hyperlink: unknown hyperlinkId → unknown-target", async () => {
    const agent = await plainAgent("nothing here");
    const m = await agent.applyCommand({
      type: "docx:remove-hyperlink",
      payload: { hyperlinkId: "no-such-id" },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-target");
  });

  it("insert + serialize round-trip: external URL hyperlink survives parse(serialize(s))", async () => {
    const agent = await plainAgent("Click here please");
    const id = paraOf(agent.getSnapshot().root.body[0]).id;
    await agent.applyCommand({
      type: "docx:insert-hyperlink",
      payload: { paragraphId: id, range: { start: 0, end: 10 }, url: "https://roundtrip.test/" },
      source: "human",
    });
    const out = await serializeDocx(agent.getSnapshot());
    const snap2 = await parseDocx(out, { idMinter: deterministicIdMinter("z") });
    const p = paraOf(snap2.root.body[0]);
    const hl = p.children.find((c) => c.kind === "hyperlink") as Hyperlink;
    expect(hl).toBeTruthy();
    expect(hl.relationshipId).toBeTruthy();
    const rels = snap2.root.relationships.get("word/document.xml") ?? [];
    const rel = rels.find((r) => r.id === hl.relationshipId);
    expect(rel?.target).toBe("https://roundtrip.test/");
    expect(rel?.targetMode).toBe("External");
  });
});
