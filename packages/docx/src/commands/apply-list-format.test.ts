import { describe, expect, it } from "vitest";
import { deterministicIdMinter, ooxml } from "@officeai/core";
import { DocxAgent } from "../agent/agent.js";
import { parseDocx } from "../parser/parse.js";
import { serializeDocx } from "../serializer/serialize.js";
import type { BlockNode, Paragraph } from "../model/types.js";
import { DEFAULT_DOC_ROOT_ATTRS, escapeXml, makeSyntheticDocx } from "../test-utils/synthetic.js";

function plainDoc(paragraphs: ReadonlyArray<string>): string {
  const ps = paragraphs
    .map((t) => `<w:p><w:r><w:t xml:space="preserve">${escapeXml(t)}</w:t></w:r></w:p>`)
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>${ps}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
}

async function loadAgent(paragraphs: ReadonlyArray<string>): Promise<DocxAgent> {
  const buf = await makeSyntheticDocx({ documentXml: plainDoc(paragraphs) });
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

function paraOf(block: BlockNode): Paragraph {
  if (block.kind !== "paragraph") throw new Error("expected paragraph");
  return block;
}

describe("docx:apply-list-format (B7)", () => {
  it("auto-mints a bullet abstract+num pair when the doc has no numbering part", async () => {
    const agent = await loadAgent(["item one"]);
    const before = agent.getSnapshot();
    expect(before.root.numbering).toBeUndefined();
    const para = paraOf(before.root.body[0]);

    const m = await agent.applyCommand({
      type: "docx:apply-list-format",
      payload: { paragraphId: para.id, format: "bullet" },
      source: "human",
    });
    expect(m.status).toBe("approved");

    const snap = agent.getSnapshot();
    expect(snap.root.numbering).toBeTruthy();
    expect(snap.root.numbering!.abstractNums.size).toBe(1);
    expect(snap.root.numbering!.nums.size).toBe(1);
    const onlyAbstract = Array.from(snap.root.numbering!.abstractNums.values())[0];
    expect(onlyAbstract.levels[0].numFmt).toBe("bullet");
    expect(snap.dirty.numbering).toBe(true);
    expect(snap.dirty.body).toBe(true);
  });

  it("auto-mints a decimal abstract+num pair on first numbered list", async () => {
    const agent = await loadAgent(["one", "two"]);
    const para = paraOf(agent.getSnapshot().root.body[0]);
    await agent.applyCommand({
      type: "docx:apply-list-format",
      payload: { paragraphId: para.id, format: "decimal" },
      source: "human",
    });
    const snap = agent.getSnapshot();
    const onlyAbstract = Array.from(snap.root.numbering!.abstractNums.values())[0];
    expect(onlyAbstract.levels[0].numFmt).toBe("decimal");
    expect(onlyAbstract.levels[0].lvlText).toBe("%1.");
  });

  it("reuses an existing matching <w:num> instead of minting a new one", async () => {
    const agent = await loadAgent(["one", "two"]);
    const p0 = paraOf(agent.getSnapshot().root.body[0]).id;
    const p1 = paraOf(agent.getSnapshot().root.body[1]).id;
    await agent.applyCommand({
      type: "docx:apply-list-format",
      payload: { paragraphId: p0, format: "bullet" },
      source: "human",
    });
    const after1 = agent.getSnapshot();
    const numCountAfter1 = after1.root.numbering!.nums.size;

    await agent.applyCommand({
      type: "docx:apply-list-format",
      payload: { paragraphId: p1, format: "bullet" },
      source: "human",
    });
    const after2 = agent.getSnapshot();
    expect(after2.root.numbering!.nums.size).toBe(numCountAfter1);
  });

  it("registers word/numbering.xml + relationship + content-type override on save", async () => {
    const agent = await loadAgent(["item"]);
    const para = paraOf(agent.getSnapshot().root.body[0]);
    await agent.applyCommand({
      type: "docx:apply-list-format",
      payload: { paragraphId: para.id, format: "bullet" },
      source: "human",
    });
    const out = await serializeDocx(agent.getSnapshot());
    const reloaded = await ooxml.OoxmlContainer.load(out);
    expect(reloaded.has("word/numbering.xml")).toBe(true);
    const relsXml = reloaded.readText("word/_rels/document.xml.rels");
    expect(relsXml).toContain("/relationships/numbering");
    expect(relsXml).toContain("numbering.xml");
    const ctXml = reloaded.readText("[Content_Types].xml");
    expect(ctXml).toContain("/word/numbering.xml");
    expect(ctXml).toContain("wordprocessingml.numbering+xml");
  });

  it("survives a save → reparse round-trip with the paragraph still pointing at the new list", async () => {
    const agent = await loadAgent(["item"]);
    const para = paraOf(agent.getSnapshot().root.body[0]);
    await agent.applyCommand({
      type: "docx:apply-list-format",
      payload: { paragraphId: para.id, format: "decimal", ilvl: 0 },
      source: "human",
    });
    const out = await serializeDocx(agent.getSnapshot());
    const reparsed = await parseDocx(out, { idMinter: deterministicIdMinter("z") });
    const p = paraOf(reparsed.root.body[0]);
    expect(p.properties.numbering).toBeTruthy();
    expect(p.properties.numbering?.ilvl).toBe(0);
    expect(reparsed.root.numbering).toBeTruthy();
    const num = reparsed.root.numbering!.nums.get(p.properties.numbering!.numId);
    expect(num).toBeTruthy();
    const abs = reparsed.root.numbering!.abstractNums.get(num!.abstractNumId);
    expect(abs?.levels[0].numFmt).toBe("decimal");
  });

  it("rejects invalid payloads (empty paragraphId, bad format, negative ilvl)", async () => {
    const agent = await loadAgent(["x"]);
    const para = paraOf(agent.getSnapshot().root.body[0]);
    const a = await agent.applyCommand({
      type: "docx:apply-list-format",
      payload: { paragraphId: "", format: "bullet" },
      source: "human",
    });
    expect(a.status).toBe("rejected");
    const b = await agent.applyCommand({
      type: "docx:apply-list-format",
      payload: { paragraphId: para.id, format: "alien" as "bullet" },
      source: "human",
    });
    expect(b.status).toBe("rejected");
    const c = await agent.applyCommand({
      type: "docx:apply-list-format",
      payload: { paragraphId: para.id, format: "bullet", ilvl: -1 },
      source: "human",
    });
    expect(c.status).toBe("rejected");
  });
});
