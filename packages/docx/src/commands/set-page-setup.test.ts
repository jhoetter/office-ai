import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import JSZip from "jszip";
import { DocxAgent } from "../agent/agent.js";
import { DEFAULT_DOC_ROOT_ATTRS, makeSyntheticDocx, plainDocxXml } from "../test-utils/synthetic.js";

async function makeAgent(): Promise<DocxAgent> {
  const buf = await makeSyntheticDocx({
    documentXml: plainDocxXml([{ text: "Hello world" }]),
  });
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

function syntheticWithMargins(
  pgSz: { w: number; h: number },
  pgMar: { top: number; right: number; bottom: number; left: number; header: number; footer: number }
): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body><w:p><w:r><w:t xml:space="preserve">Body para</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="${pgSz.w}" w:h="${pgSz.h}"/><w:pgMar w:top="${pgMar.top}" w:right="${pgMar.right}" w:bottom="${pgMar.bottom}" w:left="${pgMar.left}" w:header="${pgMar.header}" w:footer="${pgMar.footer}"/></w:sectPr></w:body></w:document>`;
}

describe("docx:set-page-setup", () => {
  it("updates pgMar on the trailing implicit section", async () => {
    const agent = await makeAgent();
    await agent.applyCommand({
      type: "docx:set-page-setup",
      payload: {
        paragraphIndex: 0,
        pgMar: { left: 720, right: 720 },
      },
      source: "human",
    });
    const snap = agent.getSnapshot();
    const section = snap.root.body.find((b) => b.kind === "section-break");
    expect(section).toBeDefined();
    if (section?.kind !== "section-break") throw new Error("expected section-break");
    expect(section.properties.pgMar?.left).toBe(720);
    expect(section.properties.pgMar?.right).toBe(720);
    // Untouched fields preserved.
    expect(section.properties.pgMar?.top).toBe(1440);
  });

  it("updates pgSz orientation and dimensions", async () => {
    const buf = await makeSyntheticDocx({
      documentXml: syntheticWithMargins(
        { w: 12240, h: 15840 },
        { top: 1440, right: 1440, bottom: 1440, left: 1440, header: 720, footer: 720 }
      ),
    });
    const agent = await DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    await agent.applyCommand({
      type: "docx:set-page-setup",
      payload: {
        paragraphIndex: 0,
        pgSz: { w: 16838, h: 11906, orient: "landscape" },
      },
      source: "human",
    });
    const snap = agent.getSnapshot();
    const section = snap.root.body.find((b) => b.kind === "section-break");
    if (section?.kind !== "section-break") throw new Error("expected section-break");
    expect(section.properties.pgSz?.w).toBe(16838);
    expect(section.properties.pgSz?.h).toBe(11906);
    expect(section.properties.pgSz?.orient).toBe("landscape");
  });

  it("clamps margins so printable width stays >= 360 twips", async () => {
    const agent = await makeAgent();
    await agent.applyCommand({
      type: "docx:set-page-setup",
      payload: {
        paragraphIndex: 0,
        pgMar: { left: 12000, right: 0 },
      },
      source: "human",
    });
    const snap = agent.getSnapshot();
    const section = snap.root.body.find((b) => b.kind === "section-break");
    if (section?.kind !== "section-break") throw new Error("expected section-break");
    const left = section.properties.pgMar?.left ?? 0;
    const right = section.properties.pgMar?.right ?? 0;
    const pageW = section.properties.pgSz?.w ?? 12240;
    expect(pageW - left - right).toBeGreaterThanOrEqual(360);
  });

  it("is a no-op when geometry is unchanged", async () => {
    const agent = await makeAgent();
    const before = agent.getSnapshot().revision;
    await agent.applyCommand({
      type: "docx:set-page-setup",
      payload: {
        paragraphIndex: 0,
        pgMar: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      },
      source: "human",
    });
    expect(agent.getSnapshot().revision).toBe(before);
  });

  it("re-serialises into a docx that still parses", async () => {
    const agent = await makeAgent();
    await agent.applyCommand({
      type: "docx:set-page-setup",
      payload: {
        paragraphIndex: 0,
        pgSz: { w: 16838, h: 11906, orient: "landscape" },
        pgMar: { left: 720, right: 720, top: 1080, bottom: 1080 },
      },
      source: "human",
    });
    const out = await agent.exportFile();
    const z = await JSZip.loadAsync(out);
    const doc = await z.file("word/document.xml")!.async("string");
    expect(doc).toContain('w:w="16838"');
    expect(doc).toContain('w:h="11906"');
    expect(doc).toMatch(/w:left="720"/);
    expect(doc).toMatch(/w:right="720"/);

    const reparsed = await DocxAgent.fromBuffer(out, { idMinter: deterministicIdMinter() });
    const section = reparsed.getSnapshot().root.body.find((b) => b.kind === "section-break");
    if (section?.kind !== "section-break") throw new Error("expected section-break");
    expect(section.properties.pgSz?.w).toBe(16838);
    expect(section.properties.pgMar?.left).toBe(720);
  });

  it("rejects empty payload", async () => {
    const agent = await makeAgent();
    const m = await agent.applyCommand({
      type: "docx:set-page-setup",
      payload: { paragraphIndex: 0 },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-payload");
  });
});
