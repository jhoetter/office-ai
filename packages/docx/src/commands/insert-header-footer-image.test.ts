import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { DocxAgent } from "../agent/agent.js";
import { parseDocx } from "../parser/parse.js";
import { DEFAULT_DOC_ROOT_ATTRS, escapeXml, makeSyntheticDocx } from "../test-utils/synthetic.js";
import type { InlineImageDrawing, Paragraph, Run } from "../model/types.js";

const HDR_NS_ATTRS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;

const HEADER_PART = "word/header1.xml";

function syntheticDocXml(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body><w:p><w:r><w:t xml:space="preserve">Body</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdH"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`;
}

function syntheticHeaderXml(text: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr ${HDR_NS_ATTRS}><w:p><w:r><w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p></w:hdr>`;
}

function syntheticDocRels(): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rIdH" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
</Relationships>`;
}

async function makeAgent(): Promise<DocxAgent> {
  const buf = await makeSyntheticDocx({
    documentXml: syntheticDocXml(),
    extra: {
      "word/_rels/document.xml.rels": syntheticDocRels(),
      [HEADER_PART]: syntheticHeaderXml("Header"),
    },
  });
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

// 1×1 PNG (transparent). Real PNG bytes — important because the
// command SHA-256s them for media de-dup.
const PNG_1x1 = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52, 0x00, 0x00,
  0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4, 0x89, 0x00, 0x00, 0x00,
  0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d,
  0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
]);

describe("docx:insert-header-footer-image", () => {
  it("appends a new paragraph with the image when paragraphIndex is omitted", async () => {
    const agent = await makeAgent();
    const m = await agent.applyCommand({
      type: "docx:insert-header-footer-image",
      payload: {
        partPath: HEADER_PART,
        data: PNG_1x1,
        mimeType: "image/png",
        width: 64,
        height: 64,
      },
      source: "human",
    });
    expect(m.status).toBe("approved");

    const snap = agent.getSnapshot();
    const header = snap.root.headersAndFooters.find((p) => p.partPath === HEADER_PART);
    expect(header).toBeTruthy();
    expect(header!.body.length).toBe(2);
    const last = header!.body[header!.body.length - 1] as Paragraph;
    expect(last.kind).toBe("paragraph");
    const imgRun = last.children.find(
      (c): c is Run => c.kind === "run" && c.children.some((rc) => rc.kind === "drawing")
    );
    expect(imgRun).toBeTruthy();
    const drawing = imgRun!.children.find((rc) => rc.kind === "drawing") as InlineImageDrawing;
    expect(drawing.subkind).toBe("inline-image");
    expect(drawing.relId).toMatch(/^rId\d+/);

    // Relationship landed in the H/F part's own rels file, NOT the
    // body's.
    const hfRels = snap.root.relationships.get(HEADER_PART) ?? [];
    expect(hfRels.some((r) => r.id === drawing.relId)).toBe(true);
    const bodyRels = snap.root.relationships.get("word/document.xml") ?? [];
    expect(bodyRels.some((r) => r.id === drawing.relId)).toBe(false);

    // Media part minted, dirty flags set.
    expect(snap.root.media.size).toBe(1);
    expect(snap.dirty.headersAndFooters.has(HEADER_PART)).toBe(true);
  });

  it("appends an image into an existing header paragraph at paragraphIndex", async () => {
    const agent = await makeAgent();
    await agent.applyCommand({
      type: "docx:insert-header-footer-image",
      payload: {
        partPath: HEADER_PART,
        paragraphIndex: 0,
        data: PNG_1x1,
        mimeType: "image/png",
        width: 32,
        height: 32,
      },
      source: "human",
    });
    const snap = agent.getSnapshot();
    const header = snap.root.headersAndFooters.find((p) => p.partPath === HEADER_PART);
    // Body length unchanged — image went into the existing paragraph.
    expect(header!.body.length).toBe(1);
    const para = header!.body[0] as Paragraph;
    const drawingRun = para.children.find(
      (c) => c.kind === "run" && c.children.some((rc) => rc.kind === "drawing")
    );
    expect(drawingRun).toBeTruthy();
  });

  it("rejects an unsupported MIME type with invalid-payload", async () => {
    const agent = await makeAgent();
    const m = await agent.applyCommand({
      type: "docx:insert-header-footer-image",
      payload: {
        partPath: HEADER_PART,
        data: PNG_1x1,
        mimeType: "application/pdf",
        width: 64,
        height: 64,
      },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-payload");
  });

  it("survives parse → insert → export → parse with the relationship intact", async () => {
    const agent = await makeAgent();
    await agent.applyCommand({
      type: "docx:insert-header-footer-image",
      payload: {
        partPath: HEADER_PART,
        data: PNG_1x1,
        mimeType: "image/png",
        width: 48,
        height: 48,
      },
      source: "human",
    });
    const buf = await agent.exportFile();
    const reparsed = await parseDocx(buf);
    const header = reparsed.root.headersAndFooters.find((p) => p.partPath === HEADER_PART);
    expect(header).toBeTruthy();
    // The header still has at least one paragraph carrying the
    // drawing, plus the original "Header" paragraph (or merged in).
    const hasDrawing = header!.body.some(
      (b) =>
        b.kind === "paragraph" &&
        b.children.some((c) => c.kind === "run" && c.children.some((rc) => rc.kind === "drawing"))
    );
    expect(hasDrawing).toBe(true);
    // Media part survived.
    expect(reparsed.root.media.size).toBeGreaterThanOrEqual(1);
  });
});
