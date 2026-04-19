import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { DocxAgent } from "../agent/agent.js";
import { parseDocx } from "../parser/parse.js";
import { serializeDocx } from "../serializer/serialize.js";
import type { BlockNode, InlineImageDrawing, Paragraph } from "../model/types.js";
import { DEFAULT_DOC_ROOT_ATTRS, escapeXml, makeSyntheticDocx } from "../test-utils/synthetic.js";

const PNG_1x1 = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52, 0x00, 0x00,
  0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4, 0x89, 0x00, 0x00, 0x00,
  0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x62, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d,
  0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
]);

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

function findInlineImage(p: Paragraph): InlineImageDrawing | null {
  for (const inline of p.children) {
    if (inline.kind !== "run") continue;
    for (const c of inline.children) {
      if (c.kind === "drawing" && c.subkind === "inline-image") return c;
    }
  }
  return null;
}

async function seedImage(agent: DocxAgent): Promise<string> {
  await agent.applyCommand({
    type: "docx:insert-image",
    payload: {
      at: { paragraph: 0 },
      data: PNG_1x1,
      mimeType: "image/png",
      width: 96,
      height: 96,
      altText: "smoke",
    },
    source: "human",
  });
  const img = findInlineImage(paraOf(agent.getSnapshot().root.body[0]));
  if (!img) throw new Error("seed: failed to insert image");
  return img.id;
}

describe("docx:set-image-properties (B6)", () => {
  it("updates width/height in EMU based on pixel inputs", async () => {
    const agent = await loadAgent(["x"]);
    const id = await seedImage(agent);
    const m = await agent.applyCommand({
      type: "docx:set-image-properties",
      payload: { imageId: id, widthPx: 200, heightPx: 150 },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const img = findInlineImage(paraOf(agent.getSnapshot().root.body[0]));
    expect(img?.cx).toBe(200 * 9525);
    expect(img?.cy).toBe(150 * 9525);
  });

  it("updates alt text and clears it when null is passed", async () => {
    const agent = await loadAgent(["x"]);
    const id = await seedImage(agent);
    await agent.applyCommand({
      type: "docx:set-image-properties",
      payload: { imageId: id, altText: "diagram of an octopus" },
      source: "human",
    });
    expect(findInlineImage(paraOf(agent.getSnapshot().root.body[0]))?.descr).toBe("diagram of an octopus");
    await agent.applyCommand({
      type: "docx:set-image-properties",
      payload: { imageId: id, altText: null },
      source: "human",
    });
    expect(findInlineImage(paraOf(agent.getSnapshot().root.body[0]))?.descr).toBeUndefined();
  });

  it("drops the cached `raw` so the serializer regenerates the drawing subtree", async () => {
    const agent = await loadAgent(["x"]);
    const id = await seedImage(agent);
    await agent.applyCommand({
      type: "docx:set-image-properties",
      payload: { imageId: id, widthPx: 64, heightPx: 64 },
      source: "human",
    });
    const img = findInlineImage(paraOf(agent.getSnapshot().root.body[0]));
    expect(img?.raw).toBeUndefined();
  });

  it("rejects when the image id is unknown", async () => {
    const agent = await loadAgent(["x"]);
    await seedImage(agent);
    const m = await agent.applyCommand({
      type: "docx:set-image-properties",
      payload: { imageId: "non-existent", widthPx: 10 },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-target");
  });

  it("rejects empty payloads", async () => {
    const agent = await loadAgent(["x"]);
    const id = await seedImage(agent);
    const m = await agent.applyCommand({
      type: "docx:set-image-properties",
      payload: { imageId: id },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("invalid-payload");
  });

  it("survives a save → reparse round-trip with the new dimensions intact", async () => {
    const agent = await loadAgent(["x"]);
    const id = await seedImage(agent);
    await agent.applyCommand({
      type: "docx:set-image-properties",
      payload: { imageId: id, widthPx: 48, heightPx: 48, altText: "round-trip" },
      source: "human",
    });
    const out = await serializeDocx(agent.getSnapshot());
    const reparsed = await parseDocx(out, { idMinter: deterministicIdMinter("z") });
    const img = findInlineImage(paraOf(reparsed.root.body[0]));
    expect(img).toBeTruthy();
    expect(img?.cx).toBe(48 * 9525);
    expect(img?.cy).toBe(48 * 9525);
    expect(img?.descr).toBe("round-trip");
  });
});
