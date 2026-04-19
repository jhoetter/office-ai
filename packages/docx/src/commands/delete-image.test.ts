import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { DocxAgent } from "../agent/agent.js";
import type { BlockNode, InlineImageDrawing, Paragraph, Run } from "../model/types.js";
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

describe("docx:delete-image (B6)", () => {
  it("removes the image leaf and leaves the paragraph + sibling text intact", async () => {
    const agent = await loadAgent(["before"]);
    await agent.applyCommand({
      type: "docx:insert-image",
      payload: {
        at: { paragraph: 0, run: 0, offset: 6 },
        data: PNG_1x1,
        mimeType: "image/png",
        width: 32,
        height: 32,
      },
      source: "human",
    });
    const seeded = paraOf(agent.getSnapshot().root.body[0]);
    const img = findInlineImage(seeded);
    expect(img).toBeTruthy();
    if (!img) return;

    const m = await agent.applyCommand({
      type: "docx:delete-image",
      payload: { imageId: img.id },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const after = paraOf(agent.getSnapshot().root.body[0]);
    expect(findInlineImage(after)).toBeNull();
    const text = after.children
      .filter((c): c is Run => c.kind === "run")
      .flatMap((r) => r.children)
      .filter((c) => c.kind === "text")
      .map((c) => (c.kind === "text" ? c.text : ""))
      .join("");
    expect(text).toBe("before");
  });

  it("rejects when no image with the given id exists", async () => {
    const agent = await loadAgent(["x"]);
    const m = await agent.applyCommand({
      type: "docx:delete-image",
      payload: { imageId: "nope" },
      source: "human",
    });
    expect(m.status).toBe("rejected");
    expect(m.rejection?.code).toBe("unknown-target");
  });
});
