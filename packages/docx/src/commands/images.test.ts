import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { deterministicIdMinter, ooxml, sha256Hex } from "@officeai/core";
import { DocxAgent } from "../agent/agent.js";
import { parseDocx } from "../parser/parse.js";
import { serializeDocx } from "../serializer/serialize.js";
import type { BlockNode, InlineImageDrawing, Paragraph, Run } from "../model/types.js";
import { DEFAULT_DOC_ROOT_ATTRS, escapeXml, makeSyntheticDocx } from "../test-utils/synthetic.js";

const FIXTURE_PATH = resolve(__dirname, "../../../../fixtures/docx/real-world/05-inline-image.docx");

/**
 * Smallest valid PNG: 1×1 transparent pixel. Used as the body of every
 * synthetic image insertion so we exercise the full media-bytes
 * round-trip without depending on a fixture asset.
 */
const PNG_1x1 = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52, 0x00, 0x00,
  0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4, 0x89, 0x00, 0x00, 0x00,
  0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x62, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d,
  0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
]);

const PNG_1x1_RED = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52, 0x00, 0x00,
  0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xde, 0x00, 0x00, 0x00,
  0x0c, 0x49, 0x44, 0x41, 0x54, 0x08, 0x99, 0x63, 0xf8, 0xcf, 0xc0, 0x00, 0x00, 0x00, 0x03, 0x00, 0x01, 0x5b,
  0xa0, 0xfa, 0x6b, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
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

function findInlineImage(p: Paragraph): InlineImageDrawing | null {
  for (const inline of p.children) {
    if (inline.kind !== "run") continue;
    for (const c of inline.children) {
      if (c.kind === "drawing" && c.subkind === "inline-image") return c;
    }
  }
  return null;
}

function paraOf(block: BlockNode): Paragraph {
  if (block.kind !== "paragraph") throw new Error("expected paragraph");
  return block;
}

describe("docx images — parser, serializer, insert-image (P1.3 / W8)", () => {
  it("parses the real-world 05-inline-image fixture into a typed InlineImageDrawing", async () => {
    const buf = await readFile(FIXTURE_PATH);
    const snap = await parseDocx(buf, { idMinter: deterministicIdMinter() });
    // The image lives in the second paragraph ("Logo: <image> ...").
    const p1 = paraOf(snap.root.body[1]);
    const img = findInlineImage(p1);
    expect(img).toBeTruthy();
    if (!img) return;
    expect(img.subkind).toBe("inline-image");
    expect(img.relId).toBe("rId7");
    expect(img.cx).toBe(228600);
    expect(img.cy).toBe(228600);
    expect(img.docPrId).toBe(1);
    // Cached subtree is present so the byte-preservation fast path is armed.
    expect(img.raw).toBeTruthy();
  });

  it("populates `media` and `relationships` on a parsed snapshot containing an image", async () => {
    const buf = await readFile(FIXTURE_PATH);
    const snap = await parseDocx(buf, { idMinter: deterministicIdMinter() });
    const mediaPath = "word/media/27af7beffa5cca2fee08379789a1d11eafbd999d.png";
    const media = snap.root.media.get(mediaPath);
    expect(media).toBeTruthy();
    expect(media?.mimeType).toBe("image/png");
    expect(media?.digest).toMatch(/^[0-9a-f]{64}$/);
    const docRels = snap.root.relationships.get("word/document.xml") ?? [];
    const imageRels = docRels.filter((r) => r.type.endsWith("/relationships/image"));
    expect(imageRels).toHaveLength(1);
    expect(imageRels[0].id).toBe("rId7");
  });

  it("round-trips the inline-image fixture byte-identically through a no-touch save", async () => {
    const buf = await readFile(FIXTURE_PATH);
    const snap = await parseDocx(buf, { idMinter: deterministicIdMinter() });
    const out = await serializeDocx(snap);
    const reloaded = await ooxml.OoxmlContainer.load(out);
    for (const path of snap.container.parts.keys()) {
      const before = sha256Hex(snap.container.readBytes(path));
      const after = sha256Hex(reloaded.readBytes(path));
      expect(after, `part ${path} should be byte-identical`).toBe(before);
    }
  });

  it("inserts a brand-new image, mints rel + media, and dirty-flags the right parts", async () => {
    const agent = await loadAgent(["Hello"]);
    const before = agent.getSnapshot();
    const m = await agent.applyCommand({
      type: "docx:insert-image",
      payload: {
        at: { paragraph: 0 },
        data: PNG_1x1,
        mimeType: "image/png",
        width: 96,
        height: 96,
        altText: "smoke test",
      },
      source: "human",
    });
    expect(m.status).toBe("approved");
    const snap = agent.getSnapshot();
    expect(snap.dirty.body).toBe(true);
    expect(snap.dirty.media.size).toBe(1);
    expect(snap.dirty.relationships.has("word/document.xml")).toBe(true);
    expect(snap.dirty.contentTypes).toBe(true);

    const img = findInlineImage(paraOf(snap.root.body[0]));
    expect(img).toBeTruthy();
    if (!img) return;
    expect(img.cx).toBe(96 * 9525);
    expect(img.cy).toBe(96 * 9525);
    expect(img.descr).toBe("smoke test");
    // Diff has both a node-inserted and a part-added change.
    expect(m.diff.changes.map((c) => c.kind).sort()).toEqual(["node-inserted", "part-added"]);
    void before;
  });

  it("mints a fresh `word/media/imageN.png` part path that does not collide with existing media", async () => {
    const agent = await loadAgent(["First", "Second"]);
    await agent.applyCommand({
      type: "docx:insert-image",
      payload: {
        at: { paragraph: 0 },
        data: PNG_1x1,
        mimeType: "image/png",
        width: 32,
        height: 32,
      },
      source: "human",
    });
    await agent.applyCommand({
      type: "docx:insert-image",
      payload: {
        at: { paragraph: 1 },
        data: PNG_1x1_RED,
        mimeType: "image/png",
        width: 32,
        height: 32,
      },
      source: "human",
    });
    const snap = agent.getSnapshot();
    const paths = Array.from(snap.root.media.keys()).sort();
    expect(paths).toEqual(["word/media/image1.png", "word/media/image2.png"]);
  });

  it("de-duplicates by SHA-256: inserting identical bytes twice reuses one media part and one relationship", async () => {
    const agent = await loadAgent(["First", "Second"]);
    await agent.applyCommand({
      type: "docx:insert-image",
      payload: {
        at: { paragraph: 0 },
        data: PNG_1x1,
        mimeType: "image/png",
        width: 32,
        height: 32,
      },
      source: "human",
    });
    const m2 = await agent.applyCommand({
      type: "docx:insert-image",
      payload: {
        at: { paragraph: 1 },
        data: PNG_1x1,
        mimeType: "image/png",
        width: 32,
        height: 32,
      },
      source: "human",
    });
    expect(m2.status).toBe("approved");
    const snap = agent.getSnapshot();
    expect(snap.root.media.size).toBe(1);
    const docRels = snap.root.relationships.get("word/document.xml") ?? [];
    const imageRels = docRels.filter((r) => r.type.endsWith("/relationships/image"));
    expect(imageRels).toHaveLength(1);

    // Both images point at the same rId.
    const img0 = findInlineImage(paraOf(snap.root.body[0]));
    const img1 = findInlineImage(paraOf(snap.root.body[1]));
    expect(img0?.relId).toBe(imageRels[0].id);
    expect(img1?.relId).toBe(imageRels[0].id);
    // The second insertion did not add a new part, so its diff carries
    // only the node-inserted change.
    expect(m2.diff.changes.map((c) => c.kind)).toEqual(["node-inserted"]);
  });

  it("assigns unique <wp:docPr id> values to every inserted image", async () => {
    const agent = await loadAgent(["a", "b", "c"]);
    for (let i = 0; i < 3; i++) {
      await agent.applyCommand({
        type: "docx:insert-image",
        payload: {
          at: { paragraph: i },
          data: i % 2 === 0 ? PNG_1x1 : PNG_1x1_RED,
          mimeType: "image/png",
          width: 16,
          height: 16,
        },
        source: "human",
      });
    }
    const snap = agent.getSnapshot();
    const docPrIds: number[] = [];
    for (const block of snap.root.body) {
      if (block.kind !== "paragraph") continue;
      const img = findInlineImage(block);
      if (img) docPrIds.push(img.docPrId);
    }
    expect(new Set(docPrIds).size).toBe(docPrIds.length);
    expect(docPrIds.length).toBe(3);
  });

  it("rejects payloads with empty bytes, non-positive dimensions, or unsupported MIME", async () => {
    const agent = await loadAgent(["x"]);
    const empty = await agent.applyCommand({
      type: "docx:insert-image",
      payload: { at: { paragraph: 0 }, data: new Uint8Array(0), mimeType: "image/png", width: 1, height: 1 },
      source: "human",
    });
    expect(empty.status).toBe("rejected");
    expect(empty.rejection?.code).toBe("invalid-payload");
    const zeroW = await agent.applyCommand({
      type: "docx:insert-image",
      payload: { at: { paragraph: 0 }, data: PNG_1x1, mimeType: "image/png", width: 0, height: 1 },
      source: "human",
    });
    expect(zeroW.status).toBe("rejected");
    const badMime = await agent.applyCommand({
      type: "docx:insert-image",
      payload: { at: { paragraph: 0 }, data: PNG_1x1, mimeType: "application/pdf", width: 1, height: 1 },
      source: "human",
    });
    expect(badMime.status).toBe("rejected");
  });

  it("rejects insertion at an out-of-range paragraph or onto a non-paragraph block", async () => {
    const agent = await loadAgent(["only"]);
    const oob = await agent.applyCommand({
      type: "docx:insert-image",
      payload: { at: { paragraph: 99 }, data: PNG_1x1, mimeType: "image/png", width: 16, height: 16 },
      source: "human",
    });
    expect(oob.status).toBe("rejected");
    expect(oob.rejection?.code).toBe("invalid-position");
  });

  it("inserts the image as a fresh run mid-paragraph, splitting the targeted run at offset", async () => {
    const agent = await loadAgent(["abcdef"]);
    await agent.applyCommand({
      type: "docx:insert-image",
      payload: {
        at: { paragraph: 0, run: 0, offset: 3 },
        data: PNG_1x1,
        mimeType: "image/png",
        width: 16,
        height: 16,
      },
      source: "human",
    });
    const snap = agent.getSnapshot();
    const p = paraOf(snap.root.body[0]);
    // Expect three runs: "abc", <image>, "def".
    const runs = p.children.filter((c): c is Run => c.kind === "run");
    expect(runs.length).toBe(3);
    const text0 = runs[0].children.find((c) => c.kind === "text");
    const text2 = runs[2].children.find((c) => c.kind === "text");
    expect(text0 && text0.kind === "text" ? text0.text : "").toBe("abc");
    expect(text2 && text2.kind === "text" ? text2.text : "").toBe("def");
    const drawing = runs[1].children[0];
    expect(drawing.kind).toBe("drawing");
  });

  it("registers the new media MIME's <Default Extension> in [Content_Types].xml on save", async () => {
    const agent = await loadAgent(["x"]);
    await agent.applyCommand({
      type: "docx:insert-image",
      payload: {
        at: { paragraph: 0 },
        data: PNG_1x1,
        mimeType: "image/png",
        width: 16,
        height: 16,
      },
      source: "human",
    });
    const snap = agent.getSnapshot();
    const buf = await serializeDocx(snap);
    const reloaded = await ooxml.OoxmlContainer.load(buf);
    const ct = ooxml.ContentTypes.load(reloaded);
    expect(ct.hasDefault("png")).toBe(true);
  });

  it("emits the new image part, relationship, and a parseable <w:drawing> on save", async () => {
    const agent = await loadAgent(["x"]);
    await agent.applyCommand({
      type: "docx:insert-image",
      payload: {
        at: { paragraph: 0 },
        data: PNG_1x1,
        mimeType: "image/png",
        width: 24,
        height: 24,
      },
      source: "human",
    });
    const snap = agent.getSnapshot();
    const buf = await serializeDocx(snap);
    // Reparse and verify the typed leaf survived the round-trip.
    const snap2 = await parseDocx(buf, { idMinter: deterministicIdMinter("z") });
    const p = paraOf(snap2.root.body[0]);
    const img = findInlineImage(p);
    expect(img).toBeTruthy();
    if (!img) return;
    expect(img.cx).toBe(24 * 9525);
    expect(img.cy).toBe(24 * 9525);
    // Media + rel + content-type all wired up.
    expect(snap2.root.media.size).toBe(1);
    const docRels = snap2.root.relationships.get("word/document.xml") ?? [];
    expect(docRels.find((r) => r.id === img.relId)?.type).toMatch(/\/relationships\/image$/);
  });

  // Regression: inserting an image into a doc whose <w:document> root only
  // declares xmlns:w / xmlns:r used to emit <wp:inline> without a wp:
  // namespace declaration anywhere in scope, which makes Word reject the
  // file as corrupt ("namespace prefix wp on inline is not defined"). The
  // serializer now declares xmlns:wp locally on the <wp:inline> element.
  it("declares xmlns:wp on <wp:inline> so Word can parse the drawing", async () => {
    const agent = await loadAgent(["x"]);
    await agent.applyCommand({
      type: "docx:insert-image",
      payload: {
        at: { paragraph: 0 },
        data: PNG_1x1,
        mimeType: "image/png",
        width: 16,
        height: 16,
      },
      source: "human",
    });
    const snap = agent.getSnapshot();
    const buf = await serializeDocx(snap);
    const reloaded = await ooxml.OoxmlContainer.load(buf);
    const docXml = reloaded.readText("word/document.xml");
    expect(docXml).toMatch(
      /<wp:inline[^>]*xmlns:wp="http:\/\/schemas\.openxmlformats\.org\/drawingml\/2006\/wordprocessingDrawing"/
    );
  });

  it("does not dirty media/rels/contentTypes when inserting bytes that match an existing media digest", async () => {
    const agent = await loadAgent(["one", "two"]);
    await agent.applyCommand({
      type: "docx:insert-image",
      payload: {
        at: { paragraph: 0 },
        data: PNG_1x1,
        mimeType: "image/png",
        width: 16,
        height: 16,
      },
      source: "human",
    });
    const beforeMedia = agent.getSnapshot().dirty.media;
    const beforeRels = agent.getSnapshot().dirty.relationships;
    await agent.applyCommand({
      type: "docx:insert-image",
      payload: {
        at: { paragraph: 1 },
        data: PNG_1x1,
        mimeType: "image/png",
        width: 16,
        height: 16,
      },
      source: "human",
    });
    const after = agent.getSnapshot();
    // The dirty sets shouldn't grow on a de-dup hit (still contain just
    // the entries seeded by the first insertion).
    expect(after.dirty.media.size).toBe(beforeMedia.size);
    expect(after.dirty.relationships.size).toBe(beforeRels.size);
  });

  it("preserves byte-identity of the parsed inline-image leaf when a sibling paragraph is mutated", async () => {
    const buf = await readFile(FIXTURE_PATH);
    const agent = await DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    // Mutate paragraph 0 ("With logo") via a text insertion. This dirties
    // body but must NOT regenerate the cached inline-image bytes (the
    // <w:drawing> in paragraph 1 should round-trip via its raw cache).
    await agent.applyCommand({
      type: "docx:insert-text",
      payload: { at: { paragraph: 0 }, text: "X " },
      source: "human",
    });
    const snap = agent.getSnapshot();
    const out = await serializeDocx(snap);
    const snap2 = await parseDocx(out, { idMinter: deterministicIdMinter("y") });
    const img = findInlineImage(paraOf(snap2.root.body[1]));
    expect(img).toBeTruthy();
    if (!img) return;
    expect(img.relId).toBe("rId7");
    expect(img.cx).toBe(228600);
    expect(img.cy).toBe(228600);
    // The media bytes must be byte-identical (no part path changed).
    const mediaPath = "word/media/27af7beffa5cca2fee08379789a1d11eafbd999d.png";
    const before = sha256Hex(agent.getSnapshot().container.readBytes(mediaPath));
    const after = sha256Hex(snap2.container.readBytes(mediaPath));
    expect(after).toBe(before);
  });
});
