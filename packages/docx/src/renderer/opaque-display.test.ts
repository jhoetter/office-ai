import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { DocxAgent } from "../agent/agent.js";
import { docToPM } from "./doc-to-pm.js";
import { docxSchema } from "./schema.js";
import { DEFAULT_DOC_ROOT_ATTRS, escapeXml, makeSyntheticDocx } from "../test-utils/synthetic.js";
import { serializeDocx } from "../serializer/serialize.js";

/**
 * Verify the renderer's opaque-display classification surfaces real-world
 * structural elements (bookmarks, fields, SDT content controls) the way a
 * Word user expects:
 *
 *   - metadata carriers render to nothing (no "[opaque]" clutter),
 *   - content-wrapper carriers surface their inner text,
 *   - placeholders fall back to the legacy "[<tag>]" chip,
 *
 * while the underlying model continues to round-trip the original bytes.
 */

interface BodyDocxOptions {
  bodyXml: string;
  extra?: Record<string, string | Uint8Array>;
}

async function makeBodyDocx(opts: BodyDocxOptions): Promise<ArrayBuffer> {
  const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS} xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
  <w:body>${opts.bodyXml}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body>
</w:document>`;
  return makeSyntheticDocx({ documentXml: xml, extra: opts.extra });
}

async function loadAgent(bodyXml: string): Promise<DocxAgent> {
  const buf = await makeBodyDocx({ bodyXml });
  return DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
}

function inlineNodes(pm: import("prosemirror-model").Node, paragraphIndex: number) {
  const para = pm.child(paragraphIndex);
  const out: import("prosemirror-model").Node[] = [];
  para.descendants((n) => {
    if (n.isText || n.isInline) out.push(n);
    return true;
  });
  return out;
}

describe("renderer opaque display", () => {
  it("renders bookmark markers as nothing (no [opaque] chip)", async () => {
    // `<w:bookmarkStart>` and `<w:bookmarkEnd>` flank a heading exactly the
    // way a real-world TOC anchor does. Their model representation is an
    // `OpaqueInline` carrier (round-trips byte-for-byte), but the renderer
    // must not surface them as visible "[opaque]" chips.
    const heading =
      "<w:p>" +
      '<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' +
      '<w:bookmarkStart w:id="0" w:name="_Toc1"/>' +
      '<w:r><w:t xml:space="preserve">Inhaltsverzeichnis</w:t></w:r>' +
      '<w:bookmarkEnd w:id="0"/>' +
      "</w:p>";
    const agent = await loadAgent(heading);
    const pm = docToPM(agent.getSnapshot());

    const inlines = inlineNodes(pm, 0);
    expect(inlines).toHaveLength(1);
    expect(inlines[0]!.type.name).toBe("text");
    expect(pm.child(0).textContent).toBe("Inhaltsverzeichnis");
    expect(JSON.stringify(pm.toJSON())).not.toContain("[opaque]");
  });

  it("renders fldChar / instrText / lastRenderedPageBreak as nothing", async () => {
    // The classic "PAGE" field consists of three opaque run-children
    // (`fldChar begin`, `instrText`, `fldChar end`) plus a typed text run
    // for the field result. The user should see only the result text.
    const para =
      "<w:p>" +
      '<w:r><w:t xml:space="preserve">Page </w:t></w:r>' +
      '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
      '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
      '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
      "<w:r><w:lastRenderedPageBreak/><w:t>1</w:t></w:r>" +
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
      "</w:p>";
    const agent = await loadAgent(para);
    const pm = docToPM(agent.getSnapshot());
    expect(pm.child(0).textContent).toBe("Page 1");
    expect(JSON.stringify(pm.toJSON())).not.toContain("[opaque]");
  });

  it("surfaces SDT content text via the unwrapped opaque_block_wrapper node", async () => {
    // A minimal SDT block wraps a paragraph that holds the user-visible
    // TOC text. Pre-P2.3, this rendered as a single "[w:sdt]" chip; with
    // SDT/TOC unwrapping the parser exposes the inner paragraphs and the
    // renderer produces a structured `opaque_block_wrapper` carrying them.
    const sdt =
      "<w:sdt>" +
      '<w:sdtPr><w:alias w:val="Inhaltsverzeichnis"/></w:sdtPr>' +
      "<w:sdtContent>" +
      '<w:p><w:r><w:t xml:space="preserve">Inhaltsverzeichnis</w:t></w:r></w:p>' +
      '<w:p><w:r><w:t xml:space="preserve">1 Einleitung ............... 1</w:t></w:r></w:p>' +
      "</w:sdtContent>" +
      "</w:sdt>";
    const agent = await loadAgent(sdt);
    const pm = docToPM(agent.getSnapshot());

    // body contributes the SDT block plus a trailing section-break block
    // (synthetic helper always appends `<w:sectPr>`).
    expect(pm.childCount).toBe(2);
    const block = pm.child(0);
    expect(block.type.name).toBe("opaque_block_wrapper");
    expect(block.attrs.tag).toBe("w:sdt");
    const json = JSON.stringify(pm.toJSON());
    expect(json).toContain("Inhaltsverzeichnis");
    expect(json).toContain("1 Einleitung");
    expect(json).not.toContain("[w:sdt]");
  });

  it("falls back to the [<tag>] placeholder for unclassified opaque blocks", async () => {
    const exotic = '<w:bizarreUnknown w:val="x"><w:foo/></w:bizarreUnknown>';
    const agent = await loadAgent(exotic);
    const pm = docToPM(agent.getSnapshot());
    const block = pm.child(0);
    expect(block.type.name).toBe("opaque_block");
    expect(block.attrs.tag).toBe("w:bizarreUnknown");
    expect(block.attrs.previewText).toBeNull();
    // toDOM round-trip uses the placeholder string when previewText is null.
    const spec = docxSchema.nodes.opaque_block.spec;
    if (typeof spec.toDOM === "function") {
      const dom = spec.toDOM(block) as [string, Record<string, string>, string];
      expect(dom[2]).toBe("[w:bizarreUnknown]");
    }
  });

  it("preserves the document bytes despite the renderer reclassification", async () => {
    // Display reclassification must NOT touch what the serializer emits.
    // We re-emit the same document and confirm the body XML round-trips
    // structurally (the snapshot is loaded back without errors and
    // exposes the same opaque carriers).
    const bookmarked =
      "<w:p>" +
      '<w:bookmarkStart w:id="0" w:name="_Toc1"/>' +
      '<w:r><w:t xml:space="preserve">Heading</w:t></w:r>' +
      '<w:bookmarkEnd w:id="0"/>' +
      "</w:p>";
    const agent = await loadAgent(bookmarked);
    const buf = await serializeDocx(agent.getSnapshot());
    const reloaded = await DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
    const para = reloaded.getSnapshot().root.body[0];
    if (!para || para.kind !== "paragraph") throw new Error("expected paragraph");
    const kinds = para.children.map((c) => c.kind);
    expect(kinds).toEqual(["opaque-inline", "run", "opaque-inline"]);
    expect((para.children[0] as { raw: { tag: string } }).raw.tag).toBe("w:bookmarkStart");
    expect((para.children[2] as { raw: { tag: string } }).raw.tag).toBe("w:bookmarkEnd");
  });

  it("does not introduce 'opaque' clutter into a multi-section paragraph paste path", async () => {
    // Smoke test: a paragraph containing text + bookmark + text + field +
    // text round-trips into a PM doc whose JSON serialisation contains no
    // "opaque" text artefact.
    const mixed =
      "<w:p>" +
      '<w:r><w:t xml:space="preserve">Before </w:t></w:r>' +
      '<w:bookmarkStart w:id="1" w:name="anchor"/>' +
      '<w:r><w:t xml:space="preserve">middle</w:t></w:r>' +
      '<w:bookmarkEnd w:id="1"/>' +
      '<w:r><w:t xml:space="preserve"> after</w:t></w:r>' +
      "</w:p>";
    void escapeXml; // re-export sanity
    const agent = await loadAgent(mixed);
    const pm = docToPM(agent.getSnapshot());
    expect(pm.child(0).textContent).toBe("Before middle after");
  });
});
