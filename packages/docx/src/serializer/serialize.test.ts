import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import { ooxml, sha256Hex, deterministicIdMinter } from "@officeai/core";
import { describe, expect, it } from "vitest";
import { parseDocx } from "../parser/parse.js";
import { makeSyntheticDocx, plainDocxXml } from "../test-utils/synthetic.js";
import { serializeDocx } from "./serialize.js";

const REPO_ROOT = resolve(__dirname, "../../../..");

function loadFixture(relPath: string): ArrayBuffer {
  const buf = readFileSync(resolve(REPO_ROOT, relPath));
  return buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength) as ArrayBuffer;
}

describe("serializeDocx", () => {
  it("re-emits unchanged container byte-identically (untouched parts)", async () => {
    const buf = await makeSyntheticDocx({
      documentXml: plainDocxXml([{ text: "Hello, world." }, { text: "Second paragraph." }]),
    });
    const snap = await parseDocx(buf, { idMinter: deterministicIdMinter() });
    const out = await serializeDocx(snap);
    const reloaded = await ooxml.OoxmlContainer.load(out);
    for (const path of snap.container.parts.keys()) {
      const beforeHash = sha256Hex(snap.container.readBytes(path));
      const afterHash = sha256Hex(reloaded.readBytes(path));
      expect(afterHash, `part ${path} should be byte-identical`).toBe(beforeHash);
    }
  });

  it("re-parses to a structurally-equivalent body after a no-op pass", async () => {
    const buf = await makeSyntheticDocx({
      documentXml: plainDocxXml([{ text: "Title", styleId: "Title" }, { text: "Body content here." }]),
    });
    const snap = await parseDocx(buf, { idMinter: deterministicIdMinter("a") });
    const out = await serializeDocx(snap);
    const snap2 = await parseDocx(out, { idMinter: deterministicIdMinter("b") });
    expect(snap2.root.body.length).toBe(snap.root.body.length);
    const p0 = snap.root.body[0];
    const q0 = snap2.root.body[0];
    if (p0.kind !== "paragraph" || q0.kind !== "paragraph") throw new Error();
    expect(q0.properties.styleId).toBe(p0.properties.styleId);
  });

  it("re-emits a lifted <w:sdt> carrier (wrapper-marker pair) when body is forced dirty", async () => {
    // Phase B of docx-fidelity-overhaul: body-level SDT carriers are
    // lifted into a `wrapper-marker(begin) ... wrapper-marker(end)`
    // pair around their inner blocks. On a dirty-body round-trip the
    // serializer must rebuild the carrier envelope from `wrapperRaw`
    // and splice the freshly serialized inner blocks back into the
    // `<w:sdtContent>` slot.
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:sdt>
    <w:sdtPr><w:alias w:val="TOC"/></w:sdtPr>
    <w:sdtContent><w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Inhaltsverzeichnis</w:t></w:r></w:p></w:sdtContent>
  </w:sdt>
  <w:p><w:r><w:t>Body</w:t></w:r></w:p>
</w:body></w:document>`;
    const buf = await makeSyntheticDocx({ documentXml: xml });
    const snap = await parseDocx(buf, { idMinter: deterministicIdMinter() });
    const dirtied = { ...snap, dirty: { ...snap.dirty, body: true } };
    const out = await serializeDocx(dirtied);
    const reparsed = await parseDocx(out, { idMinter: deterministicIdMinter("rt") });
    // Re-parse must lift the carrier the same way again.
    expect(reparsed.root.body).toHaveLength(4);
    const begin = reparsed.root.body[0];
    if (begin.kind !== "wrapper-marker") throw new Error();
    expect(begin.side).toBe("begin");
    expect(begin.wrapperRaw.tag).toBe("w:sdt");
    const inner = reparsed.root.body[1];
    if (inner.kind !== "paragraph") throw new Error();
    expect(inner.properties.styleId).toBe("Heading1");
    const end = reparsed.root.body[2];
    if (end.kind !== "wrapper-marker") throw new Error();
    expect(end.side).toBe("end");
    expect(end.wrapperId).toBe(begin.wrapperId);
    const trailing = reparsed.root.body[3];
    if (trailing.kind !== "paragraph") throw new Error();
  });

  it("round-trips the masterthesis TOC fixture (SDT carrier preserved)", async () => {
    const buf = loadFixture("fixtures/docx/real-world/07-toc-sdt.docx");
    const snap = await parseDocx(buf, { idMinter: deterministicIdMinter("toc") });
    // The fixture contains at least one TOC SDT carrier — Phase B
    // lifts it so we look for matched `wrapper-marker` pairs whose
    // captured `wrapperRaw.tag` is `w:sdt`.
    const begins = snap.root.body.filter(
      (b) => b.kind === "wrapper-marker" && b.side === "begin" && b.wrapperRaw.tag === "w:sdt"
    );
    expect(begins.length).toBeGreaterThanOrEqual(1);
    for (const begin of begins) {
      if (begin.kind !== "wrapper-marker") throw new Error();
      const endIdx = snap.root.body.findIndex(
        (b) => b.kind === "wrapper-marker" && b.side === "end" && b.wrapperId === begin.wrapperId
      );
      expect(endIdx).toBeGreaterThan(0);
    }
    // Pure round-trip with no dirty flags must reproduce the original
    // bytes byte-for-byte (the lifted children are render-only here;
    // the cached `word/document.xml` is preserved verbatim).
    const out = await serializeDocx(snap);
    const reloaded = await ooxml.OoxmlContainer.load(out);
    for (const path of snap.container.parts.keys()) {
      const beforeHash = sha256Hex(snap.container.readBytes(path));
      const afterHash = sha256Hex(reloaded.readBytes(path));
      expect(afterHash, `part ${path} byte-identical`).toBe(beforeHash);
    }
  });

  it("forces serialization when body is dirty and remains valid", async () => {
    const buf = await makeSyntheticDocx({
      documentXml: plainDocxXml([{ text: "before" }]),
    });
    const snap = await parseDocx(buf, { idMinter: deterministicIdMinter() });
    const dirtied = { ...snap, dirty: { ...snap.dirty, body: true } };
    const out = await serializeDocx(dirtied);
    const reparsed = await parseDocx(out, { idMinter: deterministicIdMinter("c") });
    const p0 = reparsed.root.body[0];
    if (p0.kind !== "paragraph") throw new Error();
    const r0 = p0.children[0];
    if (r0.kind !== "run") throw new Error();
    const t = r0.children[0];
    if (t.kind !== "text") throw new Error();
    expect(t.text).toBe("before");
  });
});
