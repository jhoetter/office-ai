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

  it("re-emits an unwrapped <w:sdt> carrier byte-identically when subtreeDirty is false", async () => {
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
    // Force body re-emission so the cached document.xml is rebuilt from the
    // typed model. Even with `subtreeDirty=false` the SDT should round-trip
    // through the typed serializer to byte-equivalent bytes (the typed
    // `children` are NOT the source of truth — `raw` is).
    const dirtied = { ...snap, dirty: { ...snap.dirty, body: true } };
    const out = await serializeDocx(dirtied);
    const reparsed = await parseDocx(out, { idMinter: deterministicIdMinter("rt") });
    const sdt = reparsed.root.body[0];
    expect(sdt.kind).toBe("opaque-block");
    if (sdt.kind !== "opaque-block") throw new Error();
    expect(sdt.raw.tag).toBe("w:sdt");
    // Children must still be present after a round-trip (parser unwraps
    // them again).
    expect(sdt.children).toBeDefined();
    expect(sdt.children).toHaveLength(1);
    const c0 = sdt.children![0];
    if (c0.kind !== "paragraph") throw new Error();
    expect(c0.properties.styleId).toBe("Heading1");
    expect(reparsed.root.body[1]?.kind).toBe("paragraph");
  });

  it("round-trips the masterthesis TOC fixture (SDT carrier preserved)", async () => {
    const buf = loadFixture("fixtures/docx/real-world/07-toc-sdt.docx");
    const snap = await parseDocx(buf, { idMinter: deterministicIdMinter("toc") });
    // The fixture contains two SDT carriers (table-of-contents wrappers).
    const sdts = snap.root.body.filter((b) => b.kind === "opaque-block" && b.raw.tag === "w:sdt");
    expect(sdts.length).toBeGreaterThanOrEqual(1);
    for (const sdt of sdts) {
      if (sdt.kind !== "opaque-block") throw new Error();
      // Each TOC SDT must surface its inner paragraphs as typed children
      // so the renderer can show the entries instead of an opaque chip.
      expect(sdt.children).toBeDefined();
      expect(sdt.children!.length).toBeGreaterThan(0);
      expect(sdt.children!.some((c) => c.kind === "paragraph")).toBe(true);
    }
    // Pure round-trip with no dirty flags must reproduce the original bytes
    // byte-for-byte (typed `children` are render-only when subtreeDirty is
    // false).
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
