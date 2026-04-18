import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { parseDocx } from "./parse.js";
import { makeSyntheticDocx, plainDocxXml } from "../test-utils/synthetic.js";

describe("parseDocx", () => {
  it("parses a minimal plain-text document", async () => {
    const buf = await makeSyntheticDocx({
      documentXml: plainDocxXml([{ text: "Hello, world." }, { text: "Second paragraph." }]),
    });
    const snap = await parseDocx(buf, { idMinter: deterministicIdMinter() });
    expect(snap.format).toBe("docx");
    expect(snap.revision).toBe(0);
    expect(snap.root.body).toHaveLength(3); // 2 paragraphs + sectPr
    const p0 = snap.root.body[0];
    expect(p0.kind).toBe("paragraph");
    if (p0.kind !== "paragraph") throw new Error("type narrowing");
    expect(p0.children).toHaveLength(1);
    const r0 = p0.children[0];
    expect(r0.kind).toBe("run");
    if (r0.kind !== "run") throw new Error("type narrowing");
    expect(r0.children[0]).toMatchObject({ kind: "text", text: "Hello, world." });
  });

  it("preserves paragraph styleId", async () => {
    const buf = await makeSyntheticDocx({
      documentXml: plainDocxXml([{ text: "Title", styleId: "Title" }]),
    });
    const snap = await parseDocx(buf, { idMinter: deterministicIdMinter() });
    const p = snap.root.body[0];
    if (p.kind !== "paragraph") throw new Error("type narrowing");
    expect(p.properties.styleId).toBe("Title");
  });

  it("parses run formatting (bold, italic, underline, color, size)", async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:p>
    <w:r>
      <w:rPr>
        <w:b/>
        <w:i/>
        <w:u w:val="single"/>
        <w:color w:val="FF0000"/>
        <w:sz w:val="28"/>
      </w:rPr>
      <w:t>Styled</w:t>
    </w:r>
  </w:p>
</w:body></w:document>`;
    const buf = await makeSyntheticDocx({ documentXml: xml });
    const snap = await parseDocx(buf, { idMinter: deterministicIdMinter() });
    const p = snap.root.body[0];
    if (p.kind !== "paragraph") throw new Error();
    const r = p.children[0];
    if (r.kind !== "run") throw new Error();
    expect(r.properties).toMatchObject({
      bold: true,
      italic: true,
      underline: true,
      color: "FF0000",
      fontSize: 28,
    });
  });

  it("lifts a body-level <w:sdt> carrier into wrapper-marker brackets + inner blocks", async () => {
    // Phase B of docx-fidelity-overhaul: body-level wrappers
    // (`<w:sdt>`, `mc:AlternateContent`, `<w:fldSimple>`, …) that
    // contain typed body blocks no longer survive as opaque-block
    // atoms. The parser splits them into
    //   `[ wrapper-marker(begin), ...inner blocks, wrapper-marker(end) ]`
    // so the page chunker can flow the inner paragraphs as regular
    // body content and the heading next to a TOC SDT no longer ends
    // up orphaned on its own page.
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:sdt><w:sdtPr><w:tag w:val="custom"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>X</w:t></w:r></w:p></w:sdtContent></w:sdt>
</w:body></w:document>`;
    const buf = await makeSyntheticDocx({ documentXml: xml });
    const snap = await parseDocx(buf, { idMinter: deterministicIdMinter() });
    const begin = snap.root.body[0];
    if (begin.kind !== "wrapper-marker") throw new Error();
    expect(begin.side).toBe("begin");
    expect(begin.wrapperRaw.tag).toBe("w:sdt");
    const inner = snap.root.body[1];
    if (inner.kind !== "paragraph") throw new Error();
    const end = snap.root.body[2];
    if (end.kind !== "wrapper-marker") throw new Error();
    expect(end.side).toBe("end");
    expect(end.wrapperId).toBe(begin.wrapperId);
  });

  it("unwraps a <w:sdt> wrapper's <w:sdtContent> into typed paragraph children", async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:sdt>
    <w:sdtPr><w:alias w:val="TOC"/></w:sdtPr>
    <w:sdtContent>
      <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Table of Contents</w:t></w:r></w:p>
      <w:p><w:r><w:t>Entry one</w:t></w:r></w:p>
    </w:sdtContent>
  </w:sdt>
</w:body></w:document>`;
    const buf = await makeSyntheticDocx({ documentXml: xml });
    const snap = await parseDocx(buf, { idMinter: deterministicIdMinter() });
    expect(snap.root.body).toHaveLength(4);
    const begin = snap.root.body[0];
    if (begin.kind !== "wrapper-marker") throw new Error();
    expect(begin.side).toBe("begin");
    expect(begin.wrapperRaw.tag).toBe("w:sdt");
    const c0 = snap.root.body[1];
    if (c0.kind !== "paragraph") throw new Error();
    expect(c0.properties.styleId).toBe("Heading1");
    expect(c0.children[0]).toMatchObject({ kind: "run" });
    const c1 = snap.root.body[2];
    if (c1.kind !== "paragraph") throw new Error();
    const r = c1.children[0];
    if (r.kind !== "run") throw new Error();
    expect(r.children[0]).toMatchObject({ kind: "text", text: "Entry one" });
    const end = snap.root.body[3];
    if (end.kind !== "wrapper-marker") throw new Error();
    expect(end.side).toBe("end");
    expect(end.wrapperId).toBe(begin.wrapperId);
  });

  it("unwraps a <w:fldSimple> wrapper into typed inline children", async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:p>
    <w:r><w:t xml:space="preserve">Page </w:t></w:r>
    <w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>
  </w:p>
</w:body></w:document>`;
    const buf = await makeSyntheticDocx({ documentXml: xml });
    const snap = await parseDocx(buf, { idMinter: deterministicIdMinter() });
    const p = snap.root.body[0];
    if (p.kind !== "paragraph") throw new Error();
    expect(p.children).toHaveLength(2);
    const op = p.children[1];
    expect(op.kind).toBe("opaque-inline");
    if (op.kind !== "opaque-inline") throw new Error();
    expect(op.raw.tag).toBe("w:fldSimple");
    expect(op.children).toBeDefined();
    expect(op.children).toHaveLength(1);
    const inner = op.children![0];
    if (inner.kind !== "run") throw new Error();
    expect(inner.children[0]).toMatchObject({ kind: "text", text: "1" });
  });

  it("preserves SDT carrier raw subtree on the wrapper-marker for byte-identical re-emission", async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:sdt><w:sdtPr><w:tag w:val="custom"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>X</w:t></w:r></w:p></w:sdtContent></w:sdt>
</w:body></w:document>`;
    const buf = await makeSyntheticDocx({ documentXml: xml });
    const snap = await parseDocx(buf, { idMinter: deterministicIdMinter() });
    const begin = snap.root.body[0];
    if (begin.kind !== "wrapper-marker") throw new Error();
    // The wrapper's raw subtree retains both <w:sdtPr> and <w:sdtContent>
    // with the original paragraph inside, so the serializer can rebuild
    // the carrier envelope on a body-dirty round-trip.
    expect(begin.wrapperRaw.subtree.length).toBeGreaterThan(0);
    expect(begin.wrapperRaw.tag).toBe("w:sdt");
  });

  it("throws DocxParseError on missing main part", async () => {
    const JSZip = (await import("jszip")).default;
    const z = new JSZip();
    z.file("foo.txt", "not a docx");
    const buf = await z.generateAsync({ type: "arraybuffer" });
    await expect(parseDocx(buf)).rejects.toMatchObject({ name: "DocxParseError", code: "missing-main-part" });
  });
});
