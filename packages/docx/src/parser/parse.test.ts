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

  it("preserves unknown elements as opaque blocks", async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:sdt><w:sdtPr><w:tag w:val="custom"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>X</w:t></w:r></w:p></w:sdtContent></w:sdt>
</w:body></w:document>`;
    const buf = await makeSyntheticDocx({ documentXml: xml });
    const snap = await parseDocx(buf, { idMinter: deterministicIdMinter() });
    const sdt = snap.root.body[0];
    expect(sdt.kind).toBe("opaque-block");
    if (sdt.kind !== "opaque-block") throw new Error();
    expect(sdt.raw.tag).toBe("w:sdt");
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
    const sdt = snap.root.body[0];
    expect(sdt.kind).toBe("opaque-block");
    if (sdt.kind !== "opaque-block") throw new Error();
    expect(sdt.raw.tag).toBe("w:sdt");
    expect(sdt.children).toBeDefined();
    expect(sdt.children).toHaveLength(2);
    const c0 = sdt.children![0];
    if (c0.kind !== "paragraph") throw new Error();
    expect(c0.properties.styleId).toBe("Heading1");
    expect(c0.children[0]).toMatchObject({ kind: "run" });
    const c1 = sdt.children![1];
    if (c1.kind !== "paragraph") throw new Error();
    const r = c1.children[0];
    if (r.kind !== "run") throw new Error();
    expect(r.children[0]).toMatchObject({ kind: "text", text: "Entry one" });
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

  it("preserves SDT carrier raw subtree for byte-identical re-emission", async () => {
    const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>
  <w:sdt><w:sdtPr><w:tag w:val="custom"/></w:sdtPr><w:sdtContent><w:p><w:r><w:t>X</w:t></w:r></w:p></w:sdtContent></w:sdt>
</w:body></w:document>`;
    const buf = await makeSyntheticDocx({ documentXml: xml });
    const snap = await parseDocx(buf, { idMinter: deterministicIdMinter() });
    const sdt = snap.root.body[0];
    if (sdt.kind !== "opaque-block") throw new Error();
    // The wrapper's raw subtree retains both <w:sdtPr> and <w:sdtContent>
    // with the original paragraph inside, regardless of the parsed children.
    expect(sdt.raw.subtree.length).toBeGreaterThan(0);
    expect(sdt.subtreeDirty).toBeFalsy();
  });

  it("throws DocxParseError on missing main part", async () => {
    const JSZip = (await import("jszip")).default;
    const z = new JSZip();
    z.file("foo.txt", "not a docx");
    const buf = await z.generateAsync({ type: "arraybuffer" });
    await expect(parseDocx(buf)).rejects.toMatchObject({ name: "DocxParseError", code: "missing-main-part" });
  });
});
