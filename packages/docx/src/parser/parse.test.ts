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

  it("throws DocxParseError on missing main part", async () => {
    const JSZip = (await import("jszip")).default;
    const z = new JSZip();
    z.file("foo.txt", "not a docx");
    const buf = await z.generateAsync({ type: "arraybuffer" });
    await expect(parseDocx(buf)).rejects.toMatchObject({ name: "DocxParseError", code: "missing-main-part" });
  });
});
