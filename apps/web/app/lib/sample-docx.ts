/**
 * Tiny synthetic DOCX produced entirely in-browser as a fallback for the
 * "I don't have a file to upload" demo flow. Uses JSZip via @officeai/docx
 * which already depends on it transitively through @officeai/core.
 *
 * Keeps the editor demo runnable with no network, no fixtures, no FS.
 */
import JSZip from "jszip";

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const PKG_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOC_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;

const DOC_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t xml:space="preserve">Welcome to officeAI</w:t></w:r></w:p>
    <w:p><w:r><w:t xml:space="preserve">This is a tiny synthetic .docx generated in your browser. It serves as a stand-in for "open file" so you can see the editor working without any upload.</w:t></w:r></w:p>
    <w:p><w:r><w:t xml:space="preserve">Type, format, or comment — every change goes through the headless agent's command bus, the same path an AI agent would use.</w:t></w:r></w:p>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
  </w:body>
</w:document>`;

export async function buildSampleDocx(): Promise<ArrayBuffer> {
  const z = new JSZip();
  z.file("[Content_Types].xml", CONTENT_TYPES);
  z.file("_rels/.rels", PKG_RELS);
  z.file("word/_rels/document.xml.rels", DOC_RELS);
  z.file("word/document.xml", DOC_XML);
  return z.generateAsync({ type: "arraybuffer" });
}
