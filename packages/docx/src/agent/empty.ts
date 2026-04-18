import JSZip from "jszip";

/**
 * Build the byte-buffer for a minimal but spec-valid blank DOCX:
 * one empty paragraph inside one section, no header/footer, no styles
 * part. The shape mirrors what `docx insert-text` etc. expect on an
 * "empty" document.
 *
 * Used by `DocxAgent.empty()` and `oa docx create --out`.
 */
export async function buildBlankDocxBuffer(): Promise<Uint8Array> {
  const z = new JSZip();
  z.file("[Content_Types].xml", CONTENT_TYPES_XML);
  z.file("_rels/.rels", PACKAGE_RELS_XML);
  z.file("word/_rels/document.xml.rels", DOC_RELS_XML);
  z.file("word/document.xml", DOCUMENT_XML);
  return (await z.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  })) as Uint8Array;
}

const CONTENT_TYPES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const PACKAGE_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOC_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`;

const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p/>
    <w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>
  </w:body>
</w:document>`;
