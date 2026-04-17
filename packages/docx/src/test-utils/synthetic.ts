import JSZip from "jszip";

/**
 * Produce a minimal but valid-enough DOCX zip from XML strings, for tests.
 * The shape matches what Word emits (modulo features we don't model).
 */
export interface SyntheticDocxOptions {
  documentXml: string;
  /** Extra parts beyond the minimum (relationships, content-types, document.xml). */
  extra?: Record<string, string | Uint8Array>;
}

export async function makeSyntheticDocx(opts: SyntheticDocxOptions): Promise<ArrayBuffer> {
  const z = new JSZip();
  z.file("[Content_Types].xml", DEFAULT_CONTENT_TYPES);
  z.file("_rels/.rels", DEFAULT_PACKAGE_RELS);
  z.file("word/_rels/document.xml.rels", DEFAULT_DOC_RELS);
  z.file("word/document.xml", opts.documentXml);
  if (opts.extra) {
    for (const [path, content] of Object.entries(opts.extra)) {
      z.file(path, content);
    }
  }
  return z.generateAsync({
    type: "arraybuffer",
    compression: "DEFLATE",
    compressionOptions: { level: 6 },
  });
}

export const DEFAULT_DOC_ROOT_ATTRS = `
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
`
  .replace(/\s+/g, " ")
  .trim();

export function plainDocxXml(paragraphs: ReadonlyArray<{ text: string; styleId?: string }>): string {
  const ps = paragraphs
    .map((p) => {
      const pPr = p.styleId ? `<w:pPr><w:pStyle w:val="${escapeXml(p.styleId)}"/></w:pPr>` : "";
      return `<w:p>${pPr}<w:r><w:t xml:space="preserve">${escapeXml(p.text)}</w:t></w:r></w:p>`;
    })
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document ${DEFAULT_DOC_ROOT_ATTRS}><w:body>${ps}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`;
}

export function escapeXml(s: string): string {
  return s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

const DEFAULT_CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const DEFAULT_PACKAGE_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DEFAULT_DOC_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`;
