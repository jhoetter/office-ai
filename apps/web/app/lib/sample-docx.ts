/**
 * Tiny synthetic DOCX produced entirely in-browser as a fallback for the
 * "I don't have a file to upload" demo flow. Uses JSZip via @officeai/docx
 * which already depends on it transitively through @officeai/core.
 *
 * IMPORTANT — why this file ships a real `word/styles.xml`:
 *   The first paragraph references `<w:pStyle w:val="Heading1"/>`. Without a
 *   styles part that DEFINES `Heading1`, Word and LibreOffice silently fall
 *   back to default formatting and the heading renders as plain text in the
 *   exported .docx. Verified end-to-end with LibreOffice headless render.
 *
 * IMPORTANT — why this file ships a real `word/theme/theme1.xml`:
 *   The heading styles below reference `<w:rFonts w:asciiTheme="majorHAnsi"/>`,
 *   which Word 2024+ resolves through the document's font scheme to "Aptos
 *   Display". Without a theme part shipped alongside the doc, Word falls back
 *   to its built-in default theme (also Aptos), but the editor's style
 *   resolver had no way to project the theme ref to a literal typeface and
 *   would mis-report "Calibri" in the toolbar. We now ship the part so the
 *   round-trip is fully self-describing — the editor and Word agree on the
 *   rendered font ("Aptos Display") for headings and ("Aptos") for body text.
 */
import JSZip from "jszip";

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
  <Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>`;

const PKG_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOC_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`;

/**
 * Minimal theme part. Only the font scheme is interesting — the editor's
 * style cascade resolver consults `majorFont.latin` / `minorFont.latin`
 * to project `<w:rFonts w:asciiTheme="majorHAnsi"/>` references in the
 * styles below to the literal "Aptos Display" / "Aptos" typefaces Word
 * 2024+ uses by default. The color and format schemes are placeholders;
 * Word ignores their absence on this minimal doc.
 */
const THEME_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
      <a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="0E2841"/></a:dk2>
      <a:lt2><a:srgbClr val="E8E8E8"/></a:lt2>
      <a:accent1><a:srgbClr val="156082"/></a:accent1>
      <a:accent2><a:srgbClr val="E97132"/></a:accent2>
      <a:accent3><a:srgbClr val="196B24"/></a:accent3>
      <a:accent4><a:srgbClr val="0F9ED5"/></a:accent4>
      <a:accent5><a:srgbClr val="A02B93"/></a:accent5>
      <a:accent6><a:srgbClr val="4EA72E"/></a:accent6>
      <a:hlink><a:srgbClr val="467886"/></a:hlink>
      <a:folHlink><a:srgbClr val="96607D"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Aptos Display" panose="02110004020202020204"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Aptos" panose="02110004020202020204"/>
        <a:ea typeface=""/>
        <a:cs typeface=""/>
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
        <a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
        <a:effectStyle><a:effectLst/></a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
        <a:solidFill><a:schemeClr val="phClr"/></a:solidFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
</a:theme>`;

// DIN A4 — 210 × 297 mm — is the default for ~99% of locales outside
// the US, and is what Word emits by default in EU installs. Twips:
// 210 mm × (1440 twip/in ÷ 25.4 mm/in) = 11905.51 → 11906
// 297 mm × (1440 twip/in ÷ 25.4 mm/in) = 16837.79 → 16838
// Margins: 2.5 cm on every side = 1417 twips (Word's German A4 default).
// Header / footer offset: 1.25 cm = 708 twips.
const DOC_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t xml:space="preserve">Welcome to office-ai</w:t></w:r></w:p>
    <w:p><w:r><w:t xml:space="preserve">This is a tiny synthetic .docx generated in your browser. It serves as a stand-in for "open file" so you can see the editor working without any upload.</w:t></w:r></w:p>
    <w:p><w:r><w:t xml:space="preserve">Type, format, or comment — every change goes through the headless agent's command bus, the same path an AI agent would use.</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;

/**
 * Minimal but valid styles.xml covering Normal + Heading 1-3 + Title.
 * Modeled on what Word emits when you save a fresh document; trimmed to
 * the essentials Word/LibreOffice need to apply visible heading styling.
 */
const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
        <w:lang w:val="en-US"/>
      </w:rPr>
    </w:rPrDefault>
    <w:pPrDefault>
      <w:pPr>
        <w:spacing w:after="160" w:line="259" w:lineRule="auto"/>
      </w:pPr>
    </w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Title">
    <w:name w:val="Title"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:spacing w:before="240" w:after="60"/>
      <w:contextualSpacing/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/>
      <w:b/>
      <w:bCs/>
      <w:sz w:val="56"/>
      <w:szCs w:val="56"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      <w:spacing w:before="240" w:after="60"/>
      <w:outlineLvl w:val="0"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/>
      <w:b/>
      <w:bCs/>
      <w:sz w:val="36"/>
      <w:szCs w:val="36"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      <w:spacing w:before="200" w:after="60"/>
      <w:outlineLvl w:val="1"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/>
      <w:b/>
      <w:bCs/>
      <w:sz w:val="28"/>
      <w:szCs w:val="28"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:pPr>
      <w:keepNext/>
      <w:keepLines/>
      <w:spacing w:before="160" w:after="60"/>
      <w:outlineLvl w:val="2"/>
    </w:pPr>
    <w:rPr>
      <w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"/>
      <w:b/>
      <w:bCs/>
      <w:sz w:val="24"/>
      <w:szCs w:val="24"/>
    </w:rPr>
  </w:style>
</w:styles>`;

export async function buildSampleDocx(): Promise<ArrayBuffer> {
  const z = new JSZip();
  z.file("[Content_Types].xml", CONTENT_TYPES);
  z.file("_rels/.rels", PKG_RELS);
  z.file("word/_rels/document.xml.rels", DOC_RELS);
  z.file("word/document.xml", DOC_XML);
  z.file("word/styles.xml", STYLES_XML);
  z.file("word/theme/theme1.xml", THEME_XML);
  return z.generateAsync({ type: "arraybuffer" });
}

// Blank companion to `buildSampleDocx`. Same package shape — we keep
// styles + theme so headings/font scheme still resolve once the user
// types — but the body is a single empty paragraph instead of the
// welcome copy. Used when the user picks "New document" on the home
// page.
const BLANK_DOC_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p/>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1417" w:right="1417" w:bottom="1417" w:left="1417" w:header="708" w:footer="708" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;

export async function buildBlankDocx(): Promise<ArrayBuffer> {
  const z = new JSZip();
  z.file("[Content_Types].xml", CONTENT_TYPES);
  z.file("_rels/.rels", PKG_RELS);
  z.file("word/_rels/document.xml.rels", DOC_RELS);
  z.file("word/document.xml", BLANK_DOC_XML);
  z.file("word/styles.xml", STYLES_XML);
  z.file("word/theme/theme1.xml", THEME_XML);
  return z.generateAsync({ type: "arraybuffer" });
}
