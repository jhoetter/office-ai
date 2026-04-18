/**
 * Tiny synthetic XLSX produced entirely in-browser as a fallback for the
 * "I don't have a file to upload" demo flow on /xlsx-editor. Mirrors
 * `sample-docx.ts`.
 *
 * Layout:
 *   A1=Name   B1=Score
 *   A2=Alex   B2=42
 *   A3=Sam    B3=37
 *   A4=Total  B4==SUM(B2:B3)
 *
 * Strings are encoded inline (`t="inlineStr"`) to skip a sharedStrings
 * part. The styles part is the SheetJS-default minimum so the parser's
 * style-table loader always has something to intern through.
 */
import JSZip from "jszip";

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;

const PKG_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;

const WORKBOOK_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`;

const WORKBOOK_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;

const SHEET_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:B4"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Name</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Score</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>Alex</t></is></c>
      <c r="B2"><v>42</v></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>Sam</t></is></c>
      <c r="B3"><v>37</v></c>
    </row>
    <row r="4">
      <c r="A4" t="inlineStr"><is><t>Total</t></is></c>
      <c r="B4"><f>SUM(B2:B3)</f><v>79</v></c>
    </row>
  </sheetData>
</worksheet>`;

const STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>
  <fills count="1"><fill><patternFill patternType="none"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;

export async function buildSampleXlsx(): Promise<ArrayBuffer> {
  const z = new JSZip();
  z.file("[Content_Types].xml", CONTENT_TYPES);
  z.file("_rels/.rels", PKG_RELS);
  z.file("xl/workbook.xml", WORKBOOK_XML);
  z.file("xl/_rels/workbook.xml.rels", WORKBOOK_RELS);
  z.file("xl/worksheets/sheet1.xml", SHEET_XML);
  z.file("xl/styles.xml", STYLES_XML);
  return z.generateAsync({ type: "arraybuffer" });
}
