// Generate "real-shape" DOCX fixtures using the `docx` MIT library.
// Run with `pnpm fixtures-real` or `make fixtures-real`.
//
// These are deliberately produced by a third-party Word-grade emitter so
// they ship the parts a real Word/LibreOffice/Google-Docs export would
// ship: `word/styles.xml`, `word/numbering.xml`, header/footer parts,
// `word/comments.xml`, inline drawings, etc. They are checked in so the
// roundtrip test corpus is hermetic, but the manifest documents
// regeneration so we can refresh them when the spec evolves.
//
// Fixture inventory (kept aligned with fixtures/docx/MANIFEST.md):
//
//   01-styled-letter.docx        — headings + body + bullets + bold/italic
//   02-report-headers-footers.docx — multi-page report with header + footer
//   03-numbered-list.docx        — w:numPr + numbering.xml
//   04-table-grid.docx           — 3×4 table with header row styling
//   05-inline-image.docx         — inline ImageRun (1×1 PNG buffer)
//   06-comments-and-changes.docx — comments + tracked insertion/deletion
//
// All fixtures stay below 50 KB (verified at write time).

import { mkdir, readFile, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";
import {
  AlignmentType,
  CommentRangeEnd,
  CommentRangeStart,
  CommentReference,
  DeletedTextRun,
  Document,
  Footer,
  Header,
  HeadingLevel,
  ImageRun,
  InsertedTextRun,
  LevelFormat,
  PageBreak,
  PageOrientation,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from "docx";

const here = dirname(fileURLToPath(import.meta.url));
const outRoot = resolve(here, "../fixtures/docx/real-world");
const MAX_SIZE = 50 * 1024;

async function write(name, doc) {
  await mkdir(outRoot, { recursive: true });
  const buf = await Packer.toBuffer(doc);
  if (buf.length > MAX_SIZE) {
    throw new Error(`${name}.docx exceeded size budget: ${buf.length} > ${MAX_SIZE} bytes`);
  }
  const path = resolve(outRoot, `${name}.docx`);
  await writeFile(path, buf);
  console.log(`✓ wrote ${path} (${buf.length} bytes)`);
}

// ── 01 — styled letter ──────────────────────────────────────────────────
async function styledLetter() {
  await write(
    "01-styled-letter",
    new Document({
      creator: "office-ai",
      title: "Acme Quarterly Letter",
      sections: [
        {
          properties: {},
          children: [
            new Paragraph({
              heading: HeadingLevel.HEADING_1,
              children: [new TextRun("Quarterly letter")],
            }),
            new Paragraph({
              heading: HeadingLevel.HEADING_2,
              children: [new TextRun("Dear shareholders,")],
            }),
            new Paragraph({
              children: [
                new TextRun("This quarter we shipped "),
                new TextRun({ text: "office-ai P0", bold: true }),
                new TextRun(", an "),
                new TextRun({ text: "AI-native", italics: true }),
                new TextRun(" DOCX editor. Highlights:"),
              ],
            }),
            new Paragraph({ bullet: { level: 0 }, children: [new TextRun("Byte-preserving roundtrip")] }),
            new Paragraph({
              bullet: { level: 0 },
              children: [new TextRun("Headless command bus with approve / reject")],
            }),
            new Paragraph({
              bullet: { level: 0 },
              children: [new TextRun("ProseMirror surface that funnels every edit")],
            }),
            new Paragraph({
              children: [
                new TextRun("Sincerely,"),
                new TextRun({ break: 1 }),
                new TextRun({ text: "The office-ai team", italics: true }),
              ],
            }),
          ],
        },
      ],
    })
  );
}

// ── 02 — multi-page report with headers/footers ─────────────────────────
async function reportWithHeadersFooters() {
  const body = [
    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Quarterly Report")] }),
    new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Executive summary")] }),
  ];
  for (let i = 0; i < 30; i++) {
    body.push(
      new Paragraph({
        children: [
          new TextRun(
            `Section paragraph ${i + 1}. Lorem ipsum dolor sit amet, consectetur adipiscing elit. ` +
              `Phasellus sit amet quam laoreet, vehicula odio in, viverra elit.`
          ),
        ],
      })
    );
  }
  body.push(new Paragraph({ children: [new PageBreak()] }));
  body.push(new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Appendix")] }));
  for (let i = 0; i < 12; i++) {
    body.push(
      new Paragraph({
        children: [new TextRun(`Appendix line ${i + 1}: details continue across pages.`)],
      })
    );
  }
  await write(
    "02-report-headers-footers",
    new Document({
      creator: "office-ai",
      sections: [
        {
          properties: {},
          headers: {
            default: new Header({
              children: [
                new Paragraph({
                  alignment: AlignmentType.RIGHT,
                  children: [new TextRun({ text: "Acme Corp — Confidential", italics: true })],
                }),
              ],
            }),
          },
          footers: {
            default: new Footer({
              children: [
                new Paragraph({
                  alignment: AlignmentType.CENTER,
                  children: [new TextRun("Page footer · office-ai fixture")],
                }),
              ],
            }),
          },
          children: body,
        },
      ],
    })
  );
}

// ── 03 — numbered list ──────────────────────────────────────────────────
async function numberedList() {
  await write(
    "03-numbered-list",
    new Document({
      creator: "office-ai",
      numbering: {
        config: [
          {
            reference: "agenda",
            levels: [
              {
                level: 0,
                format: LevelFormat.DECIMAL,
                text: "%1.",
                alignment: AlignmentType.START,
              },
              {
                level: 1,
                format: LevelFormat.LOWER_LETTER,
                text: "%2)",
                alignment: AlignmentType.START,
              },
            ],
          },
        ],
      },
      sections: [
        {
          properties: {},
          children: [
            new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Agenda")] }),
            new Paragraph({
              numbering: { reference: "agenda", level: 0 },
              children: [new TextRun("Roundtrip review")],
            }),
            new Paragraph({
              numbering: { reference: "agenda", level: 1 },
              children: [new TextRun("Bytewise invariants")],
            }),
            new Paragraph({
              numbering: { reference: "agenda", level: 1 },
              children: [new TextRun("Edge cases")],
            }),
            new Paragraph({
              numbering: { reference: "agenda", level: 0 },
              children: [new TextRun("Agent surface")],
            }),
            new Paragraph({
              numbering: { reference: "agenda", level: 0 },
              children: [new TextRun("Demo and Q&A")],
            }),
          ],
        },
      ],
    })
  );
}

// ── 04 — table with header row ──────────────────────────────────────────
async function tableGrid() {
  const headerCell = (text) =>
    new TableCell({
      children: [new Paragraph({ children: [new TextRun({ text, bold: true })] })],
    });
  const cell = (text) => new TableCell({ children: [new Paragraph({ children: [new TextRun(text)] })] });
  await write(
    "04-table-grid",
    new Document({
      creator: "office-ai",
      sections: [
        {
          properties: {},
          children: [
            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Sprint plan")] }),
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              rows: [
                new TableRow({
                  tableHeader: true,
                  children: [headerCell("Week"), headerCell("Owner"), headerCell("Outcome")],
                }),
                new TableRow({ children: [cell("W1"), cell("Alex"), cell("Real-world fixtures")] }),
                new TableRow({ children: [cell("W2"), cell("Sam"), cell("PM funnel + range edits")] }),
                new TableRow({ children: [cell("W3"), cell("Jordan"), cell("Agent reach (CLI + MCP)")] }),
              ],
            }),
            new Paragraph({ children: [new TextRun("Notes captured below.")] }),
          ],
        },
      ],
    })
  );
}

// ── 05 — inline image (1×1 solid-color PNG) ─────────────────────────────
//
// Hand-crafted 1×1 opaque-blue PNG; small enough to inline as a literal.
// Verified via `file <(echo -n ...)` to be a valid PNG. We embed it via
// docx's ImageRun so the resulting .docx ships:
//   - word/media/image1.png
//   - word/_rels/document.xml.rels with the image relationship
//   - the inline drawing markup in word/document.xml
const ONE_PIXEL_PNG = Buffer.from(
  "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c489" +
    "0000000d49444154789c63304e3bf31f00049a026573f9061f0000000049454e44ae426082",
  "hex"
);

async function inlineImage() {
  await write(
    "05-inline-image",
    new Document({
      creator: "office-ai",
      sections: [
        {
          properties: {},
          children: [
            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("With logo")] }),
            new Paragraph({
              children: [
                new TextRun("Logo: "),
                new ImageRun({
                  type: "png",
                  data: ONE_PIXEL_PNG,
                  transformation: { width: 24, height: 24 },
                }),
                new TextRun(" — inline image rendered via docx ImageRun."),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun("This fixture exercises an inline drawing relationship and a media part."),
              ],
            }),
          ],
        },
      ],
    })
  );
}

// ── 06 — comments + tracked changes ─────────────────────────────────────
async function commentsAndChanges() {
  await write(
    "06-comments-and-changes",
    new Document({
      creator: "office-ai",
      comments: {
        children: [
          {
            id: 0,
            author: "Reviewer A",
            initials: "RA",
            date: new Date("2026-04-17T10:00:00Z"),
            children: [new Paragraph("Tighten this opening sentence.")],
          },
        ],
      },
      sections: [
        {
          properties: {},
          children: [
            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Draft")] }),
            new Paragraph({
              children: [
                new CommentRangeStart(0),
                new TextRun("This document was prepared as a fixture for office-ai."),
                new CommentRangeEnd(0),
                new TextRun({ children: [new CommentReference(0)] }),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun("It also contains a "),
                new InsertedTextRun({
                  text: "tracked insertion",
                  id: 1,
                  author: "Reviewer A",
                  date: "2026-04-17T10:05:00Z",
                }),
                new TextRun(" and a "),
                new DeletedTextRun({
                  text: "deleted phrase",
                  id: 2,
                  author: "Reviewer A",
                  date: "2026-04-17T10:06:00Z",
                }),
                new TextRun("."),
              ],
            }),
          ],
        },
      ],
    })
  );
}

// ── 08 — shaded callout (single-cell table with fill + bold colored text) ──
//
// Mirrors the "Theoretische Grundlage" / "Definition" callouts in the DSR
// thesis chapter: a 1×1 table whose cell carries `<w:shd w:fill="…"/>`
// plus a paragraph with a colored bold run. Exercises Phase 2 (typed table
// shading + run-mark fidelity).
async function shadedCalloutTable() {
  const filledCell = new TableCell({
    width: { size: 100, type: WidthType.PERCENTAGE },
    shading: { fill: "DDEBF7", type: ShadingType.CLEAR, color: "auto" },
    children: [
      new Paragraph({
        children: [new TextRun({ text: "Theoretical foundation:", bold: true, color: "1B3A5C" })],
      }),
      new Paragraph({
        children: [
          new TextRun(
            "Design Science Research is a research paradigm that focuses on " +
              "creating innovative artifacts that solve real-world problems."
          ),
        ],
      }),
    ],
  });
  await write(
    "08-shaded-callout-table",
    new Document({
      creator: "office-ai",
      sections: [
        {
          properties: {},
          children: [
            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Methodology")] }),
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              rows: [new TableRow({ children: [filledCell] })],
            }),
            new Paragraph({
              children: [
                new TextRun(
                  "The callout above is a 1×1 table with cell shading. Word and the DSR " +
                    "thesis use this pattern for definition / aside boxes."
                ),
              ],
            }),
          ],
        },
      ],
    })
  );
}

// ── 09 — multi-column section ───────────────────────────────────────────
//
// One section configured for two equal-width columns with column spacing.
// Exercises Phase 5 (per-section <w:cols> projection + page-decoration grid).
async function multiColumnSection() {
  const para = (text) => new Paragraph({ children: [new TextRun(text)] });
  await write(
    "09-multi-column-section",
    new Document({
      creator: "office-ai",
      sections: [
        {
          properties: {
            column: { count: 2, space: 720, equalWidth: true },
          },
          children: [
            new Paragraph({
              heading: HeadingLevel.HEADING_1,
              children: [new TextRun("Newsletter")],
            }),
            para(
              "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Phasellus sit amet " +
                "quam laoreet, vehicula odio in, viverra elit. Suspendisse potenti."
            ),
            para(
              "Vivamus rutrum tincidunt nibh, vitae ullamcorper magna pharetra eget. Nulla " +
                "facilisi. Aliquam erat volutpat. Donec quis arcu dapibus nibh."
            ),
            para(
              "Pellentesque habitant morbi tristique senectus et netus et malesuada fames " +
                "ac turpis egestas. Curabitur vitae ipsum nec turpis tristique."
            ),
            para(
              "Nullam id orci eget mauris malesuada feugiat ut ut leo. Etiam suscipit purus " +
                "id ipsum aliquet, vitae bibendum risus accumsan."
            ),
            para(
              "Suspendisse potenti. Aenean rhoncus malesuada ipsum, eget viverra justo " +
                "vehicula nec. Praesent volutpat sapien at quam laoreet eleifend."
            ),
          ],
        },
      ],
    })
  );
}

// ── 10 — landscape section break ────────────────────────────────────────
//
// Two sections: the first portrait, the second landscape. Exercises Phase 5
// (per-section page-geometry switching) and Phase 1's per-chunk filler.
async function landscapeSection() {
  const para = (text) => new Paragraph({ children: [new TextRun(text)] });
  await write(
    "10-landscape-section",
    new Document({
      creator: "office-ai",
      sections: [
        {
          properties: {
            page: { size: { orientation: PageOrientation.PORTRAIT } },
          },
          children: [
            new Paragraph({
              heading: HeadingLevel.HEADING_1,
              children: [new TextRun("Portrait page")],
            }),
            para('Body content sized for portrait (8.5" × 11").'),
            para("More portrait body content."),
          ],
        },
        {
          properties: {
            page: { size: { orientation: PageOrientation.LANDSCAPE } },
          },
          children: [
            new Paragraph({
              heading: HeadingLevel.HEADING_1,
              children: [new TextRun("Landscape page")],
            }),
            para('Body content sized for landscape (11" × 8.5").'),
            para("Used for wide tables, charts, screenshots."),
          ],
        },
      ],
    })
  );
}

// ── 11 — anchored textbox (mc:AlternateContent + wps:txbx) ─────────────
//
// Word's high-level API in the `docx` npm package does not emit textboxes,
// but real Word/LibreOffice docs use them constantly. We hand-craft this
// fixture by re-using the styled-letter as a chassis (so it carries
// styles.xml, theme1.xml, fontTable.xml, settings.xml, etc.) and replacing
// `word/document.xml` with a body that contains a `mc:AlternateContent`
// drawing wrapping a `wps:wsp` shape with `wps:txbx` inner content.
//
// The textbox carries a small fill, a black stroke, and one paragraph of
// inner text. Round-trip is byte-equal because the parser keeps the full
// `mc:AlternateContent` subtree in its `raw` cache (see Phase 4 plan).
const TEXTBOX_DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
    xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
    xmlns:v="urn:schemas-microsoft-com:vml"
    xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:w10="urn:schemas-microsoft-com:office:word"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
    xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
    xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
    xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
    xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
    xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
    xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
    mc:Ignorable="w14 w15 wp14"><w:body><w:p><w:r><w:t xml:space="preserve">Body paragraph before the textbox.</w:t></w:r></w:p><w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="45720" distB="45720" distL="114300" distR="114300" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>1828800</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="2743200" cy="914400"/><wp:effectExtent l="0" t="0" r="19050" b="19050"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="1" name="Text Box 1"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFF2CC"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr><wps:txbx><w:txbxContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Note:</w:t></w:r></w:p><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">This is a floating textbox anchored to the page margin.</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr rot="0" spcFirstLastPara="0" vertOverflow="overflow" horzOverflow="overflow" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" numCol="1" spcCol="0" rtlCol="0" fromWordArt="0" anchor="t" anchorCtr="0" forceAA="0" compatLnSpc="1"><a:prstTxWarp prst="textNoShape"><a:avLst/></a:prstTxWarp><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape id="Text Box 1" o:spid="_x0000_s1026" type="#_x0000_t202" style="position:absolute;margin-left:144pt;margin-top:0;width:216pt;height:72pt;z-index:251659264;mso-wrap-distance-left:9pt;mso-wrap-distance-top:3.6pt;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:3.6pt;mso-position-horizontal-relative:margin;mso-position-vertical-relative:text" fillcolor="#fff2cc" strokecolor="black" strokeweight=".75pt"><v:textbox><w:txbxContent><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Note:</w:t></w:r></w:p><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">This is a floating textbox anchored to the page margin.</w:t></w:r></w:p></w:txbxContent></v:textbox><w10:wrap type="square" anchorx="margin"/></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r><w:r><w:t xml:space="preserve">Inline text continuing after the anchor.</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">Body paragraph after the textbox.</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/><w:cols w:space="720"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>`;

async function textboxFixture() {
  // Use the styled-letter as a chassis: it ships styles.xml, theme1.xml,
  // fontTable.xml, settings.xml, etc. Replace document.xml with our
  // hand-rolled textbox-bearing body.
  const chassisPath = resolve(outRoot, "01-styled-letter.docx");
  const chassisBytes = await readFile(chassisPath);
  const zip = await JSZip.loadAsync(chassisBytes);
  zip.file("word/document.xml", TEXTBOX_DOCUMENT_XML);
  const buf = await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
  if (buf.length > MAX_SIZE) {
    throw new Error(`11-textbox.docx exceeded size budget: ${buf.length} > ${MAX_SIZE} bytes`);
  }
  const path = resolve(outRoot, "11-textbox.docx");
  await writeFile(path, buf);
  console.log(`✓ wrote ${path} (${buf.length} bytes)`);
}

await styledLetter();
await reportWithHeadersFooters();
await numberedList();
await tableGrid();
await inlineImage();
await commentsAndChanges();
await shadedCalloutTable();
await multiColumnSection();
await landscapeSection();
await textboxFixture();

console.log("\nDone. See fixtures/docx/MANIFEST.md.");
