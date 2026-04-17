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

import { mkdir, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
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
  Packer,
  Paragraph,
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
      creator: "officeAI",
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
                new TextRun({ text: "officeAI P0", bold: true }),
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
                new TextRun({ text: "The officeAI team", italics: true }),
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
      creator: "officeAI",
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
                  children: [new TextRun("Page footer · officeAI fixture")],
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
      creator: "officeAI",
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
  const cell = (text) =>
    new TableCell({ children: [new Paragraph({ children: [new TextRun(text)] })] });
  await write(
    "04-table-grid",
    new Document({
      creator: "officeAI",
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
      creator: "officeAI",
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
                new TextRun(
                  "This fixture exercises an inline drawing relationship and a media part."
                ),
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
      creator: "officeAI",
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
                new TextRun("This document was prepared as a fixture for officeAI."),
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

await styledLetter();
await reportWithHeadersFooters();
await numberedList();
await tableGrid();
await inlineImage();
await commentsAndChanges();

console.log("\nDone. See fixtures/docx/MANIFEST.md.");
