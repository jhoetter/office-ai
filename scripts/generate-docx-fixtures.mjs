// Generate synthetic DOCX fixtures using the `docx` library.
// Run with `pnpm fixtures:docx` (see root package.json).
//
// Each fixture is a self-contained .docx file representing a category
// from spec/docx/feature-scope.md. They are NOT real-world documents —
// they are smoke fixtures that exercise our parser/serializer/handlers.
// See fixtures/docx/MANIFEST.md for the to-collect list of real-world
// docs we still need.

import { mkdir, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  HeadingLevel,
  AlignmentType,
  Table,
  TableRow,
  TableCell,
  WidthType,
} from "docx";

const here = dirname(fileURLToPath(import.meta.url));
const outRoot = resolve(here, "../fixtures/docx/synthetic");

async function write(name, doc) {
  await mkdir(outRoot, { recursive: true });
  const buf = await Packer.toBuffer(doc);
  const path = resolve(outRoot, `${name}.docx`);
  await writeFile(path, buf);
  console.log(`✓ wrote ${path} (${buf.length} bytes)`);
}

async function plain() {
  await write(
    "01-plain-paragraphs",
    new Document({
      creator: "officeAI",
      sections: [
        {
          properties: {},
          children: [
            new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Plain")] }),
            new Paragraph({ children: [new TextRun("First body paragraph.")] }),
            new Paragraph({ children: [new TextRun("Second body paragraph.")] }),
            new Paragraph({ children: [new TextRun("Third body paragraph.")] }),
          ],
        },
      ],
    })
  );
}

async function styled() {
  await write(
    "02-styled-runs",
    new Document({
      creator: "officeAI",
      sections: [
        {
          properties: {},
          children: [
            new Paragraph({
              heading: HeadingLevel.HEADING_1,
              children: [new TextRun({ text: "Styled runs", bold: true })],
            }),
            new Paragraph({
              children: [
                new TextRun("Plain "),
                new TextRun({ text: "bold", bold: true }),
                new TextRun(", "),
                new TextRun({ text: "italic", italics: true }),
                new TextRun(", "),
                new TextRun({ text: "underline", underline: { type: "single" } }),
                new TextRun(", "),
                new TextRun({ text: "colored", color: "9B59B6" }),
                new TextRun("."),
              ],
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun("Centered text.")],
            }),
          ],
        },
      ],
    })
  );
}

async function multiSection() {
  await write(
    "03-headings-and-body",
    new Document({
      creator: "officeAI",
      sections: [
        {
          properties: {},
          children: [
            new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Top level")] }),
            new Paragraph({ children: [new TextRun("Intro paragraph.")] }),
            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Subsection A")] }),
            new Paragraph({ children: [new TextRun("Body of subsection A.")] }),
            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Subsection B")] }),
            new Paragraph({ children: [new TextRun("Body of subsection B.")] }),
          ],
        },
      ],
    })
  );
}

async function withTable() {
  // Tables are P1 in our model — fixtures here are for round-trip
  // preservation tests (we should not destroy them on save).
  await write(
    "04-with-table",
    new Document({
      creator: "officeAI",
      sections: [
        {
          properties: {},
          children: [
            new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun("Schedule")] }),
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              rows: [
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph("Date")] }),
                    new TableCell({ children: [new Paragraph("Topic")] }),
                  ],
                }),
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph("2026-04-17")] }),
                    new TableCell({ children: [new Paragraph("Kickoff")] }),
                  ],
                }),
                new TableRow({
                  children: [
                    new TableCell({ children: [new Paragraph("2026-04-24")] }),
                    new TableCell({ children: [new Paragraph("Review")] }),
                  ],
                }),
              ],
            }),
            new Paragraph({ children: [new TextRun("Notes follow.")] }),
          ],
        },
      ],
    })
  );
}

async function long() {
  const blocks = [];
  blocks.push(
    new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Long body")] })
  );
  for (let i = 0; i < 60; i++) {
    blocks.push(
      new Paragraph({
        children: [
          new TextRun({
            text: `Paragraph ${i + 1}: lorem ipsum dolor sit amet, consectetur adipiscing elit. Phasellus sit amet quam laoreet, vehicula odio in, viverra elit.`,
          }),
        ],
      })
    );
  }
  await write(
    "05-long-body",
    new Document({ creator: "officeAI", sections: [{ properties: {}, children: blocks }] })
  );
}

await plain();
await styled();
await multiSection();
await withTable();
await long();
console.log("\nDone. See fixtures/docx/MANIFEST.md.");
