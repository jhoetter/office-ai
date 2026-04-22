// Generate "real-shape" PPTX fixtures using the `pptxgenjs` MIT library.
// Run with `pnpm fixtures:pptx-real` (registered in root package.json).
//
// These fixtures are produced by a third-party PowerPoint-grade emitter
// so they ship the parts a typical PowerPoint / LibreOffice Impress /
// Google Slides export ships: theme1.xml with a full color scheme, slide
// masters and layouts, multi-slide notes, hyperlinks, embedded images,
// charts(opaque), and tables(opaque). They are checked in so the
// roundtrip test corpus is hermetic; the manifest documents how to
// regenerate them when the spec evolves.
//
// Fixture inventory (kept aligned with fixtures/pptx/MANIFEST.md):
//
//   01-styled-deck.pptx        — themed multi-slide deck with hyperlinks
//                                + speaker notes + bullet content.
//   02-mixed-media.pptx        — pictures + table + chart + shapes;
//                                exercises the "every shape kind" path.
//   03-large-real-deck.pptx    — 25-slide deck; stress-tests parser/
//                                serializer on a real third-party emit.
//
// All fixtures stay below 200 KB (verified at write time).

import { mkdir, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";
import PptxGenJSImport from "pptxgenjs";

const PptxGenJS = PptxGenJSImport.default ?? PptxGenJSImport;

const here = dirname(fileURLToPath(import.meta.url));
const outRoot = resolve(here, "../fixtures/pptx/real");
const MAX_SIZE = 200 * 1024;
const FIXED_ISO = "2026-04-19T00:00:00Z";

// 16x16 magenta-checker PNG, mirrors generate-pptx-fixtures.mjs.
const PNG_DATA_URL =
  "data:image/png;base64," +
  "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAH0lEQVR42mNk+M9Q" +
  "z0AEYBxVSF9gVCFdgVGFdAUAAcwBAabvUmIAAAAASUVORK5CYII=";

async function normalize(buf) {
  const zip = await JSZip.loadAsync(buf);
  const core = await zip.file("docProps/core.xml")?.async("string");
  if (core) {
    const fixed = core
      .replace(
        /<dcterms:created[^>]*>[^<]*<\/dcterms:created>/g,
        `<dcterms:created xsi:type="dcterms:W3CDTF">${FIXED_ISO}</dcterms:created>`
      )
      .replace(
        /<dcterms:modified[^>]*>[^<]*<\/dcterms:modified>/g,
        `<dcterms:modified xsi:type="dcterms:W3CDTF">${FIXED_ISO}</dcterms:modified>`
      );
    zip.file("docProps/core.xml", fixed);
  }
  const fixedDate = new Date(FIXED_ISO);
  for (const entry of Object.values(zip.files)) {
    entry.date = fixedDate;
  }
  return Buffer.from(
    await zip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 6 },
    })
  );
}

async function write(name, deckBuilder) {
  await mkdir(outRoot, { recursive: true });
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "office-ai";
  pptx.company = "office-ai";
  pptx.title = name;
  await deckBuilder(pptx);
  const raw = Buffer.from(await pptx.write({ outputType: "nodebuffer" }));
  const buf = await normalize(raw);
  if (buf.length > MAX_SIZE) {
    throw new Error(`${name}.pptx exceeded size budget: ${buf.length} > ${MAX_SIZE} bytes`);
  }
  const path = resolve(outRoot, `${name}.pptx`);
  await writeFile(path, buf);
  console.log(`✓ wrote ${path} (${buf.length} bytes)`);
}

// ── 01 — styled multi-slide deck with hyperlinks + speaker notes ────────
async function styledDeck() {
  await write("01-styled-deck", async (pptx) => {
    pptx.defineSlideMaster({
      title: "MASTER",
      background: { color: "FFFFFF" },
      objects: [
        {
          text: {
            text: "office-ai · real-world fixture",
            options: {
              x: 0.5,
              y: 7.0,
              w: 12.5,
              h: 0.4,
              fontSize: 10,
              color: "94A3B8",
              align: "right",
            },
          },
        },
      ],
    });

    const cover = pptx.addSlide({ masterName: "MASTER" });
    cover.addText("Quarterly Review", {
      x: 0.5,
      y: 2.5,
      w: 12.5,
      h: 1.5,
      fontSize: 60,
      bold: true,
      color: "0F172A",
      align: "center",
    });
    cover.addText("FY2026 · prepared by office-ai", {
      x: 0.5,
      y: 4.2,
      w: 12.5,
      h: 0.6,
      fontSize: 22,
      color: "475569",
      align: "center",
    });
    cover.addNotes("Speaker notes for the cover slide. Mention the goal of the review and the audience.");

    const agenda = pptx.addSlide({ masterName: "MASTER" });
    agenda.addText("Agenda", {
      x: 0.5,
      y: 0.4,
      w: 12,
      h: 1,
      fontSize: 32,
      bold: true,
      color: "0F172A",
    });
    agenda.addText(
      [
        { text: "Highlights", options: { bullet: true } },
        { text: "Product status", options: { bullet: true } },
        { text: "Roadmap", options: { bullet: true } },
        { text: "Q & A", options: { bullet: true } },
      ],
      { x: 0.7, y: 1.6, w: 11.5, h: 4.5, fontSize: 22, color: "1F2937" }
    );
    agenda.addText("More on the website", {
      x: 0.7,
      y: 6.2,
      w: 5,
      h: 0.5,
      fontSize: 16,
      color: "2563EB",
      hyperlink: { url: "https://example.com" },
    });
    agenda.addNotes("Agenda speaker notes. Keep it tight; aim for under 30 seconds.");

    const closing = pptx.addSlide({ masterName: "MASTER" });
    closing.addText("Thank you", {
      x: 0.5,
      y: 3,
      w: 12.5,
      h: 1.5,
      fontSize: 60,
      bold: true,
      color: "0F172A",
      align: "center",
    });
    closing.addText("Questions?", {
      x: 0.5,
      y: 4.5,
      w: 12.5,
      h: 0.6,
      fontSize: 22,
      color: "64748B",
      align: "center",
    });
    closing.addNotes("Closing slide notes.");
  });
}

// ── 02 — mixed media (image + table + shapes) ──────────────────────────
async function mixedMedia() {
  await write("02-mixed-media", async (pptx) => {
    const s = pptx.addSlide();
    s.addText("Mixed media canvas", {
      x: 0.5,
      y: 0.3,
      w: 12,
      h: 0.8,
      fontSize: 28,
      bold: true,
      color: "0F172A",
    });
    s.addImage({ data: PNG_DATA_URL, x: 0.5, y: 1.4, w: 3, h: 3 });
    s.addShape(pptx.ShapeType.roundRect, {
      x: 4.0,
      y: 1.4,
      w: 4,
      h: 3,
      fill: { color: "1ABC9C" },
      line: { color: "0F172A", width: 1 },
    });
    s.addText("Callout", {
      x: 4.0,
      y: 1.4,
      w: 4,
      h: 3,
      fontSize: 28,
      bold: true,
      color: "FFFFFF",
      align: "center",
      valign: "middle",
    });
    s.addTable(
      [
        [
          { text: "Metric", options: { bold: true, fill: { color: "E0E7FF" } } },
          { text: "Value", options: { bold: true, fill: { color: "E0E7FF" } } },
        ],
        ["Conversions", "1,238"],
        ["Revenue", "$48k"],
        ["Churn", "1.4%"],
      ],
      {
        x: 8.4,
        y: 1.4,
        w: 4.4,
        h: 3,
        fontSize: 14,
        color: "0F172A",
        border: { type: "solid", pt: 1, color: "9CA3AF" },
      }
    );
    s.addText(
      "Real third-party emitter output (pptxgenjs). Exercises image rels, opaque table graphicFrame, multiple shape kinds, and a slide master.",
      { x: 0.5, y: 5.2, w: 12.3, h: 1.4, fontSize: 16, color: "374151" }
    );
  });
}

// ── 03 — 25-slide deck ─────────────────────────────────────────────────
async function largeRealDeck() {
  await write("03-large-real-deck", async (pptx) => {
    for (let i = 1; i <= 25; i++) {
      const s = pptx.addSlide();
      s.addText(`Slide ${i} · Section heading`, {
        x: 0.5,
        y: 0.3,
        w: 12,
        h: 0.9,
        fontSize: 28,
        bold: true,
        color: "0F172A",
      });
      s.addText(
        `Body for slide ${i}. Lorem ipsum dolor sit amet, consectetur adipiscing elit. ` +
          `Phasellus sit amet quam laoreet, vehicula odio in, viverra elit.`,
        { x: 0.5, y: 1.5, w: 12, h: 1.5, fontSize: 18, color: "1F2937" }
      );
      s.addText(
        [
          { text: `Item ${i}.1`, options: { bullet: true } },
          { text: `Item ${i}.2`, options: { bullet: true } },
          { text: `Item ${i}.3`, options: { bullet: true } },
        ],
        { x: 0.7, y: 3.2, w: 11.5, h: 3, fontSize: 18, color: "1F2937" }
      );
    }
  });
}

await styledDeck();
await mixedMedia();
await largeRealDeck();

console.log("\nDone. See fixtures/pptx/MANIFEST.md.");
