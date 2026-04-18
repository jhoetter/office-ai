// Generate synthetic PPTX fixtures using the `pptxgenjs` library (MIT).
// Run with `pnpm fixtures:pptx` (see root package.json).
//
// Each fixture is a self-contained .pptx file representing a category
// from spec/pptx/feature-scope.md. They are NOT real-world documents —
// they are smoke fixtures that exercise our parser/serializer/handlers.
// See fixtures/pptx/MANIFEST.md for the to-collect list of real-world
// decks we still need.

import { mkdir, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";
import PptxGenJSImport from "pptxgenjs";

// pptxgenjs ships as CJS; ESM default-import yields the constructor.
const PptxGenJS = PptxGenJSImport.default ?? PptxGenJSImport;

const here = dirname(fileURLToPath(import.meta.url));
const outRoot = resolve(here, "../fixtures/pptx/synthetic");

// Stable timestamp baked into every fixture so SHA-256 is reproducible.
const FIXED_ISO = "2026-04-01T00:00:00Z";

/**
 * pptxgenjs writes `dcterms:created` / `dcterms:modified` from `Date.now()`,
 * which makes byte-equality non-deterministic across runs. We post-process
 * the generated zip to normalize those timestamps + the zip entry mtimes.
 */
async function normalize(buf) {
  const zip = await JSZip.loadAsync(buf);
  // Override docProps/core.xml dcterms timestamps.
  const core = await zip.file("docProps/core.xml")?.async("string");
  if (core) {
    const fixedCore = core
      .replace(/<dcterms:created[^>]*>[^<]*<\/dcterms:created>/g, `<dcterms:created xsi:type="dcterms:W3CDTF">${FIXED_ISO}</dcterms:created>`)
      .replace(/<dcterms:modified[^>]*>[^<]*<\/dcterms:modified>/g, `<dcterms:modified xsi:type="dcterms:W3CDTF">${FIXED_ISO}</dcterms:modified>`);
    zip.file("docProps/core.xml", fixedCore);
  }
  // Set a fixed mtime on every entry so the zip CRCs / extra fields are stable.
  const fixedDate = new Date(FIXED_ISO);
  for (const entry of Object.values(zip.files)) {
    entry.date = fixedDate;
  }
  return Buffer.from(await zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 6 } }));
}

async function write(name, deckBuilder) {
  await mkdir(outRoot, { recursive: true });
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in (16:9)
  pptx.author = "officeAI";
  pptx.company = "officeAI";
  pptx.title = name;
  await deckBuilder(pptx);
  const raw = Buffer.from(await pptx.write({ outputType: "nodebuffer" }));
  const buf = await normalize(raw);
  const path = resolve(outRoot, `${name}.pptx`);
  await writeFile(path, buf);
  console.log(`✓ wrote ${path} (${buf.length} bytes)`);
}

async function blank() {
  await write("01-blank", async (pptx) => {
    pptx.addSlide();
  });
}

async function titleOnly() {
  await write("02-title-only", async (pptx) => {
    const slide = pptx.addSlide();
    slide.addText("Hello, PPTX", {
      x: 0.5, y: 0.5, w: 12.0, h: 1.5,
      fontSize: 44, bold: true, color: "111827",
    });
  });
}

async function titleAndContent() {
  await write("03-title-and-content", async (pptx) => {
    const slide = pptx.addSlide();
    slide.addText("Agenda", {
      x: 0.5, y: 0.4, w: 12.0, h: 1.0,
      fontSize: 32, bold: true, color: "111827",
    });
    slide.addText(
      [
        { text: "Background", options: { bullet: true } },
        { text: "Approach", options: { bullet: true } },
        { text: "Results", options: { bullet: true } },
        { text: "Next steps", options: { bullet: true } },
      ],
      { x: 0.7, y: 1.6, w: 11.5, h: 4.5, fontSize: 22, color: "1F2937" },
    );
  });
}

async function multiShape() {
  await write("04-multi-shape", async (pptx) => {
    const slide = pptx.addSlide();
    slide.addText("Multi-shape canvas", {
      x: 0.5, y: 0.3, w: 12, h: 0.8, fontSize: 28, bold: true,
    });
    slide.addShape(pptx.ShapeType.rect, {
      x: 0.6, y: 1.5, w: 3, h: 2,
      fill: { color: "9B59B6" },
      line: { color: "FFFFFF" },
    });
    slide.addShape(pptx.ShapeType.ellipse, {
      x: 4.0, y: 1.5, w: 3, h: 2,
      fill: { color: "3498DB" },
      line: { color: "FFFFFF" },
    });
    slide.addShape(pptx.ShapeType.roundRect, {
      x: 7.5, y: 1.5, w: 3, h: 2,
      fill: { color: "1ABC9C" },
      line: { color: "FFFFFF" },
    });
    slide.addText("Boxed text", {
      x: 0.6, y: 4.0, w: 5, h: 1.0,
      fontSize: 20, color: "111827",
      fill: { color: "F3F4F6" },
    });
    slide.addText("Another box", {
      x: 6.0, y: 4.0, w: 5, h: 1.0,
      fontSize: 20, italic: true, color: "374151",
      fill: { color: "FEF3C7" },
    });
  });
}

async function withImage() {
  // Embed a tiny 16x16 PNG as a base64 data URL so the fixture is hermetic.
  // 16x16 transparent-magenta-checker PNG.
  const pngDataUrl =
    "data:image/png;base64," +
    "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAH0lEQVR42mNk+M9Q" +
    "z0AEYBxVSF9gVCFdgVGFdAUAAcwBAabvUmIAAAAASUVORK5CYII=";

  await write("05-with-image", async (pptx) => {
    const slide = pptx.addSlide();
    slide.addText("Picture below", {
      x: 0.5, y: 0.4, w: 12, h: 0.8,
      fontSize: 28, bold: true,
    });
    slide.addImage({
      data: pngDataUrl,
      x: 1.0, y: 1.5, w: 4, h: 4,
    });
    slide.addText("This deck embeds one PNG to exercise media + relationships.", {
      x: 5.5, y: 2.0, w: 7, h: 1.5,
      fontSize: 18, color: "374151",
    });
  });
}

async function withTable() {
  // P0 round-trips tables as opaque graphicFrame; this fixture verifies
  // that path stays clean.
  await write("06-with-table", async (pptx) => {
    const slide = pptx.addSlide();
    slide.addText("Quarterly results", {
      x: 0.5, y: 0.4, w: 12, h: 0.8,
      fontSize: 28, bold: true,
    });
    slide.addTable(
      [
        [
          { text: "Quarter", options: { bold: true } },
          { text: "Revenue", options: { bold: true } },
          { text: "Growth", options: { bold: true } },
        ],
        ["Q1", "$120k", "+12%"],
        ["Q2", "$145k", "+21%"],
        ["Q3", "$162k", "+12%"],
        ["Q4", "$210k", "+30%"],
      ],
      {
        x: 0.7, y: 1.6, w: 11.5,
        colW: [3, 4, 4.5],
        fontSize: 18, color: "111827",
        border: { type: "solid", pt: 1, color: "9CA3AF" },
      },
    );
  });
}

async function multiSlide() {
  await write("07-multi-slide", async (pptx) => {
    const titles = [
      "Welcome",
      "Problem statement",
      "Approach",
      "Architecture",
      "Demo",
      "Results",
      "Next steps",
      "Q & A",
    ];
    for (const t of titles) {
      const slide = pptx.addSlide();
      slide.addText(t, {
        x: 0.5, y: 2.5, w: 12, h: 2,
        fontSize: 48, bold: true, align: "center", color: "111827",
      });
    }
  });
}

async function largeDeck() {
  await write("08-large-deck", async (pptx) => {
    for (let i = 1; i <= 50; i += 1) {
      const slide = pptx.addSlide();
      slide.addText(`Slide ${i}`, {
        x: 0.5, y: 0.3, w: 12, h: 0.9,
        fontSize: 28, bold: true,
      });
      slide.addText(
        `Body text for slide ${i}. Lorem ipsum dolor sit amet, consectetur adipiscing elit. Phasellus sit amet quam laoreet, vehicula odio in, viverra elit.`,
        { x: 0.5, y: 1.4, w: 12, h: 5, fontSize: 18, color: "1F2937" },
      );
    }
  });
}

async function withChart() {
  // F3: chart fixture exercises the chart parser/serializer/commands.
  // Picks a clustered bar chart with 2 series so set-chart-type, set-chart-data
  // and set-chart-title each have something to round-trip.
  await write("09-with-chart", async (pptx) => {
    const slide = pptx.addSlide();
    slide.addText("Quarterly performance", {
      x: 0.5, y: 0.4, w: 12, h: 0.8,
      fontSize: 28, bold: true, color: "111827",
    });
    slide.addChart(
      pptx.ChartType.bar,
      [
        {
          name: "Revenue",
          labels: ["Q1", "Q2", "Q3", "Q4"],
          values: [120, 145, 162, 210],
        },
        {
          name: "Expenses",
          labels: ["Q1", "Q2", "Q3", "Q4"],
          values: [80, 95, 110, 130],
        },
      ],
      {
        x: 0.5, y: 1.5, w: 12, h: 5,
        showTitle: true,
        title: "Revenue vs Expenses (k$)",
        showLegend: true,
        legendPos: "b",
        catAxisLabelColor: "1F2937",
        valAxisLabelColor: "1F2937",
      },
    );
  });
}

async function withAnimations() {
  // F4: slide-transition + simple per-shape entrance animations fixture.
  // pptxgenjs doesn't expose a high-level animations API, so this fixture
  // is built minimally and the F4 implementation will use a real-world
  // deck (added under fixtures/pptx/real/) for full animation coverage.
  await write("10-with-anim", async (pptx) => {
    const slide = pptx.addSlide();
    slide.addText("Animated title", {
      x: 0.5, y: 0.5, w: 12, h: 1.5,
      fontSize: 44, bold: true, color: "111827",
    });
    slide.addText("Body line that should appear after the title.", {
      x: 0.5, y: 2.5, w: 12, h: 1.0,
      fontSize: 24, color: "1F2937",
    });
  });
}

await blank();
await titleOnly();
await titleAndContent();
await multiShape();
await withImage();
await withTable();
await multiSlide();
await largeDeck();
await withChart();
await withAnimations();
console.log("\nDone. See fixtures/pptx/MANIFEST.md.");
