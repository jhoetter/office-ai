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
async function normalize(buf, patchZip) {
  const zip = await JSZip.loadAsync(buf);
  // Optional caller patch (e.g. inject <p:transition>/<p:timing> bytes).
  if (patchZip) await patchZip(zip);
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

async function write(name, deckBuilder, patchZip) {
  await mkdir(outRoot, { recursive: true });
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE"; // 13.333 x 7.5 in (16:9)
  pptx.author = "officeAI";
  pptx.company = "officeAI";
  pptx.title = name;
  await deckBuilder(pptx);
  const raw = Buffer.from(await pptx.write({ outputType: "nodebuffer" }));
  const buf = await normalize(raw, patchZip);
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
  // pptxgenjs doesn't expose a high-level animations API, so we inject a
  // realistic <p:transition> + <p:timing> tail directly into ppt/slides/
  // slide1.xml after pptxgenjs writes the deck. This gives us a typed
  // fade transition and two entrance animations (appear on shape 2,
  // fly-in on shape 3) for parser/serializer round-tripping.
  await write(
    "10-with-anim",
    async (pptx) => {
      const slide = pptx.addSlide();
      slide.addText("Animated title", {
        x: 0.5, y: 0.5, w: 12, h: 1.5,
        fontSize: 44, bold: true, color: "111827",
      });
      slide.addText("Body line that should appear after the title.", {
        x: 0.5, y: 2.5, w: 12, h: 1.0,
        fontSize: 24, color: "1F2937",
      });
    },
    async (zip) => {
      const slidePath = "ppt/slides/slide1.xml";
      const slideXml = await zip.file(slidePath)?.async("string");
      if (!slideXml) return;
      const insert =
        `<p:transition spd="med"><p:fade/></p:transition>` +
        `<p:timing>` +
        `<p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">` +
        `<p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq">` +
        `<p:childTnLst>` +
        // First entrance: appear on shape with cNvPrId=2 (the title body)
        `<p:par><p:cTn id="3" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" nodeType="clickEffect">` +
        `<p:childTnLst><p:set><p:cBhvr><p:cTn id="4" dur="1" fill="hold"/><p:tgtEl><p:spTgt spid="2"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set></p:childTnLst>` +
        `</p:cTn></p:par>` +
        // Second entrance: fly-in on shape with cNvPrId=3 (the second body)
        `<p:par><p:cTn id="5" presetID="2" presetClass="entr" presetSubtype="4" fill="hold" nodeType="afterEffect" dur="500">` +
        `<p:childTnLst><p:set><p:cBhvr><p:cTn id="6" dur="1" fill="hold"/><p:tgtEl><p:spTgt spid="3"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set></p:childTnLst>` +
        `</p:cTn></p:par>` +
        `</p:childTnLst>` +
        `</p:cTn></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>` +
        `</p:timing>`;
      // Splice before </p:sld>.
      const patched = slideXml.replace("</p:sld>", `${insert}</p:sld>`);
      zip.file(slidePath, patched);
    }
  );
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
await animationsGallery();
console.log("\nDone. See fixtures/pptx/MANIFEST.md.");

// ─── 11-animations-gallery ───────────────────────────────────────────────
//
// F4 v2: a "golden" fixture that exercises representatives of all four
// animation categories — Entrance, Emphasis, Exit, Motion Path — on a
// single slide. Useful as a regression bookmark for the typed
// (category, preset, direction, trigger, delay) round-trip and for
// manually opening in PowerPoint to verify visual fidelity.
//
// The OOXML below is hand-rolled because pptxgenjs has no animations
// API. It targets two shapes (cNvPrId=2 and cNvPrId=3) and stages:
//   1. Entrance / Fly In (left)               on shape 2, click
//   2. Emphasis / Spin (clockwise)            on shape 2, withPrev
//   3. Exit / Fade                            on shape 2, afterPrev
//   4. Entrance / Wipe (up)                   on shape 3, click, delay 500
//   5. Motion Path / Arc                      on shape 3, withPrev
async function animationsGallery() {
  await write(
    "11-animations-gallery",
    async (pptx) => {
      const slide = pptx.addSlide();
      slide.addText("Animation gallery", {
        x: 0.5, y: 0.5, w: 12, h: 1.5,
        fontSize: 44, bold: true, color: "111827",
      });
      slide.addText("Hover targets for emphasis + exit + motion paths.", {
        x: 0.5, y: 2.5, w: 12, h: 1.0,
        fontSize: 24, color: "1F2937",
      });
    },
    async (zip) => {
      const slidePath = "ppt/slides/slide1.xml";
      const slideXml = await zip.file(slidePath)?.async("string");
      if (!slideXml) return;
      const par = (id, presetClass, presetID, subtype, nodeType, body, extraAttrs = "") =>
        `<p:par><p:cTn id="${id}" presetID="${presetID}" presetClass="${presetClass}" presetSubtype="${subtype}" fill="hold" nodeType="${nodeType}"${extraAttrs}>` +
        `<p:childTnLst>${body}</p:childTnLst></p:cTn></p:par>`;
      const setVis = (childId, spid, value) =>
        `<p:set><p:cBhvr><p:cTn id="${childId}" dur="1" fill="hold"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>` +
        `<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>` +
        `<p:to><p:strVal val="${value}"/></p:to></p:set>`;
      const animMotion = (childId, spid, dur, path) =>
        `<p:animMotion origin="layout" path="${path}" pathEditMode="relative">` +
        `<p:cBhvr><p:cTn id="${childId}" dur="${dur}" fill="hold"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr></p:animMotion>`;
      const animRot = (childId, spid, dur, fromDeg, toDeg) =>
        `<p:animRot by="${(toDeg - fromDeg) * 60000}"><p:cBhvr><p:cTn id="${childId}" dur="${dur}" fill="hold"/>` +
        `<p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl></p:cBhvr></p:animRot>`;
      const insert =
        `<p:timing>` +
        `<p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot">` +
        `<p:childTnLst><p:seq concurrent="1" nextAc="seek">` +
        `<p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>` +
        // 1. Entrance / Fly In (left) on shape 2
        par(3, "entr", 2, 4, "clickEffect", setVis(4, 2, "visible"), ` dur="500"`) +
        // 2. Emphasis / Spin clockwise on shape 2
        par(5, "emph", 8, 0, "withEffect", animRot(6, 2, 1500, 0, 360), ` dur="1500"`) +
        // 3. Exit / Fade on shape 2
        par(7, "exit", 10, 0, "afterEffect", setVis(8, 2, "hidden"), ` dur="500"`) +
        // 4. Entrance / Wipe (up) on shape 3, delayed
        par(9, "entr", 10, 8, "clickEffect", setVis(10, 3, "visible"), ` dur="500" delay="500"`) +
        // 5. Motion Path / Arc on shape 3
        par(11, "path", 2, 0, "withEffect",
          animMotion(12, 3, 2000, "M 0 0 C 0 -0.15 0.15 -0.25 0.3 -0.25 E"),
          ` dur="2000"`) +
        `</p:childTnLst></p:cTn>` +
        `</p:seq></p:childTnLst></p:cTn></p:par></p:tnLst>` +
        `</p:timing>`;
      const patched = slideXml.replace("</p:sld>", `${insert}</p:sld>`);
      zip.file(slidePath, patched);
    }
  );
}
