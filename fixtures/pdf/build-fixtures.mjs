#!/usr/bin/env node
/**
 * Synthetic PDF fixture builder.
 *
 * Each fixture exercises one slice of the PDF feature surface so the
 * roundtrip + agent CLI suites can lock down behaviour against a stable
 * corpus. Generation is intentionally deterministic — every fixture
 * pins its `creationDate` / `modificationDate` to a fixed ISO date so
 * fixture rebuilds produce byte-identical bytes (for `git diff`-able
 * regeneration).
 *
 * Run: `node fixtures/pdf/build-fixtures.mjs`
 *
 * The script is idempotent: missing fixtures are generated, existing
 * ones are overwritten only when their bytes change. Keeps repo
 * pollution down and lets vitest specs treat the fixtures as
 * checked-in source-of-truth artifacts.
 */
import { mkdir, readFile, writeFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const HERE = dirname(fileURLToPath(import.meta.url));

// `pdf-lib` is hoisted into individual package node_modules under
// pnpm; resolve it relative to one of the consumers so this script
// runs from the workspace root without requiring fixtures/pdf to
// declare its own dependency.
const require = createRequire(
  join(HERE, "..", "..", "packages", "pdf-edit", "package.json"),
);
const {
  PageSizes,
  PDFArray,
  PDFDocument,
  PDFHexString,
  PDFName,
  PDFNumber,
  PDFString,
  StandardFonts,
  degrees,
  rgb,
} = require("pdf-lib");

// Pinned timestamp so rebuilds produce byte-stable fixtures.
const FIXED_DATE = new Date(Date.UTC(2026, 3, 20, 0, 0, 0));

/**
 * Pin every per-document timestamp so fixtures are byte-stable across
 * rebuilds. pdf-lib stamps `CreationDate` / `ModDate` on save so we
 * have to re-set them after constructing the document.
 */
function freezeDates(pdf) {
  pdf.setCreationDate(FIXED_DATE);
  pdf.setModificationDate(FIXED_DATE);
}

/**
 * Save a fixture only if its bytes differ from the current on-disk
 * copy. Returns one of `"created" | "updated" | "unchanged"` for the
 * idempotency log.
 */
async function writeIfChanged(name, bytes) {
  const path = join(HERE, name);
  const buf = Buffer.from(bytes);
  if (!existsSync(path)) {
    await mkdir(dirname(path), { recursive: true });
    await writeFile(path, buf);
    return "created";
  }
  const current = await readFile(path);
  if (current.equals(buf)) return "unchanged";
  await writeFile(path, buf);
  return "updated";
}

// ── Fixture builders ─────────────────────────────────────────────────────

async function buildSimpleText1Page() {
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const page = pdf.addPage([612, 792]);
  page.drawText("Single-page fixture for the PDF viewer roundtrip suite.", {
    x: 50,
    y: 720,
    size: 14,
    font,
    color: rgb(0, 0, 0),
  });
  page.drawText("Body line two — exercises plain-text projection.", {
    x: 50,
    y: 696,
    size: 12,
    font,
  });
  freezeDates(pdf);
  return pdf.save();
}

async function buildSimpleText3Page() {
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  for (let i = 1; i <= 3; i++) {
    const page = pdf.addPage([612, 792]);
    page.drawText(`Page ${i} of 3 — synthetic body text.`, {
      x: 50,
      y: 720,
      size: 18,
      font,
    });
    page.drawText(
      `Lorem ipsum dolor sit amet, page ${i} marker token PAGE_${i}.`,
      { x: 50, y: 690, size: 11, font },
    );
  }
  freezeDates(pdf);
  return pdf.save();
}

async function buildMetadataRich() {
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const page = pdf.addPage([612, 792]);
  page.drawText("Metadata fixture — every Info dict field is set.", {
    x: 50,
    y: 720,
    size: 14,
    font,
  });
  pdf.setTitle("Metadata-Rich Fixture");
  pdf.setAuthor("Office AI Night Shift");
  pdf.setSubject("Roundtrip metadata coverage");
  pdf.setKeywords(["pdf", "fixture", "metadata", "roundtrip"]);
  pdf.setCreator("officeai/build-fixtures");
  pdf.setProducer("officeai/pdf-lib");
  freezeDates(pdf);
  return pdf.save();
}

async function buildRotatedPages() {
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const angles = [0, 90, 180, 270];
  for (const angle of angles) {
    const page = pdf.addPage([612, 792]);
    page.drawText(`Rotation ${angle}°`, { x: 60, y: 720, size: 28, font });
    page.setRotation(degrees(angle));
  }
  freezeDates(pdf);
  return pdf.save();
}

/**
 * pdf-lib has no high-level outline API. We hand-roll the outline tree
 * using PDF object refs so the generated `/Outlines` survives a
 * pdf-lib roundtrip and shows up under PDF.js's `getOutline()`.
 */
async function buildWithOutline() {
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const titles = ["Chapter 1 — Introduction", "Chapter 2 — Body", "Chapter 3 — Conclusion"];
  const pageRefs = titles.map((title) => {
    const page = pdf.addPage([612, 792]);
    page.drawText(title, { x: 50, y: 720, size: 22, font });
    page.drawText("Outline anchor body text.", { x: 50, y: 690, size: 12, font });
    return page.ref;
  });

  const ctx = pdf.context;
  const outlinesRef = ctx.nextRef();
  const childRefs = titles.map(() => ctx.nextRef());
  const subRefs = [ctx.nextRef(), ctx.nextRef()];

  const dest = (pageRef) => ctx.obj([pageRef, PDFName.of("Fit")]);

  // Top-level outlines dictionary
  ctx.assign(
    outlinesRef,
    ctx.obj({
      Type: PDFName.of("Outlines"),
      First: childRefs[0],
      Last: childRefs[childRefs.length - 1],
      Count: childRefs.length + subRefs.length,
    }),
  );

  // Chapter entries
  for (let i = 0; i < childRefs.length; i++) {
    const entry = {
      Title: PDFHexString.fromText(titles[i]),
      Parent: outlinesRef,
      Dest: dest(pageRefs[i]),
    };
    if (i > 0) entry.Prev = childRefs[i - 1];
    if (i < childRefs.length - 1) entry.Next = childRefs[i + 1];
    if (i === 1) {
      // Two-level: chapter 2 has two children
      entry.First = subRefs[0];
      entry.Last = subRefs[1];
      entry.Count = 2;
    }
    ctx.assign(childRefs[i], ctx.obj(entry));
  }

  // Sub-entries under chapter 2
  ctx.assign(
    subRefs[0],
    ctx.obj({
      Title: PDFHexString.fromText("§2.1 First subsection"),
      Parent: childRefs[1],
      Next: subRefs[1],
      Dest: dest(pageRefs[1]),
    }),
  );
  ctx.assign(
    subRefs[1],
    ctx.obj({
      Title: PDFHexString.fromText("§2.2 Second subsection"),
      Parent: childRefs[1],
      Prev: subRefs[0],
      Dest: dest(pageRefs[1]),
    }),
  );

  const catalog = pdf.catalog;
  catalog.set(PDFName.of("Outlines"), outlinesRef);
  catalog.set(PDFName.of("PageMode"), PDFName.of("UseOutlines"));

  freezeDates(pdf);
  return pdf.save();
}

/**
 * AcroForm with all four widget kinds. Filled-in variant just toggles
 * values before save.
 */
async function buildAcroForm({ prefilled }) {
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const page = pdf.addPage([612, 792]);
  page.drawText(
    prefilled ? "Form fixture (pre-filled)" : "Form fixture (blank)",
    { x: 50, y: 740, size: 16, font },
  );

  const form = pdf.getForm();

  page.drawText("First name:", { x: 50, y: 700, size: 11, font });
  const tf = form.createTextField("first.name");
  tf.setText(prefilled ? "Ada" : "");
  tf.addToPage(page, { x: 140, y: 690, width: 200, height: 20 });

  page.drawText("Agree to terms:", { x: 50, y: 660, size: 11, font });
  const cb = form.createCheckBox("agree");
  if (prefilled) cb.check();
  cb.addToPage(page, { x: 160, y: 656, width: 16, height: 16 });

  page.drawText("Country:", { x: 50, y: 620, size: 11, font });
  const dd = form.createDropdown("country");
  dd.addOptions(["DE", "US", "FR", "JP"]);
  if (prefilled) dd.select("DE");
  dd.addToPage(page, { x: 140, y: 612, width: 120, height: 22 });

  page.drawText("Plan:", { x: 50, y: 580, size: 11, font });
  const rg = form.createRadioGroup("plan");
  rg.addOptionToPage("free", page, { x: 140, y: 575, width: 14, height: 14 });
  page.drawText("free", { x: 158, y: 578, size: 10, font });
  rg.addOptionToPage("pro", page, { x: 200, y: 575, width: 14, height: 14 });
  page.drawText("pro", { x: 218, y: 578, size: 10, font });
  if (prefilled) rg.select("pro");

  freezeDates(pdf);
  return pdf.save();
}

async function buildWithLinkAnnot() {
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const page = pdf.addPage([612, 792]);
  const linkText = "Visit cursor.com (URL link annotation)";
  page.drawText(linkText, { x: 50, y: 720, size: 14, font, color: rgb(0, 0, 1) });

  // URI Link annotation — hand-rolled because pdf-lib has no high-level
  // link API; the writer in @officeai/pdf-annotations does the same.
  const ctx = pdf.context;
  const linkDict = ctx.obj({
    Type: PDFName.of("Annot"),
    Subtype: PDFName.of("Link"),
    Rect: ctx.obj([
      PDFNumber.of(50),
      PDFNumber.of(715),
      PDFNumber.of(330),
      PDFNumber.of(735),
    ]),
    Border: ctx.obj([PDFNumber.of(0), PDFNumber.of(0), PDFNumber.of(1)]),
    A: ctx.obj({
      Type: PDFName.of("Action"),
      S: PDFName.of("URI"),
      URI: PDFString.of("https://cursor.com"),
    }),
  });
  const linkRef = ctx.register(linkDict);
  page.node.set(PDFName.of("Annots"), ctx.obj([linkRef]));

  freezeDates(pdf);
  return pdf.save();
}

async function buildWithHighlightAnnot() {
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const page = pdf.addPage([612, 792]);
  page.drawText("Highlight + sticky-note annotations exercise.", {
    x: 50,
    y: 720,
    size: 14,
    font,
  });

  const ctx = pdf.context;

  // Highlight rect over the title.
  const highlightRect = [50, 715, 360, 735];
  const quad = ctx.obj([
    PDFNumber.of(50),
    PDFNumber.of(735),
    PDFNumber.of(360),
    PDFNumber.of(735),
    PDFNumber.of(50),
    PDFNumber.of(715),
    PDFNumber.of(360),
    PDFNumber.of(715),
  ]);
  const highlightDict = ctx.obj({
    Type: PDFName.of("Annot"),
    Subtype: PDFName.of("Highlight"),
    Rect: ctx.obj(highlightRect.map((n) => PDFNumber.of(n))),
    QuadPoints: quad,
    C: ctx.obj([PDFNumber.of(1), PDFNumber.of(0.95), PDFNumber.of(0)]),
    T: PDFString.of("Office AI"),
    Contents: PDFHexString.fromText("Important phrase"),
    F: PDFNumber.of(4),
  });
  const highlightRef = ctx.register(highlightDict);

  // Sticky note next to the highlight.
  const noteDict = ctx.obj({
    Type: PDFName.of("Annot"),
    Subtype: PDFName.of("Text"),
    Rect: ctx.obj([
      PDFNumber.of(380),
      PDFNumber.of(715),
      PDFNumber.of(404),
      PDFNumber.of(739),
    ]),
    Open: false,
    Name: PDFName.of("Comment"),
    T: PDFString.of("Office AI"),
    Contents: PDFHexString.fromText("This is a sticky note for the highlighted phrase."),
    C: ctx.obj([PDFNumber.of(1), PDFNumber.of(0.85), PDFNumber.of(0.2)]),
    F: PDFNumber.of(4),
  });
  const noteRef = ctx.register(noteDict);

  page.node.set(PDFName.of("Annots"), ctx.obj([highlightRef, noteRef]));

  freezeDates(pdf);
  return pdf.save();
}

async function buildMultiSizePages() {
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const sizes = [
    { label: "Letter", size: PageSizes.Letter },
    { label: "A4", size: PageSizes.A4 },
    { label: "Legal", size: PageSizes.Legal },
    { label: "A3", size: PageSizes.A3 },
  ];
  for (const { label, size } of sizes) {
    const page = pdf.addPage(size);
    page.drawText(`${label}: ${Math.round(size[0])} × ${Math.round(size[1])} pt`, {
      x: 50,
      y: size[1] - 60,
      size: 18,
      font,
    });
  }
  freezeDates(pdf);
  return pdf.save();
}

async function buildLarge50Page() {
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  for (let i = 1; i <= 50; i++) {
    const page = pdf.addPage([612, 792]);
    page.drawText(`Page ${i} of 50`, { x: 50, y: 740, size: 16, font });
    page.drawText(
      `Token P${i.toString().padStart(2, "0")} — perf-budget filler text.`,
      { x: 50, y: 712, size: 11, font },
    );
  }
  freezeDates(pdf);
  return pdf.save();
}

/**
 * "Signed-then-modified" — a real signed PDF requires PKCS#12 keys
 * which are out of scope for synthetic fixtures. Instead we add a
 * /Sig form field with a visible widget so the parser detects
 * `signatureCount === 1`. The fixture intentionally has no actual
 * crypto signature (an "incremental update past signature" would
 * invalidate it anyway), which is exactly the "signed then modified"
 * scenario we want to test against.
 */
async function buildSignedThenModified() {
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const page = pdf.addPage([612, 792]);
  page.drawText("Document with a signature widget (unsigned).", {
    x: 50,
    y: 720,
    size: 14,
    font,
  });
  page.drawText("Signed by: ____________________", {
    x: 50,
    y: 660,
    size: 12,
    font,
  });

  const ctx = pdf.context;

  // Build the signature widget annotation + field dictionary.
  const widgetRect = ctx.obj([
    PDFNumber.of(170),
    PDFNumber.of(650),
    PDFNumber.of(380),
    PDFNumber.of(680),
  ]);
  const sigFieldRef = ctx.nextRef();
  const sigFieldDict = ctx.obj({
    FT: PDFName.of("Sig"),
    T: PDFString.of("signature.0"),
    Type: PDFName.of("Annot"),
    Subtype: PDFName.of("Widget"),
    F: PDFNumber.of(4),
    Rect: widgetRect,
    P: page.ref,
  });
  ctx.assign(sigFieldRef, sigFieldDict);

  // Wire the widget into the page's /Annots.
  const existingAnnots = page.node.get(PDFName.of("Annots"));
  if (existingAnnots instanceof PDFArray) {
    existingAnnots.push(sigFieldRef);
  } else {
    page.node.set(PDFName.of("Annots"), ctx.obj([sigFieldRef]));
  }

  // Register the AcroForm with SigFlags=3 (signatures-exist + append-only).
  const acroFormDict = ctx.obj({
    Fields: ctx.obj([sigFieldRef]),
    SigFlags: PDFNumber.of(3),
  });
  const acroFormRef = ctx.register(acroFormDict);
  pdf.catalog.set(PDFName.of("AcroForm"), acroFormRef);

  freezeDates(pdf);
  // Save without object streams so the AcroForm dictionary is plainly
  // visible in the byte stream — easier on the parser fallbacks.
  const saved = await pdf.save({ useObjectStreams: false });

  return saved;
}

// ── Driver ────────────────────────────────────────────────────────────────

const FIXTURES = [
  { name: "simple-text-1page.pdf", build: buildSimpleText1Page, maxBytes: 80_000 },
  { name: "simple-text-3page.pdf", build: buildSimpleText3Page, maxBytes: 80_000 },
  { name: "metadata-rich.pdf", build: buildMetadataRich, maxBytes: 80_000 },
  { name: "rotated-pages.pdf", build: buildRotatedPages, maxBytes: 80_000 },
  { name: "with-outline.pdf", build: buildWithOutline, maxBytes: 80_000 },
  { name: "acroform-basic.pdf", build: () => buildAcroForm({ prefilled: false }), maxBytes: 80_000 },
  { name: "acroform-prefilled.pdf", build: () => buildAcroForm({ prefilled: true }), maxBytes: 80_000 },
  { name: "with-link-annot.pdf", build: buildWithLinkAnnot, maxBytes: 80_000 },
  { name: "with-highlight-annot.pdf", build: buildWithHighlightAnnot, maxBytes: 80_000 },
  { name: "multi-size-pages.pdf", build: buildMultiSizePages, maxBytes: 80_000 },
  { name: "large-50page.pdf", build: buildLarge50Page, maxBytes: 320_000 },
  { name: "signed-then-modified.pdf", build: buildSignedThenModified, maxBytes: 80_000 },
];

async function main() {
  let created = 0;
  let updated = 0;
  let unchanged = 0;
  for (const fixture of FIXTURES) {
    const bytes = await fixture.build();
    if (bytes.length > fixture.maxBytes) {
      throw new Error(
        `${fixture.name}: ${bytes.length} bytes exceeds budget of ${fixture.maxBytes} bytes`,
      );
    }
    const status = await writeIfChanged(fixture.name, bytes);
    process.stdout.write(
      `  ${status === "unchanged" ? "·" : status === "created" ? "+" : "~"} ${fixture.name.padEnd(28)} ${bytes.length.toString().padStart(7)} bytes (${status})\n`,
    );
    if (status === "created") created++;
    else if (status === "updated") updated++;
    else unchanged++;
  }
  process.stdout.write(
    `\n${FIXTURES.length} fixtures: ${created} created, ${updated} updated, ${unchanged} unchanged.\n`,
  );
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
