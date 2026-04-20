#!/usr/bin/env node
/**
 * Round-trip attribute fidelity audit.
 *
 * For every fixture under `fixtures/{docx,xlsx,pptx}/...`, this
 * script:
 *
 *   1. parses the bytes into the typed snapshot via the
 *      format-specific agent
 *   2. exports the snapshot back to bytes
 *   3. parses the exported bytes again
 *   4. walks both snapshots in lockstep and counts how many of a
 *      curated set of formatting attributes survived the trip.
 *
 * The output is a human-readable per-format table plus a JSON
 * summary written to `docs/build-log/roundtrip-audit-night.json`
 * for downstream tooling. Exit code stays 0 even on diffs because
 * this is a *reporting* tool — the existing
 * `run-libreoffice-roundtrip.mjs` is the one that gates `make
 * verify`.
 *
 * What we check (today):
 *   - DOCX: paragraph count, paragraph alignment, run.bold/italic
 *           /font.family / font.size, page setup (cx/cy/margins),
 *           list numbering ids, hyperlinks
 *   - XLSX: cell value/formula presence, style.font/size/color,
 *           horizontal alignment, merged ranges, column widths,
 *           charts (presence by sheet)
 *   - PPTX: slide count, shape count per slide, run.bold/italic
 *           /size/color, picture count, chart count
 *
 * Adding more attributes is intentionally low-friction: each
 * format reducer returns a flat `Map<keyof Tally, number>` so a
 * new field is one new entry in the tally object plus one walker
 * line.
 */

import { existsSync, mkdirSync, readdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, "..");

const FORMATS = [
  {
    id: "docx",
    extension: ".docx",
    fixtureDirs: ["fixtures/docx/real-world"],
    agentEntry: "packages/docx/dist/index.js",
    agentName: "DocxAgent",
    tally: tallyDocx,
  },
  {
    id: "xlsx",
    extension: ".xlsx",
    fixtureDirs: ["fixtures/xlsx/synthetic", "fixtures/xlsx/real-world"],
    agentEntry: "packages/xlsx/dist/index.js",
    agentName: "XlsxAgent",
    tally: tallyXlsx,
  },
  {
    id: "pptx",
    extension: ".pptx",
    fixtureDirs: ["fixtures/pptx/synthetic", "fixtures/pptx/real"],
    agentEntry: "packages/pptx/dist/index.js",
    agentName: "PptxAgent",
    tally: tallyPptx,
  },
  {
    id: "pdf",
    extension: ".pdf",
    fixtureDirs: ["fixtures/pdf"],
    agentEntry: "packages/pdf/dist/index.js",
    agentName: "PdfAgent",
    tally: tallyPdf,
    // PDFs use a bespoke audit (parse + valid bytes + markdown sanity)
    // instead of the generic tally diff because the serializer
    // intentionally rewrites the byte stream and only guarantees
    // attribute fidelity for the page-rotation / reorder / metadata
    // subset (see packages/pdf/src/serializer/serialize.ts).
    audit: auditPdfFixture,
  },
];

async function loadAgent(fmt) {
  const distEntry = resolve(root, fmt.agentEntry);
  if (!existsSync(distEntry)) {
    throw new Error(
      `Agent entry not found at ${distEntry}. Run \`pnpm --filter @officeai/${fmt.id} build\` first.`
    );
  }
  const mod = await import(pathToFileURL(distEntry).href);
  const Agent = mod[fmt.agentName];
  if (!Agent) throw new Error(`Module ${distEntry} does not export ${fmt.agentName}.`);
  return Agent;
}

function listFixtures(fmt) {
  const out = [];
  for (const rel of fmt.fixtureDirs) {
    const dir = resolve(root, rel);
    if (!existsSync(dir)) continue;
    let entries;
    try {
      entries = readdirSync(dir);
    } catch {
      continue;
    }
    for (const f of entries.sort()) {
      if (f.toLowerCase().endsWith(fmt.extension)) out.push(join(dir, f));
    }
  }
  return out;
}

/** ── DOCX walker ─────────────────────────────────────────────── */
function tallyDocx(snapshot) {
  const t = blankTally();
  const root = snapshot.root;
  for (const block of root.body ?? []) {
    if (block.kind !== "paragraph") continue;
    bump(t, "paragraphs");
    const p = block;
    if (p.properties?.alignment) bump(t, "paragraph-alignment");
    if (p.properties?.numPr?.numId !== undefined) bump(t, "paragraph-list");
    for (const child of p.children ?? []) {
      if (child.kind === "comment-reference") bump(t, "comment-references");
      if (child.kind !== "run") continue;
      const r = child;
      if (r.properties?.bold) bump(t, "run-bold");
      if (r.properties?.italic) bump(t, "run-italic");
      if (r.properties?.fontFamily) bump(t, "run-font-family");
      // Multi-script + theme font slots (P1 expansion).
      if (r.properties?.fontFamilyHAnsi) bump(t, "run-font-hAnsi");
      if (r.properties?.fontFamilyEastAsia) bump(t, "run-font-eastAsia");
      if (r.properties?.fontFamilyComplexScript) bump(t, "run-font-cs");
      if (r.properties?.fontFamilyAsciiTheme) bump(t, "run-font-asciiTheme");
      if (r.properties?.fontFamilyHAnsiTheme) bump(t, "run-font-hAnsiTheme");
      if (r.properties?.fontFamilyEastAsiaTheme) bump(t, "run-font-eastAsiaTheme");
      if (r.properties?.fontFamilyComplexScriptTheme) bump(t, "run-font-csTheme");
      if (typeof r.properties?.fontSize === "number") bump(t, "run-font-size");
      if (r.properties?.color) bump(t, "run-color");
      if (r.properties?.highlight) bump(t, "run-highlight");
      for (const leaf of r.children ?? []) {
        if (leaf.kind === "text") bump(t, "text-leaves");
        else if (leaf.kind === "drawing") bump(t, "drawings");
        else if (leaf.kind === "page-number-field") bump(t, "page-number-fields");
      }
    }
  }
  // Header/footer parts and any tables nested inside them.
  for (const part of root.headersAndFooters ?? []) {
    bump(t, "header-footer-parts");
    for (const block of part.body ?? []) {
      if (block.kind === "table") bump(t, "header-footer-tables");
    }
  }
  // Page setup lives on the last sectPr; we hash it as a single
  // string so a missing margin or pgSz still gets flagged.
  const sect = root.sections?.[root.sections.length - 1];
  if (sect?.pgSz) {
    bump(t, "page-size");
    if (sect.pgSz.orient === "landscape") bump(t, "page-landscape");
  }
  if (sect?.pgMar) bump(t, "page-margins");
  return t;
}

/** ── XLSX walker ─────────────────────────────────────────────── */
function tallyXlsx(snapshot) {
  const t = blankTally();
  const wb = snapshot.root;
  for (const sheet of wb.sheets ?? []) {
    bump(t, "sheets");
    if (sheet.charts && sheet.charts.length > 0) bump(t, "charts", sheet.charts.length);
    if (sheet.images && sheet.images.length > 0) bump(t, "images", sheet.images.length);
    if (sheet.merges) bump(t, "merges", sheet.merges.length);
    if (sheet.cols) {
      for (const col of sheet.cols) {
        if (typeof col.width === "number") bump(t, "col-widths");
      }
    }
    // Opaque sheet parts (P0 + P2 round-trip work). One bump per
    // non-empty blob so a fixture losing any of them shows up as a
    // negative delta in the JSON summary.
    if (sheet.hyperlinksXml) bump(t, "hyperlinks-blocks");
    if (sheet.tablePartsXml) bump(t, "table-parts-blocks");
    if (sheet.colsXml) bump(t, "cols-blocks");
    if (sheet.sheetViewsXml) bump(t, "sheet-views-blocks");
    if (sheet.sheetProtectionXml) bump(t, "sheet-protection-blocks");
    if (sheet.pageMarginsXml) bump(t, "page-margins-blocks");
    if (sheet.pageSetupXml) bump(t, "page-setup-blocks");
    if (sheet.headerFooterXml) bump(t, "header-footer-blocks");
    if (sheet.legacyDrawingXml) bump(t, "legacy-drawing-blocks");
    if (sheet.ignoredErrorsXml) bump(t, "ignored-errors-blocks");
    for (const cell of sheet.cells?.values?.() ?? []) {
      bump(t, "cells");
      if (cell.formula) {
        bump(t, "formulas");
        // Shared / array formula encoding (P2). Counts both master
        // and follower cells so a regression that demotes the group
        // to per-cell formulas surfaces as a negative delta.
        if (cell.formula.kind === "shared") bump(t, "formulas-shared");
        if (cell.formula.kind === "array") bump(t, "formulas-array");
        if (cell.formula.isMaster) bump(t, "formulas-shared-master");
      }
      if (cell.styleId !== undefined) bump(t, "styled-cells");
    }
    // Typed conditional formats and comment rich-text blobs (P2).
    if (sheet.conditionalFormats) bump(t, "cond-formats", sheet.conditionalFormats.length);
    if (sheet.opaqueConditionalFormats) bump(t, "cond-formats-opaque", sheet.opaqueConditionalFormats.length);
    if (sheet.comments) {
      for (const c of sheet.comments) {
        bump(t, "comments");
        if (c.textXml) bump(t, "comments-rich-text");
      }
    }
  }
  // Style table
  const styles = wb.styles;
  if (styles) {
    if (styles.fonts) bump(t, "fonts", styles.fonts.length);
    if (styles.fills) bump(t, "fills", styles.fills.length);
    if (styles.numFmts) bump(t, "num-fmts", styles.numFmts.length);
  }
  return t;
}

/** ── PDF walker ──────────────────────────────────────────────── */
function tallyPdf(snapshot) {
  const t = blankTally();
  const root = snapshot.root;
  bump(t, "pages", root.pages?.length ?? 0);
  bump(t, "outline-entries", flattenOutlineCount(root.outline ?? []));
  bump(t, "annotations", root.annotations?.length ?? 0);
  bump(t, "form-fields", root.formFields?.length ?? 0);
  bump(t, "signatures", root.signatureCount ?? 0);
  for (const p of root.pages ?? []) {
    if (p.rotation && p.rotation !== 0) bump(t, "rotated-pages");
    if (p.hasTextLayer) bump(t, "pages-with-text");
  }
  const md = root.metadata ?? {};
  for (const k of ["title", "author", "subject", "keywords", "creator", "producer"]) {
    if (md[k] !== undefined) bump(t, `meta-${k}`);
  }
  return t;
}

function flattenOutlineCount(nodes) {
  let n = 0;
  for (const node of nodes) {
    n += 1 + flattenOutlineCount(node.children ?? []);
  }
  return n;
}

/**
 * PDF-specific per-fixture auditor. Three gates:
 *   (a) PdfAgent.fromBuffer succeeds
 *   (b) agent.exportFile() returns bytes that start with %PDF-
 *   (c) agent.toMarkdown() emits at least one page heading (and the
 *       title when one is present in the Info dict)
 */
async function auditPdfFixture(Agent, path) {
  const buf = readFileSync(path);
  const before = await Agent.fromBuffer(buf);
  const beforeTally = tallyPdf(before.getSnapshot());
  const exported = Buffer.from(await before.exportFile());
  const isPdf = exported.length > 5 && exported.slice(0, 5).toString("ascii") === "%PDF-";
  const after = await Agent.fromBuffer(exported);
  const afterTally = tallyPdf(after.getSnapshot());
  const md = before.toMarkdown();
  const hasPageHeading = /^### Page \d+/m.test(md);
  const expectedTitle = before.getSnapshot().root.metadata?.title;
  const hasTitleHeading =
    expectedTitle === undefined ? true : md.startsWith(`# ${expectedTitle}`);
  const diffs = diffTallies(beforeTally, afterTally);
  const losses = diffs.filter((d) => d.delta < 0);
  return {
    fixture: path,
    ok: isPdf && hasPageHeading && hasTitleHeading && losses.length === 0,
    losses,
    gains: diffs.filter((d) => d.delta > 0),
    diffs,
    pdf: {
      bytes: exported.byteLength,
      isPdfMagic: isPdf,
      hasPageHeading,
      hasTitleHeading,
      pages: before.getSnapshot().root.pages.length,
      markdownBytes: md.length,
    },
  };
}

/** ── PPTX walker ─────────────────────────────────────────────── */
function tallyPptx(snapshot) {
  const t = blankTally();
  const pres = snapshot.root;
  bump(t, "slides", pres.slides?.length ?? 0);
  for (const slide of pres.slides ?? []) {
    if (slide.transition) {
      bump(t, "transitions");
      if (slide.transition.raw) bump(t, "transitions-with-raw");
    }
    for (const shape of slide.shapes ?? []) {
      bump(t, "shapes");
      if (shape.kind === "picture") bump(t, "pictures");
      if (shape.kind === "chart") bump(t, "charts");
      if (shape.kind === "connector") {
        bump(t, "connectors");
        if (shape.stroke?.colorTheme) bump(t, "connectors-theme-color");
        if (shape.stroke?.dash && shape.stroke.dash !== "solid") bump(t, "connectors-dashed");
      }
      if (shape.kind === "text" || shape.kind === "shape") {
        for (const para of shape.text?.paragraphs ?? []) {
          if (para.alignment) bump(t, "paragraph-alignment");
          for (const run of para.runs ?? []) {
            if (run.properties?.bold) bump(t, "run-bold");
            if (run.properties?.italic) bump(t, "run-italic");
            if (typeof run.properties?.fontSize === "number") bump(t, "run-font-size");
            if (run.properties?.color) bump(t, "run-color");
            if (run.properties?.fontFamily) bump(t, "run-font-family");
            // Multi-script + theme font slots (P1 expansion).
            if (run.properties?.fontFamilyEastAsia) bump(t, "run-font-eastAsia");
            if (run.properties?.fontFamilyComplexScript) bump(t, "run-font-cs");
            if (run.properties?.fontFamilySymbol) bump(t, "run-font-symbol");
            if (run.properties?.fontFamilyLatinTheme) bump(t, "run-font-latinTheme");
            if (run.properties?.fontFamilyEastAsiaTheme) bump(t, "run-font-eastAsiaTheme");
            if (run.properties?.fontFamilyComplexScriptTheme) bump(t, "run-font-csTheme");
          }
        }
      }
    }
  }
  return t;
}

function blankTally() {
  return new Map();
}

function bump(t, key, delta = 1) {
  t.set(key, (t.get(key) ?? 0) + delta);
}

function diffTallies(before, after) {
  const keys = new Set([...before.keys(), ...after.keys()]);
  const rows = [];
  for (const k of [...keys].sort()) {
    const a = before.get(k) ?? 0;
    const b = after.get(k) ?? 0;
    rows.push({ attribute: k, before: a, after: b, delta: b - a });
  }
  return rows;
}

async function auditFormat(fmt) {
  const Agent = await loadAgent(fmt);
  const fixtures = listFixtures(fmt);
  if (fixtures.length === 0) {
    return { id: fmt.id, fixtures: [], skipped: true };
  }
  const results = [];
  for (const path of fixtures) {
    let row;
    try {
      if (typeof fmt.audit === "function") {
        row = await fmt.audit(Agent, path);
      } else {
        const buf = readFileSync(path);
        const agentBefore = await Agent.fromBuffer(buf);
        const before = fmt.tally(agentBefore.getSnapshot());
        const exported = Buffer.from(await agentBefore.exportFile());
        const agentAfter = await Agent.fromBuffer(exported);
        const after = fmt.tally(agentAfter.getSnapshot());
        const diffs = diffTallies(before, after);
        const losses = diffs.filter((d) => d.delta < 0);
        const gains = diffs.filter((d) => d.delta > 0);
        row = {
          fixture: path,
          ok: losses.length === 0,
          losses,
          gains,
          diffs,
        };
      }
    } catch (err) {
      row = { fixture: path, ok: false, error: String(err?.message ?? err) };
    }
    results.push(row);
  }
  return { id: fmt.id, fixtures: results };
}

function fmtRow(r) {
  if (r.error) return `  ✗ ${r.fixture.split("/").slice(-2).join("/")}: ERROR ${r.error}`;
  const name = r.fixture.split("/").slice(-2).join("/");
  if (r.ok && r.gains.length === 0) {
    const total = r.diffs.reduce((a, d) => a + d.before, 0);
    return `  ✓ ${name} (${total} attrs, exact match)`;
  }
  const lossSummary = r.losses.map((l) => `${l.attribute} ${l.before}→${l.after}`).join(", ");
  const gainSummary = r.gains.map((g) => `${g.attribute} ${g.before}→${g.after}`).join(", ");
  if (r.ok && r.gains.length > 0) {
    return `  ⚠ ${name}: gained ${gainSummary}`;
  }
  if (lossSummary && gainSummary) {
    return `  ✗ ${name}: lost ${lossSummary}; gained ${gainSummary}`;
  }
  return `  ✗ ${name}: lost ${lossSummary}`;
}

function parseArgs(argv) {
  const args = { product: null };
  for (let i = 0; i < argv.length; i++) {
    if (argv[i] === "--product" && i + 1 < argv.length) {
      args.product = argv[i + 1];
      i++;
    } else if (argv[i].startsWith("--product=")) {
      args.product = argv[i].slice("--product=".length);
    }
  }
  return args;
}

/**
 * Per-fixture summary row in the JSON envelope. Pulls the PDF-specific
 * fields out of the audit row when present so consumers don't have to
 * reach into `losses`/`gains`.
 */
function summarizeRow(r) {
  return {
    fixture: r.fixture.replace(root + "/", ""),
    ok: r.ok,
    losses: r.losses ?? [],
    ...(r.pdf ? { pdf: r.pdf } : {}),
    ...(r.error ? { error: r.error } : {}),
  };
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const filtered = args.product ? FORMATS.filter((f) => f.id === args.product) : FORMATS;
  if (filtered.length === 0) {
    console.error(
      `--product "${args.product}" not recognised. Known products: ${FORMATS.map((f) => f.id).join(", ")}`,
    );
    process.exit(2);
  }

  // Merge mode: when invoked with --product, preserve previously
  // recorded entries for the other formats so the JSON stays a
  // complete snapshot. Without --product (the default 'audit
  // everything' invocation) the file is rewritten from scratch.
  const outDir = resolve(root, "docs/build-log");
  mkdirSync(outDir, { recursive: true });
  const jsonOut = resolve(outDir, "roundtrip-audit-night.json");
  let summary = { generatedAt: new Date().toISOString(), formats: [] };
  if (args.product && existsSync(jsonOut)) {
    try {
      const prev = JSON.parse(readFileSync(jsonOut, "utf8"));
      if (prev && Array.isArray(prev.formats)) {
        summary.formats = prev.formats.filter((f) => f.id !== args.product);
      }
    } catch {
      // Falling through is safe — we'll just rebuild a fresh file.
    }
  }

  for (const fmt of filtered) {
    process.stdout.write(`\n=== ${fmt.id.toUpperCase()} ===\n`);
    let res;
    try {
      res = await auditFormat(fmt);
    } catch (err) {
      console.error(`  cannot audit ${fmt.id}: ${err.message}`);
      summary.formats.push({ id: fmt.id, error: err.message });
      continue;
    }
    if (res.skipped) {
      console.log("  (no fixtures present, skipped)");
      summary.formats.push({ id: fmt.id, skipped: true });
      continue;
    }
    for (const r of res.fixtures) {
      console.log(fmtRow(r));
    }
    const ok = res.fixtures.filter((r) => r.ok).length;
    const total = res.fixtures.length;
    console.log(`  → ${ok}/${total} fixtures clean`);
    summary.formats.push({
      id: fmt.id,
      total,
      ok,
      fixtures: res.fixtures.map(summarizeRow),
    });
  }
  writeFileSync(jsonOut, JSON.stringify(summary, null, 2));
  console.log(`\nJSON summary → ${jsonOut.replace(root + "/", "")}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
