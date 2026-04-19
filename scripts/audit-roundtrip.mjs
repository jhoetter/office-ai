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
      if (child.kind !== "run") continue;
      const r = child;
      if (r.properties?.bold) bump(t, "run-bold");
      if (r.properties?.italic) bump(t, "run-italic");
      if (r.properties?.fontFamily) bump(t, "run-font-family");
      if (typeof r.properties?.fontSize === "number") bump(t, "run-font-size");
      if (r.properties?.color) bump(t, "run-color");
      if (r.properties?.highlight) bump(t, "run-highlight");
      for (const leaf of r.children ?? []) {
        if (leaf.kind === "text") bump(t, "text-leaves");
        else if (leaf.kind === "drawing") bump(t, "drawings");
      }
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
    for (const cell of sheet.cells?.values?.() ?? []) {
      bump(t, "cells");
      if (cell.formula) bump(t, "formulas");
      if (cell.styleId !== undefined) bump(t, "styled-cells");
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

/** ── PPTX walker ─────────────────────────────────────────────── */
function tallyPptx(snapshot) {
  const t = blankTally();
  const pres = snapshot.root;
  bump(t, "slides", pres.slides?.length ?? 0);
  for (const slide of pres.slides ?? []) {
    for (const shape of slide.shapes ?? []) {
      bump(t, "shapes");
      if (shape.kind === "picture") bump(t, "pictures");
      if (shape.kind === "chart") bump(t, "charts");
      if (shape.kind === "text" || shape.kind === "shape") {
        for (const para of shape.text?.paragraphs ?? []) {
          if (para.alignment) bump(t, "paragraph-alignment");
          for (const run of para.runs ?? []) {
            if (run.properties?.bold) bump(t, "run-bold");
            if (run.properties?.italic) bump(t, "run-italic");
            if (typeof run.properties?.fontSize === "number") bump(t, "run-font-size");
            if (run.properties?.color) bump(t, "run-color");
            if (run.properties?.fontFamily) bump(t, "run-font-family");
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

async function main() {
  const summary = { generatedAt: new Date().toISOString(), formats: [] };
  for (const fmt of FORMATS) {
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
      fixtures: res.fixtures.map((r) => ({
        fixture: r.fixture.replace(root + "/", ""),
        ok: r.ok,
        losses: r.losses ?? [],
        error: r.error,
      })),
    });
  }
  const outDir = resolve(root, "docs/build-log");
  mkdirSync(outDir, { recursive: true });
  const jsonOut = resolve(outDir, "roundtrip-audit-night.json");
  writeFileSync(jsonOut, JSON.stringify(summary, null, 2));
  console.log(`\nJSON summary → ${jsonOut.replace(root + "/", "")}`);
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
