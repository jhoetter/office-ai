#!/usr/bin/env node
/**
 * Diagnostic for the masterthesis fixture: dumps the body block
 * structure, identifies opaque carriers, and reports what the
 * Phase 1 page chunker produces *without* live measurement (i.e. the
 * fallback used on first paint and in tests).
 *
 * Run: node scripts/inspect-masterthesis.mjs
 */
import { readFile } from "node:fs/promises";
import { resolve } from "node:path";
import { DocxAgent, chunkIntoPages } from "@officeai/docx";

const FIXTURE = resolve(
  process.cwd(),
  "..",
  "Masterthesis_Rohfassung_Verification_Asymmetry Kopie.docx"
);

const buf = await readFile(FIXTURE);
const agent = await DocxAgent.fromBuffer(buf);
const snap = agent.getSnapshot();
const body = snap.root.body;

console.log(`=== Body has ${body.length} blocks ===`);

const counts = {};
for (const b of body) {
  counts[b.kind] = (counts[b.kind] ?? 0) + 1;
  if (b.kind === "opaque-block") {
    const tag = b.raw?.tag ?? "?";
    counts[`opaque:${tag}`] = (counts[`opaque:${tag}`] ?? 0) + 1;
  }
}
console.log("Kind histogram:", counts);

console.log("\n=== First 80 blocks (compact) ===");
for (let i = 0; i < Math.min(80, body.length); i++) {
  const b = body[i];
  let label = b.kind;
  if (b.kind === "paragraph") {
    const text = paragraphText(b).slice(0, 60).replace(/\s+/g, " ");
    label = `p[${b.properties.styleId ?? ""}]${
      b.properties.pageBreakBefore ? "[pbb]" : ""
    }${b.properties.keepNext ? "[keep]" : ""}${
      hasPageBreakChild(b) ? "[hardpb]" : ""
    }${hasLastRendered(b) ? "[last]" : ""} "${text}"`;
  } else if (b.kind === "table") {
    label = `table rows=${b.rows.length}`;
  } else if (b.kind === "section-break") {
    label = `sectPr w:${b.properties.pgSz?.w ?? "?"} h:${b.properties.pgSz?.h ?? "?"}`;
  } else if (b.kind === "wrapper-marker") {
    label = `wrapper-${b.side} ${b.wrapperRaw.tag} id=${b.wrapperId}`;
  } else if (b.kind === "opaque-block") {
    const childCount = b.children?.length ?? 0;
    label = `opaque:${b.raw.tag} children=${childCount}`;
    if (childCount > 0) {
      const firstFew = b.children
        .slice(0, 3)
        .map((c) => (c.kind === "paragraph" ? paragraphText(c).slice(0, 30) : c.kind))
        .join(" | ");
      label += ` first="${firstFew}"`;
    }
  }
  console.log(`  ${String(i).padStart(3)}: ${label}`);
}

console.log("\n=== Pagination (no measure) ===");
const chunks = chunkIntoPages(snap);
console.log(`Total pages: ${chunks.length}`);
for (let i = 0; i < Math.min(10, chunks.length); i++) {
  const c = chunks[i];
  console.log(
    `  page ${c.pageNumber}: blocks ${c.startBlock}..${c.endBlock} (n=${
      c.endBlock - c.startBlock
    }) sect=${c.sectionIndex}`
  );
}

console.log("\n=== Detail: block 62 (quadrant table) ===");
const t = body[62];
if (t && t.kind === "table") {
  console.log("table.tblPr:", JSON.stringify(t.tblPr ?? null, null, 2));
  console.log("table.tableProperties:", JSON.stringify(t.tableProperties ?? null, null, 2));
  console.log("table.properties:", JSON.stringify(t.properties ?? null, null, 2));
  console.log("table.gridCols:", JSON.stringify(t.gridCols ?? null, null, 2));
  console.log("Top-level keys:", Object.keys(t));
  if (t.rows[0]?.cells[0]) {
    const cell = t.rows[0].cells[0];
    console.log("\ncell0,0 keys:", Object.keys(cell));
    console.log("cell0,0.properties:", JSON.stringify(cell.properties ?? null, null, 2));
  }
}

console.log("\n=== Detail: block 25 (TOC SDT) ===");
const sdt = body[25];
if (sdt && sdt.kind === "opaque-block") {
  console.log("tag:", sdt.raw.tag);
  console.log("children count:", sdt.children?.length ?? 0);
  console.log("first 3 children kinds:", (sdt.children ?? []).slice(0, 3).map((c) => c.kind));
  if (sdt.children?.[0]?.kind === "paragraph") {
    console.log("first child styleId:", sdt.children[0].properties.styleId);
    console.log("first child text:", paragraphText(sdt.children[0]));
  }
}

console.log("\n=== Looking for tables and opaque-blocks ===");
for (let i = 0; i < body.length; i++) {
  const b = body[i];
  if (b.kind === "table" || (b.kind === "opaque-block" && b.children && b.children.length > 0)) {
    if (b.kind === "table") {
      const sampleCellText = b.rows[0]?.cells[0]?.body
        ?.map((bb) => (bb.kind === "paragraph" ? paragraphText(bb) : ""))
        .join(" ")
        .slice(0, 40);
      console.log(
        `  block ${i}: table rows=${b.rows.length} cols~${b.rows[0]?.cells.length ?? 0} first-cell="${sampleCellText}"`
      );
    } else {
      console.log(
        `  block ${i}: opaque ${b.raw.tag} children=${b.children.length}`
      );
    }
  }
}

function paragraphText(p) {
  let s = "";
  for (const inline of p.children) {
    if (inline.kind === "run") {
      for (const c of inline.children) {
        if (c.kind === "text") s += c.text;
        else if (c.kind === "tab") s += "\t";
      }
    } else if (inline.kind === "hyperlink") {
      for (const r of inline.children) {
        for (const c of r.children) {
          if (c.kind === "text") s += c.text;
        }
      }
    }
  }
  return s;
}

function hasPageBreakChild(p) {
  for (const inline of p.children) {
    if (inline.kind !== "run") continue;
    for (const c of inline.children) if (c.kind === "page-break") return true;
  }
  return false;
}

function hasLastRendered(p) {
  for (const inline of p.children) {
    if (inline.kind !== "run") continue;
    for (const c of inline.children) if (c.kind === "last-rendered-page-break") return true;
  }
  return false;
}
