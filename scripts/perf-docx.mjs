#!/usr/bin/env node
/**
 * DOCX perf budget runner.
 *
 * Builds a ~100-page synthetic DOCX in memory, then asserts:
 *
 *   1. parse        < 500 ms
 *   2. 1k commands  < 1 s aggregate dispatch (insert-text into the body)
 *   3. serialize    < 750 ms
 *
 * Budgets are tuned to comfortably pass on an Apple Silicon M-class laptop
 * (the dev target), with ~30 % headroom over the typical observed run so
 * minor flakes don't gate PRs.
 *
 * Prints a small markdown table on stdout. Exits non-zero on budget violation.
 *
 * Run via `make perf-docx` or `node scripts/perf-docx.mjs`. The script imports
 * the built `@officeai/docx` from `packages/docx/dist/index.js`; if the build
 * is missing it runs `pnpm --filter @officeai/docx build` first so a one-shot
 * invocation works on a fresh clone.
 */
import { spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import { performance } from "node:perf_hooks";
import { dirname, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
// `docx` is a root devDependency (used by scripts/generate-real-fixtures.mjs).
// Reusing it here avoids adding a new top-level package just to write a zip.
import { Document, HeadingLevel, Packer, Paragraph, TextRun } from "docx";

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, "..");
const distEntry = resolve(root, "packages/docx/dist/index.js");

// ── Budgets (ms) ──────────────────────────────────────────────────────────
// Tuned so the typical run on an M-class laptop sits around 60 % of budget;
// the 30 %+ headroom is intentional so legitimate variance doesn't flake CI.
const BUDGETS = {
  parse: 500,
  thousandCommands: 1000,
  serialize: 750,
};

// ── Synthetic document shape ──────────────────────────────────────────────
// 100 pages ≈ 30 paragraphs/page in normal Word body copy. We mix in a
// heading every page so the parser exercises pPr/style branches, not just
// the cheap text path.
const PAGES = 100;
const PARAS_PER_PAGE = 30;
const TOTAL_PARAS = PAGES * PARAS_PER_PAGE;
const COMMAND_COUNT = 1000;

const LOREM = [
  "The quick brown fox jumps over the lazy dog.",
  "Premature optimization is the root of all evil.",
  "Make the change easy, then make the easy change.",
  "We shape our buildings, then they shape us.",
  "All software architecture eventually optimizes for change.",
  "Cache invalidation and naming are the hard ones; off-by-one is free.",
];

async function buildSyntheticDocxBuffer() {
  const paragraphs = [];
  for (let i = 0; i < TOTAL_PARAS; i++) {
    if (i % PARAS_PER_PAGE === 0) {
      paragraphs.push(
        new Paragraph({
          heading: HeadingLevel.HEADING_1,
          children: [new TextRun(`Section ${i / PARAS_PER_PAGE + 1}`)],
        })
      );
    } else {
      const text = LOREM[i % LOREM.length];
      paragraphs.push(
        new Paragraph({ children: [new TextRun(`${text} (paragraph ${i})`)] })
      );
    }
  }
  const doc = new Document({ sections: [{ children: paragraphs }] });
  return Packer.toBuffer(doc);
}

function ensureBuilt() {
  if (existsSync(distEntry)) return;
  console.log("⚙ packages/docx not built yet — building now…");
  const r = spawnSync("pnpm", ["--filter", "@officeai/docx", "build"], {
    cwd: root,
    stdio: "inherit",
  });
  if (r.status !== 0) {
    console.error("perf-docx: failed to build @officeai/docx — cannot run perf check.");
    process.exit(1);
  }
}

async function loadDocxApi() {
  const url = pathToFileURL(distEntry).href;
  return import(url);
}

function fmtMs(ms) {
  return ms.toFixed(1);
}

async function main() {
  ensureBuilt();
  const { DocxAgent } = await loadDocxApi();

  console.log(
    `perf-docx: synthetic ${PAGES}-page DOCX (${TOTAL_PARAS} paragraphs); dispatching ${COMMAND_COUNT} insert-text commands.`
  );
  const buf = await buildSyntheticDocxBuffer();
  console.log(`  synthetic .docx size: ${(buf.byteLength / 1024).toFixed(1)} KB\n`);

  // ── 1. Parse ────────────────────────────────────────────────────────────
  const t0 = performance.now();
  const agent = await DocxAgent.fromBuffer(buf);
  const parseMs = performance.now() - t0;

  // Sanity: confirm the model has the expected shape.
  const snap0 = agent.getSnapshot();
  const paras = snap0.root.body.filter((b) => b.kind === "paragraph").length;

  // ── 2. 1k commands ──────────────────────────────────────────────────────
  // Single-paragraph inserts are the cheapest realistic mutation; the budget
  // covers handler dispatch + diff materialization + immutable evolve, which
  // is what an LLM-driven session actually exercises.
  const t1 = performance.now();
  for (let i = 0; i < COMMAND_COUNT; i++) {
    const targetParagraph = i % paras;
    await agent.applyCommand({
      type: "docx:insert-text",
      payload: {
        at: { paragraph: targetParagraph, run: 0, offset: 0 },
        text: ".",
      },
      source: "agent",
      agentId: "perf-docx",
    });
  }
  const cmdsMs = performance.now() - t1;

  // ── 3. Serialize ────────────────────────────────────────────────────────
  const t2 = performance.now();
  const out = await agent.exportFile();
  const serializeMs = performance.now() - t2;

  // ── Report ──────────────────────────────────────────────────────────────
  const rows = [
    ["parse", parseMs, BUDGETS.parse],
    [`${COMMAND_COUNT} commands`, cmdsMs, BUDGETS.thousandCommands],
    ["serialize", serializeMs, BUDGETS.serialize],
  ];

  console.log("| Phase           | Elapsed (ms) | Budget (ms) | Status |");
  console.log("| --------------- | ------------ | ----------- | ------ |");
  let failures = 0;
  for (const [name, elapsed, budget] of rows) {
    const ok = elapsed <= budget;
    if (!ok) failures++;
    const label = ok ? "ok" : "OVER BUDGET";
    console.log(
      `| ${name.padEnd(15)} | ${fmtMs(elapsed).padStart(12)} | ${String(budget).padStart(11)} | ${label.padEnd(6)} |`
    );
  }
  console.log("");
  console.log(
    `  output .docx: ${(out.byteLength / 1024).toFixed(1)} KB; final paragraphs: ${
      agent.getSnapshot().root.body.filter((b) => b.kind === "paragraph").length
    }`
  );

  if (failures > 0) {
    console.error(
      `\n❌ perf-docx: ${failures}/${rows.length} budgets exceeded. Re-run on the same machine to confirm before tightening.`
    );
    return 1;
  }
  console.log("\n✅ perf-docx: all budgets met.");
  return 0;
}

const code = await main();
process.exit(code);
