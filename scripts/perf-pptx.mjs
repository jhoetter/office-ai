#!/usr/bin/env node
/**
 * PPTX perf budget runner.
 *
 * Builds a synthetic 100-slide deck via `pptxgenjs` (the same emitter that
 * generates our synthetic fixtures), then asserts:
 *
 *   1. parse        < 750 ms
 *   2. 1k commands  < 1 s aggregate dispatch (`pptx:set-text` against the
 *                    primary text shape on each slide, round-robined)
 *   3. serialize    < 1 s
 *
 * Budgets are tuned to comfortably pass on an Apple Silicon M-class laptop
 * (the dev target), with ~30 % headroom over the typical observed run so
 * minor flakes don't gate PRs.
 *
 * Run via `make perf-pptx` or `node scripts/perf-pptx.mjs`. The script
 * imports the built `@officeai/pptx` from `packages/pptx/dist/index.js`;
 * if the build is missing it runs `pnpm --filter @officeai/pptx build`
 * first so a one-shot invocation works on a fresh clone.
 */
import { spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import { performance } from "node:perf_hooks";
import { dirname, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import PptxGenJSImport from "pptxgenjs";

const PptxGenJS = PptxGenJSImport.default ?? PptxGenJSImport;

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, "..");
const distEntry = resolve(root, "packages/pptx/dist/index.js");

// ── Budgets (ms) ──────────────────────────────────────────────────────────
const BUDGETS = {
  parse: 750,
  thousandCommands: 1000,
  serialize: 1000,
};

// ── Synthetic deck shape ──────────────────────────────────────────────────
// 100 slides × 1 title shape = enough to exercise the slides loop, the
// shape-id resolver, and the dirty-slides set. We keep the shape count
// per slide low so the budget targets per-slide overhead rather than
// shape-count scaling (separate concern).
const SLIDES = 100;
const COMMAND_COUNT = 1000;

async function buildSyntheticDeckBuffer() {
  const pptx = new PptxGenJS();
  pptx.layout = "LAYOUT_WIDE";
  pptx.author = "officeAI";
  pptx.company = "officeAI";
  pptx.title = "perf-pptx synthetic deck";
  for (let i = 0; i < SLIDES; i++) {
    const slide = pptx.addSlide();
    slide.addText(`Slide ${i + 1}`, {
      x: 0.5,
      y: 0.5,
      w: 12,
      h: 1.5,
      fontSize: 36,
      bold: true,
      color: "111827",
    });
    slide.addText(`Body content for slide ${i + 1}.`, {
      x: 0.5,
      y: 2.5,
      w: 12,
      h: 4,
      fontSize: 18,
      color: "374151",
    });
  }
  return Buffer.from(await pptx.write({ outputType: "nodebuffer" }));
}

function ensureBuilt() {
  if (existsSync(distEntry)) return;
  console.log("⚙ packages/pptx not built yet — building now…");
  const r = spawnSync("pnpm", ["--filter", "@officeai/pptx", "build"], { cwd: root, stdio: "inherit" });
  if (r.status !== 0) {
    console.error("perf-pptx: failed to build @officeai/pptx — cannot run perf check.");
    process.exit(1);
  }
}

async function loadPptxApi() {
  const url = pathToFileURL(distEntry).href;
  return import(url);
}

const fmtMs = (ms) => ms.toFixed(1);

/**
 * Find the first text-shape id on each slide. We need the IDs ahead of
 * time so the dispatch loop measures only command apply, not lookup.
 */
function collectFirstTextShapeIds(snapshot) {
  const ids = [];
  for (let i = 0; i < snapshot.root.slides.length; i++) {
    const slide = snapshot.root.slides[i];
    const shape = slide.shapes.find((s) => s.kind === "text");
    ids.push(shape ? { slideIndex: i, shapeId: shape.id } : null);
  }
  return ids.filter((x) => x !== null);
}

async function main() {
  ensureBuilt();
  const { PptxAgent } = await loadPptxApi();

  console.log(
    `perf-pptx: synthetic ${SLIDES}-slide deck; dispatching ${COMMAND_COUNT} pptx:set-text commands.`
  );
  const buf = await buildSyntheticDeckBuffer();
  console.log(`  synthetic .pptx size: ${(buf.byteLength / 1024).toFixed(1)} KB\n`);

  const t0 = performance.now();
  const agent = await PptxAgent.fromBuffer(buf);
  const parseMs = performance.now() - t0;

  const snap0 = agent.getSnapshot();
  const targets = collectFirstTextShapeIds(snap0);
  if (targets.length === 0) {
    console.error("perf-pptx: no text shapes found in synthetic deck — cannot dispatch commands.");
    return 1;
  }

  const t1 = performance.now();
  for (let i = 0; i < COMMAND_COUNT; i++) {
    const target = targets[i % targets.length];
    await agent.applyCommand({
      type: "pptx:set-text",
      payload: {
        slideIndex: target.slideIndex,
        shapeId: target.shapeId,
        text: `cmd-${i}`,
      },
      source: "agent",
      agentId: "perf-pptx",
    });
  }
  const cmdsMs = performance.now() - t1;

  const t2 = performance.now();
  const out = await agent.exportFile();
  const serializeMs = performance.now() - t2;

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
  console.log(`  output .pptx: ${(out.byteLength / 1024).toFixed(1)} KB; slides: ${snap0.root.slides.length}`);

  if (failures > 0) {
    console.error(
      `\n❌ perf-pptx: ${failures}/${rows.length} budgets exceeded. Re-run on the same machine to confirm before tightening.`
    );
    return 1;
  }
  console.log("\n✅ perf-pptx: all budgets met.");
  return 0;
}

const code = await main();
process.exit(code);
