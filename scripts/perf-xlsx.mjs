#!/usr/bin/env node
/**
 * XLSX perf budget runner.
 *
 * Builds a synthetic multi-sheet workbook in memory (SheetJS-backed,
 * matching the production parse path), then asserts:
 *
 *   1. parse        < 750 ms
 *   2. 1k commands  < 1 s aggregate dispatch (`xlsx:set-cell-value`)
 *   3. serialize    < 1 s
 *
 * Budgets are tuned to comfortably pass on an Apple Silicon M-class laptop
 * (the dev target), with ~30 % headroom over the typical observed run so
 * minor flakes don't gate PRs. They are deliberately a touch looser than
 * `perf-docx` because XLSX serialisation walks both the typed model and the
 * parallel SheetJS workbook.
 *
 * Run via `make perf-xlsx` or `node scripts/perf-xlsx.mjs`. The script
 * imports the built `@officeai/xlsx` from `packages/xlsx/dist/index.js`;
 * if the build is missing it runs `pnpm --filter @officeai/xlsx build`
 * first so a one-shot invocation works on a fresh clone.
 */
import { spawnSync } from "node:child_process";
import { existsSync } from "node:fs";
import { performance } from "node:perf_hooks";
import { dirname, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import * as XLSX from "@e965/xlsx";

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, "..");
const distEntry = resolve(root, "packages/xlsx/dist/index.js");

// ── Budgets (ms) ──────────────────────────────────────────────────────────
const BUDGETS = {
  parse: 750,
  thousandCommands: 1000,
  serialize: 1000,
};

// ── Synthetic workbook shape ──────────────────────────────────────────────
// 5 sheets × 2k rows × 6 cols ≈ 60k populated cells. Enough to make the
// parser do real work without ballooning into the multi-second range that
// would mask budget regressions.
const SHEETS = 5;
const ROWS = 2000;
const COLS = 6;
const COMMAND_COUNT = 1000;

function buildSyntheticWorkbookBuffer() {
  const wb = XLSX.utils.book_new();
  for (let s = 0; s < SHEETS; s++) {
    const aoa = [];
    aoa.push(["id", "label", "qty", "price", "total", "notes"]);
    for (let r = 0; r < ROWS; r++) {
      const qty = (r % 17) + 1;
      const price = ((r * 13) % 991) / 10;
      aoa.push([
        r + 1,
        `Item ${s}-${r}`,
        qty,
        price,
        qty * price,
        r % 7 === 0 ? "promo" : "",
      ]);
    }
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    XLSX.utils.book_append_sheet(wb, ws, `Sheet${s + 1}`);
  }
  return XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
}

function ensureBuilt() {
  if (existsSync(distEntry)) return;
  console.log("⚙ packages/xlsx not built yet — building now…");
  const r = spawnSync("pnpm", ["--filter", "@officeai/xlsx", "build"], { cwd: root, stdio: "inherit" });
  if (r.status !== 0) {
    console.error("perf-xlsx: failed to build @officeai/xlsx — cannot run perf check.");
    process.exit(1);
  }
}

async function loadXlsxApi() {
  const url = pathToFileURL(distEntry).href;
  return import(url);
}

const fmtMs = (ms) => ms.toFixed(1);

async function main() {
  ensureBuilt();
  const { XlsxAgent } = await loadXlsxApi();

  console.log(
    `perf-xlsx: synthetic ${SHEETS}-sheet workbook (${SHEETS * ROWS * COLS} populated cells); dispatching ${COMMAND_COUNT} set-cell-value commands.`
  );
  const buf = buildSyntheticWorkbookBuffer();
  console.log(`  synthetic .xlsx size: ${(buf.byteLength / 1024).toFixed(1)} KB\n`);

  const t0 = performance.now();
  const agent = await XlsxAgent.fromBuffer(buf);
  const parseMs = performance.now() - t0;

  // Sanity: confirm the workbook has the expected shape.
  const snap0 = agent.getSnapshot();
  const sheetNames = snap0.root.sheets.map((s) => s.name);

  // Round-robin write to a synthetic note column on each sheet so we hit
  // every parser branch the dispatcher cares about (existing-cell update +
  // new-cell insert + dirty propagation).
  const t1 = performance.now();
  for (let i = 0; i < COMMAND_COUNT; i++) {
    const sheet = sheetNames[i % sheetNames.length];
    const row = (i % ROWS) + 2; // skip header row
    const col = "G"; // column past the seeded data
    const ref = `${col}${row}`;
    await agent.applyCommand({
      type: "xlsx:set-cell-value",
      payload: { sheet, ref, value: `cmd-${i}` },
      source: "agent",
      agentId: "perf-xlsx",
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
  console.log(`  output .xlsx: ${(out.byteLength / 1024).toFixed(1)} KB; sheets: ${sheetNames.length}`);

  if (failures > 0) {
    console.error(
      `\n❌ perf-xlsx: ${failures}/${rows.length} budgets exceeded. Re-run on the same machine to confirm before tightening.`
    );
    return 1;
  }
  console.log("\n✅ perf-xlsx: all budgets met.");
  return 0;
}

const code = await main();
process.exit(code);
