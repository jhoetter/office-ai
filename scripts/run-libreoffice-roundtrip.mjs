#!/usr/bin/env node
/**
 * LibreOffice headless roundtrip runner — multi-format.
 *
 * For each fixture in the configured directories, this script:
 *   1. Renders the input to PDF via `soffice --headless --convert-to pdf`.
 *   2. Loads the input through the format's @officeai agent, exports it
 *      back to bytes, and renders THAT to PDF as well.
 *   3. Asserts both conversions exit 0 and emit no "repair" / "error"
 *      messages on stderr (case-insensitive).
 *
 * Usage:
 *   node scripts/run-libreoffice-roundtrip.mjs --format docx
 *   node scripts/run-libreoffice-roundtrip.mjs --format xlsx
 *   node scripts/run-libreoffice-roundtrip.mjs --format pptx
 *
 * Defaults to `--format docx` for backwards compatibility with the
 * original DOCX-only entry point.
 *
 * Exit semantics (so it composes with `make verify` cleanly):
 *   - exit 0  → all conversions clean.
 *   - exit 0 + warning → `soffice` not on PATH (graceful skip; the CI job
 *     installs LibreOffice explicitly so this branch only runs on dev
 *     machines that don't have it).
 *   - exit 0 + warning → no fixtures present (XLSX synthetic dir might
 *     be empty pre-`pnpm fixtures-xlsx`; pptx-real isn't always built).
 *   - exit 1  → at least one conversion failed or surfaced repair text.
 */

import { spawnSync } from "node:child_process";
import { existsSync, mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, "..");

const REPAIR_HINTS = [/\brepair/i, /\berror\b/i, /\bcorrupt/i, /unable to load/i, /failed to/i];

// ── per-format config ────────────────────────────────────────────────
//
// `fixtureDirs` is an ordered list — every existing dir contributes
// fixtures. `agentEntry` is the *built* dist entry we dynamic-import;
// `agentName` is the export we look for on that module.
const FORMATS = {
  docx: {
    extension: ".docx",
    fixtureDirs: ["fixtures/docx/real-world"],
    agentEntry: "packages/docx/dist/index.js",
    agentName: "DocxAgent",
  },
  xlsx: {
    extension: ".xlsx",
    fixtureDirs: ["fixtures/xlsx/synthetic", "fixtures/xlsx/real-world"],
    agentEntry: "packages/xlsx/dist/index.js",
    agentName: "XlsxAgent",
  },
  pptx: {
    extension: ".pptx",
    fixtureDirs: ["fixtures/pptx/synthetic", "fixtures/pptx/real"],
    agentEntry: "packages/pptx/dist/index.js",
    agentName: "PptxAgent",
  },
};

function parseArgs(argv) {
  const args = { format: "docx" };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === "--format" && argv[i + 1]) {
      args.format = argv[++i];
    } else if (a.startsWith("--format=")) {
      args.format = a.slice("--format=".length);
    } else if (a === "--help" || a === "-h") {
      args.help = true;
    }
  }
  return args;
}

function findSoffice() {
  // `which`/`where` shells out cheaper than booting LibreOffice. macOS
  // homebrew and Linux package managers both put it on PATH; CI installs
  // it via apt, which also lands on PATH.
  const probe = spawnSync(process.platform === "win32" ? "where" : "which", ["soffice"], {
    encoding: "utf8",
  });
  if (probe.status === 0) {
    const found = probe.stdout.split(/\r?\n/).find((l) => l.trim().length > 0);
    if (found) return found.trim();
  }
  return null;
}

function listFixtures(format) {
  const out = [];
  for (const rel of format.fixtureDirs) {
    const dir = resolve(root, rel);
    if (!existsSync(dir)) continue;
    let entries;
    try {
      entries = readdirSync(dir);
    } catch {
      continue;
    }
    for (const f of entries.sort()) {
      if (f.toLowerCase().endsWith(format.extension)) {
        out.push(join(dir, f));
      }
    }
  }
  return out;
}

function convertToPdf(soffice, input, outDir) {
  const result = spawnSync(
    soffice,
    ["--headless", "--norestore", "--nologo", "--convert-to", "pdf", "--outdir", outDir, input],
    { encoding: "utf8", timeout: 60_000 }
  );
  return {
    code: result.status,
    stdout: result.stdout ?? "",
    stderr: result.stderr ?? "",
    error: result.error,
  };
}

function flagsRepair(text) {
  return REPAIR_HINTS.some((re) => re.test(text));
}

async function loadAgent(format) {
  const distEntry = resolve(root, format.agentEntry);
  if (!existsSync(distEntry)) {
    throw new Error(
      `Agent entry not found at ${distEntry}. Run \`pnpm --filter @officeai/${format.agentName.replace("Agent", "").toLowerCase()} build\` first.`
    );
  }
  const url = pathToFileURL(distEntry).href;
  const mod = await import(url);
  const Agent = mod[format.agentName];
  if (!Agent) {
    throw new Error(`Module ${distEntry} does not export ${format.agentName}.`);
  }
  return Agent;
}

async function roundtripBuffer(Agent, input) {
  const buf = readFileSync(input);
  const agent = await Agent.fromBuffer(buf);
  const out = await agent.exportFile();
  return Buffer.from(out);
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help) {
    console.log("Usage: run-libreoffice-roundtrip.mjs --format docx|xlsx|pptx");
    return 0;
  }
  const format = FORMATS[args.format];
  if (!format) {
    console.error(`Unknown --format ${args.format}. Use one of: docx, xlsx, pptx.`);
    return 1;
  }

  const soffice = findSoffice();
  if (!soffice) {
    console.warn(
      "⚠ soffice not found on PATH — skipping LibreOffice roundtrip.\n" +
        "  Install LibreOffice (e.g. `brew install --cask libreoffice` or\n" +
        "  `apt-get install libreoffice`) to run the full check locally.\n" +
        "  CI installs it explicitly so the gate still runs in CI."
    );
    return 0;
  }

  const Agent = await loadAgent(format);
  const fixtures = listFixtures(format);
  if (fixtures.length === 0) {
    console.warn(
      `⚠ no ${args.format} fixtures found in ${format.fixtureDirs.join(", ")}.\n` +
        `  Generate them with \`pnpm fixtures-${args.format}\` (synthetic) or the relevant fixtures-real script.`
    );
    return 0;
  }

  console.log(`✓ format: ${args.format}`);
  console.log(`✓ using soffice at ${soffice}`);
  console.log(`✓ checking ${fixtures.length} fixtures\n`);

  const workDir = mkdtempSync(join(tmpdir(), `officeai-libre-${args.format}-`));
  const inputPdfDir = join(workDir, "pdf-input");
  const roundtripDir = join(workDir, "roundtrip");
  const roundtripPdfDir = join(workDir, "pdf-roundtrip");
  mkdirSync(inputPdfDir, { recursive: true });
  mkdirSync(roundtripDir, { recursive: true });
  mkdirSync(roundtripPdfDir, { recursive: true });

  let failures = 0;
  let skipped = 0;
  for (const input of fixtures) {
    const name = input.split("/").pop();
    process.stdout.write(`  ${name} … `);

    const r1 = convertToPdf(soffice, input, inputPdfDir);
    if (r1.error || r1.code !== 0 || flagsRepair(r1.stderr) || flagsRepair(r1.stdout)) {
      // The *input* fixture itself fails to render cleanly through
      // LibreOffice (e.g. corrupt embedded PNG, libpng IDAT warnings).
      // That is not a regression caused by our roundtrip — it's a dirty
      // baseline. Skip with a warning so we don't mask real failures
      // but also don't fail the gate on pre-existing fixture issues.
      skipped++;
      console.log("skip (input fixture dirty before roundtrip)");
      console.log("    stderr:", (r1.stderr || "").trim().split("\n").slice(0, 2).join(" | ") || "(empty)");
      continue;
    }

    let roundBuf;
    try {
      roundBuf = await roundtripBuffer(Agent, input);
    } catch (err) {
      failures++;
      console.log(`FAIL (${format.agentName} roundtrip)`);
      console.log("    ", err instanceof Error ? err.message : String(err));
      continue;
    }
    const roundPath = join(roundtripDir, name);
    writeFileSync(roundPath, roundBuf);

    const r2 = convertToPdf(soffice, roundPath, roundtripPdfDir);
    if (r2.error || r2.code !== 0 || flagsRepair(r2.stderr) || flagsRepair(r2.stdout)) {
      failures++;
      console.log("FAIL (roundtrip → PDF)");
      console.log("    stderr:", r2.stderr.trim() || "(empty)");
      console.log("    stdout:", r2.stdout.trim() || "(empty)");
      continue;
    }

    console.log("ok");
  }

  if (failures === 0) {
    rmSync(workDir, { recursive: true, force: true });
    const checked = fixtures.length - skipped;
    if (skipped > 0) {
      console.log(
        `\n✅ ${checked}/${fixtures.length} ${args.format} fixtures roundtrip clean through LibreOffice ` +
          `(${skipped} skipped — input fixture dirty before roundtrip).`
      );
    } else {
      console.log(`\n✅ all ${fixtures.length} ${args.format} fixtures roundtrip clean through LibreOffice.`);
    }
    return 0;
  }

  console.error(
    `\n❌ ${failures} of ${fixtures.length} ${args.format} fixtures failed (${skipped} skipped). Artifacts in ${workDir}`
  );
  return 1;
}

const code = await main();
process.exit(code);
