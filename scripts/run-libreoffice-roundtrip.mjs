#!/usr/bin/env node
/**
 * LibreOffice headless roundtrip runner.
 *
 * For each .docx in fixtures/docx/real-world/, this script:
 *   1. Renders the input to PDF via `soffice --headless --convert-to pdf`.
 *   2. Loads the input through @officeai/docx, exports it back to bytes,
 *      and renders THAT to PDF as well.
 *   3. Asserts both conversions exit 0 and emit no "repair" / "error"
 *      messages on stderr (case-insensitive).
 *
 * Exit semantics (so it composes with `make verify` cleanly):
 *   - exit 0  → all conversions clean.
 *   - exit 0 + warning → `soffice` not on PATH (graceful skip; the CI job
 *     installs LibreOffice explicitly so this branch only runs on dev
 *     machines that don't have it).
 *   - exit 1  → at least one conversion failed or surfaced repair text.
 *
 * Run via: `make roundtrip-libre` or `node scripts/run-libreoffice-roundtrip.mjs`.
 */

import { spawnSync } from "node:child_process";
import { mkdirSync, mkdtempSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { pathToFileURL } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, "..");
const fixtureDir = resolve(root, "fixtures/docx/real-world");

const REPAIR_HINTS = [/\brepair/i, /\berror\b/i, /\bcorrupt/i, /unable to load/i, /failed to/i];

function findSoffice() {
  // `which`/`where` shells out cheaper than booting LibreOffice. macOS
  // `homebrew` and Linux package managers both put it on PATH; CI installs
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

function listFixtures() {
  try {
    return readdirSync(fixtureDir)
      .filter((f) => f.toLowerCase().endsWith(".docx"))
      .sort()
      .map((f) => join(fixtureDir, f));
  } catch {
    return [];
  }
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

async function loadDocxAgent() {
  // Resolve through the workspace so the script picks up the built tarball
  // OR the local source via tsconfig path mapping. We import dynamically so
  // a missing build doesn't blow up `--help`-style invocations.
  const distEntry = resolve(root, "packages/docx/dist/index.js");
  const url = pathToFileURL(distEntry).href;
  return import(url);
}

async function roundtripBuffer(input) {
  const { DocxAgent } = await loadDocxAgent();
  const buf = readFileSync(input);
  const agent = await DocxAgent.fromBuffer(buf);
  const out = await agent.exportFile();
  return Buffer.from(out);
}

async function main() {
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
  const fixtures = listFixtures();
  if (fixtures.length === 0) {
    console.warn(`⚠ no fixtures found in ${fixtureDir}. Run \`pnpm fixtures-real\` first.`);
    return 0;
  }

  console.log(`✓ using soffice at ${soffice}`);
  console.log(`✓ checking ${fixtures.length} fixtures from ${fixtureDir}\n`);

  const workDir = mkdtempSync(join(tmpdir(), "officeai-libre-"));
  const inputPdfDir = join(workDir, "pdf-input");
  const roundtripDir = join(workDir, "roundtrip");
  const roundtripPdfDir = join(workDir, "pdf-roundtrip");
  mkdirSync(inputPdfDir, { recursive: true });
  mkdirSync(roundtripDir, { recursive: true });
  mkdirSync(roundtripPdfDir, { recursive: true });

  let failures = 0;
  for (const input of fixtures) {
    const name = input.split("/").pop();
    process.stdout.write(`  ${name} … `);

    // 1. Original → PDF.
    const r1 = convertToPdf(soffice, input, inputPdfDir);
    if (r1.error || r1.code !== 0 || flagsRepair(r1.stderr) || flagsRepair(r1.stdout)) {
      failures++;
      console.log("FAIL (input → PDF)");
      console.log("    stderr:", r1.stderr.trim() || "(empty)");
      console.log("    stdout:", r1.stdout.trim() || "(empty)");
      continue;
    }

    // 2. Roundtrip via DocxAgent → second .docx → PDF.
    let roundBuf;
    try {
      roundBuf = await roundtripBuffer(input);
    } catch (err) {
      failures++;
      console.log("FAIL (DocxAgent roundtrip)");
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

  // Best effort: leave artifacts on failure so a developer can inspect.
  if (failures === 0) {
    rmSync(workDir, { recursive: true, force: true });
    console.log(`\n✅ all ${fixtures.length} fixtures roundtrip clean through LibreOffice.`);
    return 0;
  }

  console.error(`\n❌ ${failures} of ${fixtures.length} fixtures failed. Artifacts in ${workDir}`);
  return 1;
}

const code = await main();
process.exit(code);
