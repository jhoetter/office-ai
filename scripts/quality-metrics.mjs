#!/usr/bin/env node
/**
 * `make metrics` — holistic repo snapshot.
 *
 * Pure reporting tool. Counts source LOC, test files, test cases,
 * fixtures by format, scripts, and CI jobs. Never fails the build.
 * Used to keep the README / build log honest about scale and to spot
 * skew (e.g. one product is shipping 3× the code with 0.1× the tests).
 *
 * Output is plain text + a markdown table so `make metrics > metrics.md`
 * is reasonable.
 *
 * Usage:
 *   node scripts/quality-metrics.mjs
 *   node scripts/quality-metrics.mjs --json    # machine-readable
 *
 * Performance: walks the repo synchronously; should finish in well
 * under a second on a workspace this size. No network, no spawning.
 */

import { readdirSync, readFileSync, statSync } from "node:fs";
import { dirname, extname, join, relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, "..");

const argSet = new Set(process.argv.slice(2));
const AS_JSON = argSet.has("--json");

const SOURCE_EXTS = new Set([".ts", ".tsx", ".js", ".jsx", ".mjs", ".cjs"]);
const TEST_EXTS = new Set([".test.ts", ".test.tsx", ".spec.ts", ".spec.tsx"]);
const PRUNE = new Set([
  "node_modules",
  ".next",
  ".turbo",
  "dist",
  "build",
  "coverage",
  "playwright-report",
  ".git",
  "vendor",
]);

/** Recursively walks `dir`, yielding absolute paths to non-pruned files. */
function* walk(dir) {
  let entries;
  try {
    entries = readdirSync(dir, { withFileTypes: true });
  } catch {
    return;
  }
  for (const entry of entries) {
    if (PRUNE.has(entry.name)) continue;
    const abs = join(dir, entry.name);
    if (entry.isDirectory()) yield* walk(abs);
    else yield abs;
  }
}

/**
 * Returns true if the path matches `<base>.test.<ext>` or
 * `<base>.spec.<ext>`. extname() only returns the trailing .ts / .tsx so
 * we have to look at the full filename.
 */
function isTestFile(name) {
  for (const ext of TEST_EXTS) if (name.endsWith(ext)) return true;
  return false;
}

function loc(absPath) {
  try {
    const buf = readFileSync(absPath, "utf8");
    if (!buf) return 0;
    return buf.split("\n").length;
  } catch {
    return 0;
  }
}

/**
 * Cheap heuristic: count `it(`, `test(`, `describe(`, and `bench(`
 * call-sites. Doesn't try to be a parser — close enough for a
 * dashboard, and fast.
 */
function countTestCases(absPath) {
  try {
    const txt = readFileSync(absPath, "utf8");
    const matches = txt.match(/\b(?:it|test|bench)\s*\(/g);
    return matches ? matches.length : 0;
  } catch {
    return 0;
  }
}

function tally(dirAbs) {
  let sourceFiles = 0;
  let sourceLoc = 0;
  let testFiles = 0;
  let testLoc = 0;
  let testCases = 0;
  for (const file of walk(dirAbs)) {
    const ext = extname(file);
    if (!SOURCE_EXTS.has(ext)) continue;
    const isTest = isTestFile(file);
    const lines = loc(file);
    if (isTest) {
      testFiles++;
      testLoc += lines;
      testCases += countTestCases(file);
    } else {
      sourceFiles++;
      sourceLoc += lines;
    }
  }
  return { sourceFiles, sourceLoc, testFiles, testLoc, testCases };
}

function countFixtures(dirAbs, exts) {
  let n = 0;
  let bytes = 0;
  for (const file of walk(dirAbs)) {
    if (!exts.some((e) => file.toLowerCase().endsWith(e))) continue;
    n++;
    try {
      bytes += statSync(file).size;
    } catch {
      // ignore
    }
  }
  return { count: n, bytes };
}

function fmtBytes(b) {
  if (b < 1024) return `${b} B`;
  if (b < 1024 * 1024) return `${(b / 1024).toFixed(1)} KB`;
  return `${(b / 1024 / 1024).toFixed(2)} MB`;
}

function pad(s, n, align = "left") {
  s = String(s);
  if (s.length >= n) return s;
  const fill = " ".repeat(n - s.length);
  return align === "right" ? fill + s : s + fill;
}

function printTable(headers, rows) {
  const widths = headers.map((h, i) => Math.max(h.length, ...rows.map((r) => String(r[i]).length)));
  const sep = "| " + widths.map((w) => "-".repeat(w)).join(" | ") + " |";
  console.log("| " + headers.map((h, i) => pad(h, widths[i])).join(" | ") + " |");
  console.log(sep);
  for (const r of rows) {
    console.log(
      "| " + r.map((c, i) => pad(c, widths[i], typeof c === "number" ? "right" : "left")).join(" | ") + " |"
    );
  }
}

function listCiJobs() {
  const ciYml = resolve(root, ".github/workflows/ci.yml");
  let txt;
  try {
    txt = readFileSync(ciYml, "utf8");
  } catch {
    return { count: 0, advisory: 0, jobs: [] };
  }
  // Job entries are top-level keys under `jobs:`. We don't need a real
  // YAML parser here — every job in this file uses 2-space indent and
  // the `name:` field. Keep the scan dead simple and tolerant.
  const lines = txt.split("\n");
  const jobs = [];
  let inJobs = false;
  let current = null;
  for (const line of lines) {
    if (line.startsWith("jobs:")) {
      inJobs = true;
      continue;
    }
    if (!inJobs) continue;
    if (/^\S/.test(line)) {
      // Top-level key after `jobs:` ends the block.
      inJobs = false;
      continue;
    }
    const jobMatch = line.match(/^  ([a-zA-Z0-9_-]+):\s*$/);
    if (jobMatch) {
      if (current) jobs.push(current);
      current = { id: jobMatch[1], name: jobMatch[1], advisory: false };
      continue;
    }
    if (current) {
      const nameMatch = line.match(/^    name:\s*(.+)$/);
      if (nameMatch) current.name = nameMatch[1].trim();
      const advMatch = line.match(/^    continue-on-error:\s*true/);
      if (advMatch) current.advisory = true;
    }
  }
  if (current) jobs.push(current);
  return {
    count: jobs.length,
    advisory: jobs.filter((j) => j.advisory).length,
    jobs,
  };
}

function main() {
  const packages = [
    "packages/core",
    "packages/docx",
    "packages/xlsx",
    "packages/pptx",
    "packages/agent",
    "packages/ui",
    "apps/web",
  ];
  const perPackage = packages.map((p) => ({ name: p, ...tally(resolve(root, p)) }));

  const fixtures = {
    docx: countFixtures(resolve(root, "fixtures/docx"), [".docx"]),
    xlsx: countFixtures(resolve(root, "fixtures/xlsx"), [".xlsx"]),
    pptx: countFixtures(resolve(root, "fixtures/pptx"), [".pptx"]),
  };

  const scripts = (() => {
    let n = 0;
    for (const f of walk(resolve(root, "scripts"))) if (extname(f) === ".mjs") n++;
    return n;
  })();

  const ci = listCiJobs();

  if (AS_JSON) {
    console.log(JSON.stringify({ packages: perPackage, fixtures, scripts, ci }, null, 2));
    return;
  }

  console.log(`# office-ai — repo metrics`);
  console.log(`# generated by scripts/quality-metrics.mjs (${new Date().toISOString()})`);
  console.log("");

  console.log(`## Source / tests by package`);
  console.log("");
  printTable(
    ["package", "src files", "src LOC", "test files", "test LOC", "test cases"],
    perPackage.map((p) => [
      relative(root, resolve(root, p.name)),
      p.sourceFiles,
      p.sourceLoc,
      p.testFiles,
      p.testLoc,
      p.testCases,
    ])
  );
  const total = perPackage.reduce(
    (acc, p) => ({
      sourceFiles: acc.sourceFiles + p.sourceFiles,
      sourceLoc: acc.sourceLoc + p.sourceLoc,
      testFiles: acc.testFiles + p.testFiles,
      testLoc: acc.testLoc + p.testLoc,
      testCases: acc.testCases + p.testCases,
    }),
    { sourceFiles: 0, sourceLoc: 0, testFiles: 0, testLoc: 0, testCases: 0 }
  );
  console.log("");
  console.log(
    `total: ${total.sourceFiles} src files / ${total.sourceLoc} LOC; ${total.testFiles} test files / ${total.testLoc} LOC / ${total.testCases} cases.`
  );
  console.log("");

  console.log(`## Fixtures`);
  console.log("");
  printTable(
    ["format", "count", "size"],
    [
      ["docx", fixtures.docx.count, fmtBytes(fixtures.docx.bytes)],
      ["xlsx", fixtures.xlsx.count, fmtBytes(fixtures.xlsx.bytes)],
      ["pptx", fixtures.pptx.count, fmtBytes(fixtures.pptx.bytes)],
    ]
  );
  console.log("");

  console.log(`## CI / scripts`);
  console.log("");
  console.log(`- top-level scripts/*.mjs : ${scripts}`);
  console.log(`- CI jobs                  : ${ci.count} (${ci.advisory} advisory)`);
  console.log("");
  for (const j of ci.jobs) {
    console.log(`  - ${j.id} → "${j.name}"${j.advisory ? "  (advisory)" : ""}`);
  }
}

main();
