#!/usr/bin/env node
/**
 * Fixture matrix gate.
 *
 * Validates `fixtures/MATRIX.json` as the shared, machine-readable
 * inventory for roundtrip, MCP and web smoke tests.
 */

import { existsSync, readdirSync, readFileSync, statSync } from "node:fs";
import { extname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = resolve(fileURLToPath(import.meta.url), "..", "..");
const MATRIX_PATH = join(ROOT, "fixtures/MATRIX.json");
const FORMATS = new Set(["docx", "xlsx", "pptx", "pdf"]);
const COMPLEXITIES = new Set(["simple", "complex"]);
const ORIGINS = new Set(["synthetic", "generated-real-shape", "collected-real", "public-domain"]);
const EXPECTED_BEHAVIORS = new Set([
  "import",
  "projection",
  "noop-roundtrip",
  "mutation-roundtrip",
  "preserve-opaque",
  "diagnose-unsupported",
  "performance-smoke",
]);
const EXTENSIONS = new Map([
  ["docx", ".docx"],
  ["xlsx", ".xlsx"],
  ["pptx", ".pptx"],
  ["pdf", ".pdf"],
]);
const FIXTURE_ROOTS = ["fixtures/docx", "fixtures/xlsx", "fixtures/pptx", "fixtures/pdf"];

function main() {
  const violations = [];
  if (!existsSync(MATRIX_PATH)) {
    console.error("fixture-matrix: missing fixtures/MATRIX.json");
    return 1;
  }

  const matrix = JSON.parse(readFileSync(MATRIX_PATH, "utf8"));
  if (matrix.version !== 1) violations.push("matrix.version must be 1");
  if (typeof matrix.updated !== "string" || !/^\d{4}-\d{2}-\d{2}$/.test(matrix.updated)) {
    violations.push("matrix.updated must be YYYY-MM-DD");
  }
  if (!matrix.policy || typeof matrix.policy !== "object") {
    violations.push("matrix.policy is required");
  }
  if (!Array.isArray(matrix.fixtures)) {
    violations.push("matrix.fixtures must be an array");
  }

  const fixtures = Array.isArray(matrix.fixtures) ? matrix.fixtures : [];
  const seenIds = new Set();
  const seenPaths = new Set();
  const counts = new Map();
  const riskByFormat = new Map();

  for (const f of fixtures) {
    const label = typeof f?.id === "string" ? f.id : "<missing-id>";
    if (typeof f?.id !== "string" || f.id.length === 0) {
      violations.push(`${label}: id is required`);
    } else if (seenIds.has(f.id)) {
      violations.push(`${label}: duplicate id`);
    } else {
      seenIds.add(f.id);
    }

    if (!FORMATS.has(f?.format)) violations.push(`${label}: invalid format ${String(f?.format)}`);
    if (!COMPLEXITIES.has(f?.complexity)) {
      violations.push(`${label}: complexity must be simple or complex`);
    }
    if (!ORIGINS.has(f?.origin)) violations.push(`${label}: invalid origin ${String(f?.origin)}`);
    if (typeof f?.path !== "string" || f.path.length === 0) {
      violations.push(`${label}: path is required`);
    } else {
      const abs = join(ROOT, f.path);
      if (!existsSync(abs)) {
        violations.push(`${label}: file does not exist at ${f.path}`);
      } else if (!statSync(abs).isFile()) {
        violations.push(`${label}: path is not a file`);
      } else if (statSync(abs).size <= 0) {
        violations.push(`${label}: file is empty`);
      }
      seenPaths.add(f.path);
      const expectedExt = EXTENSIONS.get(f.format);
      if (expectedExt && extname(f.path) !== expectedExt) {
        violations.push(`${label}: path extension must be ${expectedExt}`);
      }
    }

    if (typeof f?.license !== "string" || f.license.length === 0) {
      violations.push(`${label}: license is required`);
    }
    if (typeof f?.source !== "string" || f.source.length === 0) {
      violations.push(`${label}: source is required`);
    }
    if (!Array.isArray(f?.features) || f.features.length === 0) {
      violations.push(`${label}: features must be a non-empty array`);
    }
    if (!Array.isArray(f?.expectedBehaviors) || f.expectedBehaviors.length === 0) {
      violations.push(`${label}: expectedBehaviors must be a non-empty array`);
    } else {
      for (const b of f.expectedBehaviors) {
        if (!EXPECTED_BEHAVIORS.has(b)) violations.push(`${label}: unknown expectedBehavior ${String(b)}`);
      }
    }
    if (!Array.isArray(f?.knownRisks)) {
      violations.push(`${label}: knownRisks must be an array`);
    }

    if (FORMATS.has(f?.format) && COMPLEXITIES.has(f?.complexity)) {
      const key = `${f.format}:${f.complexity}`;
      counts.set(key, (counts.get(key) ?? 0) + 1);
    }
    if (FORMATS.has(f?.format) && Array.isArray(f?.knownRisks) && f.knownRisks.length > 0) {
      const behaviors = new Set(f.expectedBehaviors ?? []);
      if (behaviors.has("preserve-opaque") || behaviors.has("diagnose-unsupported")) {
        riskByFormat.set(f.format, (riskByFormat.get(f.format) ?? 0) + 1);
      }
    }
  }

  for (const format of FORMATS) {
    for (const complexity of COMPLEXITIES) {
      const count = counts.get(`${format}:${complexity}`) ?? 0;
      if (count < 3) {
        violations.push(`${format}: expected at least 3 ${complexity} fixtures, found ${count}`);
      }
    }
    if ((riskByFormat.get(format) ?? 0) < 1) {
      violations.push(`${format}: expected at least one risk fixture with preserve/diagnose behavior`);
    }
  }

  const discovered = discoverFixtureFiles();
  for (const path of discovered) {
    if (!seenPaths.has(path)) violations.push(`unindexed fixture file: ${path}`);
  }
  for (const path of seenPaths) {
    if (!discovered.has(path)) violations.push(`matrix path is outside known fixture roots: ${path}`);
  }

  console.log("fixture-matrix check");
  console.log("────────────────────");
  for (const format of FORMATS) {
    const simple = counts.get(`${format}:simple`) ?? 0;
    const complex = counts.get(`${format}:complex`) ?? 0;
    const risky = riskByFormat.get(format) ?? 0;
    console.log(
      `  ${format.padEnd(5)} simple=${String(simple).padStart(2)}  complex=${String(complex).padStart(2)}  risk=${String(risky).padStart(2)}`
    );
  }
  console.log(`  total fixtures=${fixtures.length}`);
  console.log("");

  if (violations.length === 0) {
    console.log("fixture-matrix: OK");
    return 0;
  }
  console.log("fixture-matrix: FAILED");
  console.log("");
  for (const v of violations) console.log(`  ✖ ${v}`);
  console.log("");
  console.log(`Total ${violations.length} violation(s).`);
  return 1;
}

function discoverFixtureFiles() {
  const out = new Set();
  for (const root of FIXTURE_ROOTS) {
    walk(join(ROOT, root), (file) => {
      const ext = extname(file);
      if (![...EXTENSIONS.values()].includes(ext)) return;
      out.add(file.slice(ROOT.length + 1));
    });
  }
  return out;
}

function walk(dir, visit) {
  let entries;
  try {
    entries = readdirSync(dir);
  } catch {
    return;
  }
  for (const entry of entries) {
    const full = join(dir, entry);
    const st = statSync(full);
    if (st.isDirectory()) walk(full, visit);
    else if (st.isFile()) visit(full);
  }
}

process.exit(main());
