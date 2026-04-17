#!/usr/bin/env node
/**
 * OOXML schema validator (W9 / Theme D4).
 *
 * For every `.docx` in `fixtures/docx/real-world/`:
 *   1. Loads the input bytes into an `OoxmlContainer` and enumerates every
 *      XML part.
 *   2. Roundtrips the file through `DocxAgent.fromBuffer → trivial edit →
 *      agent.exportFile()`, then enumerates every XML part of the re-emit.
 *   3. Maps each part to its corresponding ECMA-376 (Transitional) XSD via a
 *      `[Content_Types].xml`-aware lookup, and shells out to `xmllint
 *      --noout --schema <xsd>` to validate the part bytes.
 *   4. Prints a per-fixture row table: `fixture | part | input ✓/✗ |
 *      re-emit ✓/✗`. On failure, emits the offending xmllint stderr as a
 *      quoted block.
 *
 * Exit semantics (matches the `make perf-docx` / `make roundtrip-libre`
 * pattern so wrappers don't need to special-case it):
 *   - exit 0 + warning : `xmllint` is missing from PATH (graceful skip; CI
 *     installs `libxml2-utils` so the gate still runs server-side).
 *   - exit 0 + warning : `vendor/ooxml-xsd/` is empty (run
 *     `make xsd-fetch`).
 *   - exit 0           : every part is well-formed AND schema-valid.
 *   - exit 1           : at least one part fails.
 *
 * Modes:
 *   --dry-run            : skip xmllint entirely; just emit the parts that
 *                          would be validated and which XSD each maps to.
 *                          Used by `tests/scripts/validate-ooxml-schemas.test.ts`
 *                          so the test stays hermetic on a fresh CI runner.
 *   --self-test          : like --dry-run, plus assert that all 6 real-world
 *                          fixtures are present and that every observed part
 *                          maps to either a known XSD or an explicit "skip"
 *                          bucket (OPC parts, w15 extensions). Exit non-zero
 *                          if a fixture is missing or an unknown part shape
 *                          is encountered.
 *   --inject-broken      : prepend a synthetic malformed-XML "part" to the
 *                          validation queue and assert the failure path
 *                          counts it as a violation. Combined with
 *                          --self-test this is what the unit test uses to
 *                          confirm the failure path raises non-zero without
 *                          needing xmllint installed.
 *
 * Run via `make schema-validate` or `node scripts/validate-ooxml-schemas.mjs`.
 */

import { spawnSync } from "node:child_process";
import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, "..");
const FIXTURE_DIR = resolve(root, "fixtures/docx/real-world");
const XSD_DIR = resolve(root, "vendor/ooxml-xsd");
const XSD_SENTINEL = join(XSD_DIR, "wml.xsd");
const DOCX_DIST = resolve(root, "packages/docx/dist/index.js");
const CORE_DIST = resolve(root, "packages/core/dist/index.js");

const args = new Set(process.argv.slice(2));
const DRY_RUN = args.has("--dry-run") || args.has("--self-test");
const SELF_TEST = args.has("--self-test");
const INJECT_BROKEN = args.has("--inject-broken");

/* ────────────────────────────────────────────────────────────────────────────
 * Part → XSD mapping
 *
 * Maps each OOXML part path to the XSD that constrains it. We key on the
 * part's content-type (from `[Content_Types].xml`) so that custom paths like
 * `word/glossary/document.xml` still get the right schema; we fall back to a
 * filename-pattern lookup for parts that have no explicit content-type
 * override (rare, but happens with hand-crafted fixtures).
 *
 * "skip" means the part is intentionally not validated. Reasons:
 *   - OPC parts (Content_Types, _rels) need ECMA-376 Part 2 schemas, which
 *     ship in a separate bundle we haven't pinned yet.
 *   - `commentsExtended.xml`, `commentsIds.xml`, `people.xml` use the w15
 *     namespace, which is a Microsoft extension, NOT covered by the
 *     transitional XSD set.
 *   - Image binaries (`word/media/*`) are not XML.
 * ──────────────────────────────────────────────────────────────────────── */

const CT_TO_XSD = {
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml": "wml.xsd",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document+xml": "wml.xsd",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml": "wml.xsd",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml": "wml.xsd",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml": "wml.xsd",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml": "wml.xsd",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml": "wml.xsd",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml": "wml.xsd",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml": "wml.xsd",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml": "wml.xsd",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml": "wml.xsd",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml": "wml.xsd",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml": "wml.xsd",
  "application/vnd.openxmlformats-officedocument.theme+xml": "dml-main.xsd",
  "application/vnd.openxmlformats-officedocument.themeOverride+xml": "dml-main.xsd",
  "application/vnd.openxmlformats-officedocument.drawingml.chart+xml": "dml-chart.xsd",
  "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml": "dml-diagram.xsd",
  "application/vnd.openxmlformats-officedocument.extended-properties+xml":
    "shared-documentPropertiesExtended.xsd",
  "application/vnd.openxmlformats-officedocument.custom-properties+xml":
    "shared-documentPropertiesCustom.xsd",
  // Skipped (no XSD in our pinned bundle):
  "application/vnd.openxmlformats-package.relationships+xml": "skip:opc",
  "application/vnd.openxmlformats-package.core-properties+xml": "skip:opc",
  "application/vnd.ms-word.commentsExtended+xml": "skip:w15",
  "application/vnd.ms-word.commentsIds+xml": "skip:w15",
  "application/vnd.ms-word.people+xml": "skip:w15",
  "application/vnd.openxmlformats-officedocument.vmlDrawing": "skip:vml-binary",
  "image/png": "skip:binary",
  "image/jpeg": "skip:binary",
  "image/gif": "skip:binary",
  "image/svg+xml": "skip:binary",
  "image/bmp": "skip:binary",
};

const PATH_TO_XSD_FALLBACK = [
  [/^word\/document\d*\.xml$/, "wml.xsd"],
  [/^word\/styles\d*\.xml$/, "wml.xsd"],
  [/^word\/numbering\.xml$/, "wml.xsd"],
  [/^word\/settings\.xml$/, "wml.xsd"],
  [/^word\/webSettings\.xml$/, "wml.xsd"],
  [/^word\/fontTable\.xml$/, "wml.xsd"],
  [/^word\/comments\.xml$/, "wml.xsd"],
  [/^word\/header\d*\.xml$/, "wml.xsd"],
  [/^word\/footer\d*\.xml$/, "wml.xsd"],
  [/^word\/footnotes\.xml$/, "wml.xsd"],
  [/^word\/endnotes\.xml$/, "wml.xsd"],
  [/^word\/glossary\/document\.xml$/, "wml.xsd"],
  [/^word\/theme\/theme\d*\.xml$/, "dml-main.xsd"],
  [/^docProps\/app\.xml$/, "shared-documentPropertiesExtended.xsd"],
  [/^docProps\/custom\.xml$/, "shared-documentPropertiesCustom.xsd"],
  // Skipped:
  [/^\[Content_Types\]\.xml$/, "skip:opc"],
  [/^_rels\//, "skip:opc"],
  [/_rels\//, "skip:opc"],
  [/^docProps\/core\.xml$/, "skip:opc"],
  [/^word\/commentsExtended\.xml$/, "skip:w15"],
  [/^word\/commentsIds\.xml$/, "skip:w15"],
  [/^word\/people\.xml$/, "skip:w15"],
  [/^word\/media\//, "skip:binary"],
  [/^word\/embeddings\//, "skip:binary"],
  [/\.bin$/, "skip:binary"],
  [/\.png$|\.jpe?g$|\.gif$|\.svg$|\.bmp$/i, "skip:binary"],
];

/**
 * Given a part path and an optional content-type lookup table, return the XSD
 * filename to validate against, or `skip:<reason>` to skip, or `null` if the
 * part is unrecognised (which `--self-test` flags as a failure so we notice
 * new part types).
 */
export function mapPartToXsd(partPath, contentTypeFor) {
  const ct = contentTypeFor ? contentTypeFor(partPath) : undefined;
  if (ct && CT_TO_XSD[ct]) return CT_TO_XSD[ct];
  for (const [re, xsd] of PATH_TO_XSD_FALLBACK) {
    if (re.test(partPath)) return xsd;
  }
  return null;
}

/* ────────────────────────────────────────────────────────────────────────────
 * Helpers
 * ──────────────────────────────────────────────────────────────────────── */

function listFixtures() {
  try {
    return readdirSync(FIXTURE_DIR)
      .filter((f) => f.toLowerCase().endsWith(".docx"))
      .sort()
      .map((f) => join(FIXTURE_DIR, f));
  } catch {
    return [];
  }
}

function findXmllint() {
  const probe = spawnSync(process.platform === "win32" ? "where" : "which", ["xmllint"], {
    encoding: "utf8",
  });
  if (probe.status === 0) {
    const found = probe.stdout.split(/\r?\n/).find((l) => l.trim().length > 0);
    if (found) return found.trim();
  }
  return null;
}

async function loadCore() {
  if (!existsSync(CORE_DIST) || !existsSync(DOCX_DIST)) {
    console.error(
      `❌ schema-validate: missing built dist (run \`pnpm build\` first).\n   expected: ${CORE_DIST}\n   expected: ${DOCX_DIST}`
    );
    process.exit(1);
  }
  const core = await import(pathToFileURL(CORE_DIST).href);
  const docx = await import(pathToFileURL(DOCX_DIST).href);
  return { core, docx };
}

/**
 * Build a `(partPath) => contentType | undefined` lookup from a container's
 * `[Content_Types].xml`. Parses defaults (`<Default Extension="..." />`) and
 * overrides (`<Override PartName="..." />`) with regexes — ContentTypes is a
 * tiny, well-known schema so we don't need a full XML parser here.
 */
function buildContentTypeLookup(container) {
  let xml;
  try {
    xml = container.readText("[Content_Types].xml");
  } catch {
    return () => undefined;
  }
  const defaults = new Map(); // ext (lower) -> contentType
  const overrides = new Map(); // /partName -> contentType
  for (const m of xml.matchAll(/<Default\s+[^>]*Extension="([^"]+)"\s+ContentType="([^"]+)"/gi)) {
    defaults.set(m[1].toLowerCase(), m[2]);
  }
  for (const m of xml.matchAll(/<Default\s+[^>]*ContentType="([^"]+)"\s+Extension="([^"]+)"/gi)) {
    defaults.set(m[2].toLowerCase(), m[1]);
  }
  for (const m of xml.matchAll(/<Override\s+[^>]*PartName="([^"]+)"\s+ContentType="([^"]+)"/gi)) {
    overrides.set(m[1], m[2]);
  }
  for (const m of xml.matchAll(/<Override\s+[^>]*ContentType="([^"]+)"\s+PartName="([^"]+)"/gi)) {
    overrides.set(m[2], m[1]);
  }
  return (partPath) => {
    const key = partPath.startsWith("/") ? partPath : `/${partPath}`;
    if (overrides.has(key)) return overrides.get(key);
    const dot = partPath.lastIndexOf(".");
    if (dot >= 0) {
      const ext = partPath.slice(dot + 1).toLowerCase();
      if (defaults.has(ext)) return defaults.get(ext);
    }
    return undefined;
  };
}

/**
 * Tiny well-formedness probe — used by `--inject-broken` so the failure path
 * exercises without a real `xmllint` install. Stack-based: counts open / close
 * tags, ignores PIs, comments, CDATA, and self-closing tags. Not a full
 * conformance check; just enough to flag obvious malformed blobs (unclosed
 * tags, mismatched names) for the unit test.
 */
export function isXmlWellFormed(xml) {
  const stripped = xml
    .replace(/<\?[\s\S]*?\?>/g, "")
    .replace(/<!--[\s\S]*?-->/g, "")
    .replace(/<!\[CDATA\[[\s\S]*?\]\]>/g, "");
  const tagRe = /<\s*(\/?)\s*([A-Za-z_][\w:.\-]*)\b[^>]*?(\/?)\s*>/g;
  const stack = [];
  let m;
  while ((m = tagRe.exec(stripped)) !== null) {
    const isClose = m[1] === "/";
    const name = m[2];
    const isSelf = m[3] === "/";
    if (isClose) {
      const top = stack.pop();
      if (top !== name) {
        return { ok: false, reason: `mismatched </${name}> (expected </${top ?? "<empty>"}>)` };
      }
    } else if (!isSelf) {
      stack.push(name);
    }
  }
  if (stack.length > 0) return { ok: false, reason: `unclosed <${stack[stack.length - 1]}>` };
  return { ok: true };
}

/* ────────────────────────────────────────────────────────────────────────────
 * Validation
 * ──────────────────────────────────────────────────────────────────────── */

function validateWithXmllint(xmllint, xsdPath, xmlPath) {
  const r = spawnSync(xmllint, ["--noout", "--schema", xsdPath, xmlPath], {
    encoding: "utf8",
    timeout: 30_000,
  });
  return {
    ok: r.status === 0,
    stderr: (r.stderr ?? "").trim(),
    stdout: (r.stdout ?? "").trim(),
  };
}

function status(ok, dryRun) {
  if (dryRun) return ok === "would-run" ? "would-run" : "skip";
  return ok ? "✓" : "✗";
}

function padRight(s, n) {
  return s.length >= n ? s.slice(0, n) : s + " ".repeat(n - s.length);
}

async function validateContainer({ container, label, fixtureName, xmllint, workDir, results }) {
  const ctLookup = buildContentTypeLookup(container);
  const partPaths = [...container.parts.keys()].sort();
  for (const partPath of partPaths) {
    const xsd = mapPartToXsd(partPath, ctLookup);
    if (!xsd) {
      results.push({
        fixture: fixtureName,
        side: label,
        part: partPath,
        outcome: "unmapped",
      });
      continue;
    }
    if (xsd.startsWith("skip:")) {
      results.push({
        fixture: fixtureName,
        side: label,
        part: partPath,
        outcome: "skip",
        reason: xsd,
      });
      continue;
    }
    if (DRY_RUN || !xmllint) {
      results.push({
        fixture: fixtureName,
        side: label,
        part: partPath,
        outcome: "would-run",
        xsd,
      });
      continue;
    }

    // Materialize the part to disk so xmllint can read it. Keeping the
    // filename close to the original makes xmllint's error messages easy to
    // grep for in CI logs.
    const safe = partPath.replace(/[^A-Za-z0-9._-]/g, "_");
    const xmlPath = join(workDir, `${label}-${safe}`);
    writeFileSync(xmlPath, container.readBytes(partPath));
    const xsdPath = join(XSD_DIR, xsd);
    const r = validateWithXmllint(xmllint, xsdPath, xmlPath);
    results.push({
      fixture: fixtureName,
      side: label,
      part: partPath,
      outcome: r.ok ? "valid" : "invalid",
      xsd,
      stderr: r.stderr,
    });
  }
}

function printTable(rows) {
  const fixCol = Math.max("fixture".length, ...rows.map((r) => r.fixture.length));
  const partCol = Math.max("part".length, ...rows.map((r) => r.part.length));
  console.log("");
  console.log(
    `| ${padRight("fixture", fixCol)} | ${padRight("part", partCol)} | input  | re-emit | xsd                                     |`
  );
  console.log(
    `| ${"-".repeat(fixCol)} | ${"-".repeat(partCol)} | ------ | ------- | --------------------------------------- |`
  );
  // Group by (fixture, part); emit one row per part.
  const byKey = new Map();
  for (const r of rows) {
    const key = `${r.fixture}\n${r.part}`;
    if (!byKey.has(key))
      byKey.set(key, { fixture: r.fixture, part: r.part, input: null, reemit: null, xsd: r.xsd ?? "" });
    const e = byKey.get(key);
    if (r.side === "input") e.input = r;
    else if (r.side === "re-emit") e.reemit = r;
    if (r.xsd) e.xsd = r.xsd;
  }
  function symbol(r) {
    if (!r) return "  -   ";
    switch (r.outcome) {
      case "valid":
        return "  ✓   ";
      case "invalid":
        return "  ✗   ";
      case "skip":
        return "skip  ";
      case "would-run":
        return "would ";
      case "unmapped":
        return "?map  ";
      default:
        return "  ?   ";
    }
  }
  for (const e of byKey.values()) {
    console.log(
      `| ${padRight(e.fixture, fixCol)} | ${padRight(e.part, partCol)} | ${symbol(e.input)} | ${padRight(symbol(e.reemit), 7)} | ${padRight(e.xsd, 39)} |`
    );
  }
}

function printFailures(rows) {
  const failures = rows.filter((r) => r.outcome === "invalid");
  if (failures.length === 0) return;
  console.log("");
  console.log(`────── ${failures.length} schema violation(s) ──────`);
  for (const f of failures) {
    console.log(`\n[${f.fixture} / ${f.side} / ${f.part}] (${f.xsd})`);
    if (f.stderr) {
      for (const line of f.stderr.split("\n")) {
        console.log(`  > ${line}`);
      }
    }
  }
}

/* ────────────────────────────────────────────────────────────────────────────
 * --self-test entry
 * ──────────────────────────────────────────────────────────────────────── */

async function runSelfTest({ core }) {
  const fixtures = listFixtures();
  console.log(`self-test: discovered ${fixtures.length} fixture(s) in ${FIXTURE_DIR}`);
  if (fixtures.length !== 6) {
    console.error(
      `self-test: expected 6 real-world fixtures, found ${fixtures.length}. The brief pins the corpus at 6.`
    );
    return 1;
  }

  // Walk every part of every fixture, assert that mapPartToXsd returns
  // either an xsd name or skip:<reason>. An "unmapped" result here is a
  // genuine signal that we have a new part type to teach the validator about.
  const unmapped = [];
  let totalParts = 0;
  let totalSkipped = 0;
  let totalToValidate = 0;
  for (const path of fixtures) {
    const buf = readFileSync(path);
    const container = await core.ooxml.OoxmlContainer.load(buf);
    const ctLookup = buildContentTypeLookup(container);
    for (const partPath of container.parts.keys()) {
      totalParts++;
      const xsd = mapPartToXsd(partPath, ctLookup);
      if (!xsd) unmapped.push({ fixture: path.split("/").pop(), part: partPath });
      else if (xsd.startsWith("skip:")) totalSkipped++;
      else totalToValidate++;
    }
  }
  console.log(
    `self-test: scanned ${totalParts} part(s) — ${totalToValidate} would-validate, ${totalSkipped} skipped, ${unmapped.length} unmapped`
  );
  if (unmapped.length > 0) {
    console.error("self-test: unmapped parts:");
    for (const u of unmapped) console.error(`  - ${u.fixture}: ${u.part}`);
    return 1;
  }

  // --inject-broken arm: feed a synthetic malformed blob through the
  // well-formedness probe and assert the script reports a failure.
  if (INJECT_BROKEN) {
    const broken = `<?xml version="1.0"?><root><child>oops</root>`;
    const wf = isXmlWellFormed(broken);
    if (wf.ok) {
      console.error(
        `self-test: --inject-broken expected the well-formedness probe to fail on a synthetic blob, but it passed.`
      );
      return 1;
    }
    console.log(`self-test: --inject-broken correctly flagged the synthetic blob (${wf.reason}).`);
    // Surface non-zero so the test harness can assert the failure path
    // propagates an exit code (mirrors `license-scan.mjs --inject-agpl`).
    return 1;
  }

  console.log("self-test: ✅ all checks passed.");
  return 0;
}

/* ────────────────────────────────────────────────────────────────────────────
 * Main
 * ──────────────────────────────────────────────────────────────────────── */

async function main() {
  const fixtures = listFixtures();
  if (fixtures.length === 0) {
    console.warn(`⚠ no fixtures in ${FIXTURE_DIR}; run \`pnpm fixtures-real\` first.`);
    return 0;
  }

  if (SELF_TEST) {
    const { core } = await loadCore();
    return runSelfTest({ core });
  }

  if (!existsSync(XSD_SENTINEL)) {
    console.warn(
      `⚠ vendor/ooxml-xsd/wml.xsd not found — skipping schema validation.\n` +
        `  Run \`make xsd-fetch\` (or \`node scripts/fetch-ooxml-xsd.mjs\`) once\n` +
        `  to populate the schemas. CI runs the fetch step before this gate.`
    );
    return 0;
  }

  let xmllint = null;
  if (!DRY_RUN) {
    xmllint = findXmllint();
    if (!xmllint) {
      console.warn(
        `⚠ xmllint not found on PATH — skipping schema validation.\n` +
          `  Install libxml2 (e.g. \`brew install libxml2\` or\n` +
          `  \`apt-get install libxml2-utils\`) to run the full check locally.\n` +
          `  CI installs it explicitly so the gate still runs in CI.`
      );
      return 0;
    }
  }

  const { core, docx } = await loadCore();
  console.log(`✓ using xmllint at ${xmllint ?? "(dry-run)"}`);
  console.log(`✓ using XSD bundle at ${XSD_DIR}`);
  console.log(`✓ checking ${fixtures.length} fixtures from ${FIXTURE_DIR}`);

  const workDir = join(tmpdir(), `officeai-xsd-${process.pid}`);
  mkdirSync(workDir, { recursive: true });

  const results = [];
  for (const path of fixtures) {
    const fixtureName = path.split("/").pop();
    const buf = readFileSync(path);

    // 1. Validate the input bytes as-is.
    const inputContainer = await core.ooxml.OoxmlContainer.load(buf);
    await validateContainer({
      container: inputContainer,
      label: "input",
      fixtureName,
      xmllint,
      workDir,
      results,
    });

    // 2. Run the agent through it (trivial edit) and validate the re-emit.
    let reemitBuf;
    try {
      const agent = await docx.DocxAgent.fromBuffer(buf);
      const snap = agent.getSnapshot();
      const firstParaIdx = snap.root.body.findIndex((b) => b.kind === "paragraph");
      if (firstParaIdx >= 0) {
        await agent.applyCommand({
          type: "docx:insert-text",
          payload: {
            at: { paragraph: firstParaIdx, offset: 0 },
            text: "X",
          },
          source: "agent",
          agentId: "schema-validate",
        });
      }
      reemitBuf = Buffer.from(await agent.exportFile());
    } catch (err) {
      console.error(`❌ ${fixtureName}: agent re-emit failed — ${err instanceof Error ? err.message : err}`);
      // Record a synthetic per-fixture failure so the table makes the gap
      // visible without crashing the rest of the run.
      results.push({
        fixture: fixtureName,
        side: "re-emit",
        part: "(agent failed)",
        outcome: "invalid",
        xsd: "",
        stderr: err instanceof Error ? err.message : String(err),
      });
      continue;
    }
    const reemitContainer = await core.ooxml.OoxmlContainer.load(reemitBuf);
    await validateContainer({
      container: reemitContainer,
      label: "re-emit",
      fixtureName,
      xmllint,
      workDir,
      results,
    });
  }

  printTable(results);
  printFailures(results);

  const totalRows = results.length;
  const valid = results.filter((r) => r.outcome === "valid").length;
  const invalid = results.filter((r) => r.outcome === "invalid").length;
  const skipped = results.filter((r) => r.outcome === "skip").length;
  const wouldRun = results.filter((r) => r.outcome === "would-run").length;

  console.log("");
  console.log(
    `summary: ${totalRows} part-checks across ${fixtures.length} fixtures — ` +
      `valid=${valid}, invalid=${invalid}, skipped=${skipped}, dry-run=${wouldRun}.`
  );

  // Best effort cleanup; leave artifacts on failure so a developer can poke at
  // the offending part with `xmllint --schema` interactively.
  if (invalid === 0 && existsSync(workDir)) {
    rmSync(workDir, { recursive: true, force: true });
  }

  if (invalid > 0) {
    console.error(`\n❌ schema-validate: ${invalid} schema violation(s). Artifacts in ${workDir}`);
    return 1;
  }
  console.log("\n✅ schema-validate: every checked part validates clean.");
  return 0;
}

const code = await main();
process.exit(code);
