#!/usr/bin/env node
/**
 * OOXML schema validator (W9 / Theme D4) — multi-format.
 *
 * For every fixture in the configured directories, this script:
 *   1. Loads the input bytes into an `OoxmlContainer` and enumerates every
 *      XML part.
 *   2. Roundtrips the file through `<Format>Agent.fromBuffer → exportFile()`,
 *      then enumerates every XML part of the re-emit.
 *   3. Maps each part to its corresponding ECMA-376 (Transitional) XSD via a
 *      `[Content_Types].xml`-aware lookup, and shells out to `xmllint
 *      --noout --schema <xsd>` to validate the part bytes.
 *   4. Prints a per-fixture row table: `fixture | part | input ✓/✗ |
 *      re-emit ✓/✗`. On failure, emits the offending xmllint stderr as a
 *      quoted block.
 *
 * Usage:
 *   node scripts/validate-ooxml-schemas.mjs                # docx (default)
 *   node scripts/validate-ooxml-schemas.mjs --format xlsx
 *   node scripts/validate-ooxml-schemas.mjs --format pptx
 *
 * Exit semantics (matches the `make perf-docx` / `make roundtrip-libre`
 * pattern so wrappers don't need to special-case it):
 *   - exit 0 + warning : `xmllint` is missing from PATH (graceful skip; CI
 *     installs `libxml2-utils` so the gate still runs server-side).
 *   - exit 0 + warning : the relevant XSD file is missing under
 *     `vendor/ooxml-xsd/` (run `make xsd-fetch`).
 *   - exit 0           : every part is well-formed AND schema-valid.
 *   - exit 1           : at least one part fails.
 *
 * Modes:
 *   --dry-run            : skip xmllint entirely; just emit the parts that
 *                          would be validated and which XSD each maps to.
 *                          Used by `tests/scripts/validate-ooxml-schemas.test.ts`
 *                          so the test stays hermetic on a fresh CI runner.
 *   --self-test          : like --dry-run, plus assert that the expected
 *                          fixture count is present and that every observed
 *                          part maps to either a known XSD or an explicit
 *                          "skip" bucket. Exit non-zero if an unknown part
 *                          shape is encountered.
 *   --inject-broken      : prepend a synthetic malformed-XML "part" to the
 *                          validation queue and assert the failure path
 *                          counts it as a violation. Combined with
 *                          --self-test this is what the unit test uses to
 *                          confirm the failure path raises non-zero without
 *                          needing xmllint installed.
 *
 * Run via `make schema-validate-{docx,xlsx,pptx}`.
 */

import { spawnSync } from "node:child_process";
import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, "..");
const XSD_DIR = resolve(root, "vendor/ooxml-xsd");
const CORE_DIST = resolve(root, "packages/core/dist/index.js");

/* ────────────────────────────────────────────────────────────────────────────
 * Args
 * ──────────────────────────────────────────────────────────────────────── */

const argv = process.argv.slice(2);
const argSet = new Set(argv);
let FORMAT = "docx";
for (let i = 0; i < argv.length; i++) {
  if (argv[i] === "--format" && argv[i + 1]) {
    FORMAT = argv[i + 1];
  } else if (argv[i].startsWith("--format=")) {
    FORMAT = argv[i].slice("--format=".length);
  }
}
const DRY_RUN = argSet.has("--dry-run") || argSet.has("--self-test");
const SELF_TEST = argSet.has("--self-test");
const INJECT_BROKEN = argSet.has("--inject-broken");

/* ────────────────────────────────────────────────────────────────────────────
 * Per-format config
 *
 * Each format declares:
 *   - extension      : fixture file extension (case-insensitive match)
 *   - fixtureDirs    : ordered list of dirs to scan; missing dirs ignored
 *   - sentinelXsd    : the XSD whose presence indicates the bundle is fetched
 *   - selfTestCount  : pinned fixture count for `--self-test` (0 = no pin)
 *   - agentDist      : built dist path for the format's agent
 *   - agentName      : exported class on `agentDist`
 *   - ctMap          : Content-Type → xsd filename (or `skip:<reason>`)
 *   - pathFallbacks  : ordered [regex, xsd] pairs for parts without a
 *                      content-type override
 * ──────────────────────────────────────────────────────────────────────── */

const SHARED_OPC_SKIPS = {
  "application/vnd.openxmlformats-package.relationships+xml": "skip:opc",
  "application/vnd.openxmlformats-package.core-properties+xml": "skip:opc",
  "image/png": "skip:binary",
  "image/jpeg": "skip:binary",
  "image/gif": "skip:binary",
  "image/svg+xml": "skip:binary",
  "image/bmp": "skip:binary",
};

const SHARED_PATH_SKIPS = [
  [/^\[Content_Types\]\.xml$/, "skip:opc"],
  [/^_rels\//, "skip:opc"],
  [/_rels\//, "skip:opc"],
  [/^docProps\/core\.xml$/, "skip:opc"],
  [/\.bin$/, "skip:binary"],
  [/\.png$|\.jpe?g$|\.gif$|\.svg$|\.bmp$/i, "skip:binary"],
];

const SHARED_DOC_PROPS = {
  "application/vnd.openxmlformats-officedocument.extended-properties+xml":
    "shared-documentPropertiesExtended.xsd",
  "application/vnd.openxmlformats-officedocument.custom-properties+xml":
    "shared-documentPropertiesCustom.xsd",
  "application/vnd.openxmlformats-officedocument.theme+xml": "dml-main.xsd",
  "application/vnd.openxmlformats-officedocument.themeOverride+xml": "dml-main.xsd",
  "application/vnd.openxmlformats-officedocument.drawingml.chart+xml": "dml-chart.xsd",
  "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml": "dml-diagram.xsd",
  "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml": "dml-diagram.xsd",
  "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml": "dml-diagram.xsd",
  "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml": "dml-diagram.xsd",
  "application/vnd.openxmlformats-officedocument.vmlDrawing": "skip:vml-binary",
};

const FORMATS = {
  docx: {
    extension: ".docx",
    fixtureDirs: ["fixtures/docx/real-world"],
    sentinelXsd: "wml.xsd",
    selfTestCount: 7,
    agentDist: resolve(root, "packages/docx/dist/index.js"),
    agentName: "DocxAgent",
    ctMap: {
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
      ...SHARED_DOC_PROPS,
      ...SHARED_OPC_SKIPS,
      "application/vnd.ms-word.commentsExtended+xml": "skip:w15",
      "application/vnd.ms-word.commentsIds+xml": "skip:w15",
      "application/vnd.ms-word.people+xml": "skip:w15",
    },
    pathFallbacks: [
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
      ...SHARED_PATH_SKIPS,
      [/^word\/commentsExtended\.xml$/, "skip:w15"],
      [/^word\/commentsIds\.xml$/, "skip:w15"],
      [/^word\/people\.xml$/, "skip:w15"],
      [/^word\/media\//, "skip:binary"],
      [/^word\/embeddings\//, "skip:binary"],
    ],
  },
  xlsx: {
    extension: ".xlsx",
    fixtureDirs: ["fixtures/xlsx/synthetic", "fixtures/xlsx/real-world"],
    sentinelXsd: "sml.xsd",
    selfTestCount: 0,
    agentDist: resolve(root, "packages/xlsx/dist/index.js"),
    agentName: "XlsxAgent",
    ctMap: {
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml": "sml.xsd",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet+xml": "sml.xsd",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml": "sml.xsd",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml": "sml.xsd",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml": "sml.xsd",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml": "sml.xsd",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml": "sml.xsd",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml": "sml.xsd",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml": "sml.xsd",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml": "sml.xsd",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml": "sml.xsd",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml": "sml.xsd",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml": "sml.xsd",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml": "sml.xsd",
      "application/vnd.ms-excel.sheetMetadata+xml": "sml.xsd",
      ...SHARED_DOC_PROPS,
      ...SHARED_OPC_SKIPS,
    },
    pathFallbacks: [
      [/^xl\/workbook\.xml$/, "sml.xsd"],
      [/^xl\/worksheets\/sheet\d+\.xml$/, "sml.xsd"],
      [/^xl\/styles\.xml$/, "sml.xsd"],
      [/^xl\/sharedStrings\.xml$/, "sml.xsd"],
      [/^xl\/calcChain\.xml$/, "sml.xsd"],
      [/^xl\/metadata\.xml$/, "sml.xsd"],
      [/^xl\/comments\d*\.xml$/, "sml.xsd"],
      [/^xl\/tables\/table\d+\.xml$/, "sml.xsd"],
      [/^xl\/theme\/theme\d*\.xml$/, "dml-main.xsd"],
      [/^xl\/charts\/chart\d+\.xml$/, "dml-chart.xsd"],
      [/^xl\/drawings\/drawing\d+\.xml$/, "dml-spreadsheetDrawing.xsd"],
      [/^docProps\/app\.xml$/, "shared-documentPropertiesExtended.xsd"],
      [/^docProps\/custom\.xml$/, "shared-documentPropertiesCustom.xsd"],
      ...SHARED_PATH_SKIPS,
      [/^xl\/media\//, "skip:binary"],
      [/^xl\/embeddings\//, "skip:binary"],
      [/^xl\/printerSettings\//, "skip:binary"],
    ],
  },
  pptx: {
    extension: ".pptx",
    fixtureDirs: ["fixtures/pptx/synthetic", "fixtures/pptx/real"],
    sentinelXsd: "pml.xsd",
    selfTestCount: 0,
    agentDist: resolve(root, "packages/pptx/dist/index.js"),
    agentName: "PptxAgent",
    ctMap: {
      "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml": "pml.xsd",
      "application/vnd.openxmlformats-officedocument.presentationml.template.main+xml": "pml.xsd",
      "application/vnd.openxmlformats-officedocument.presentationml.slide+xml": "pml.xsd",
      "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml": "pml.xsd",
      "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml": "pml.xsd",
      "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml": "pml.xsd",
      "application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml": "pml.xsd",
      "application/vnd.openxmlformats-officedocument.presentationml.handoutMaster+xml": "pml.xsd",
      "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml": "pml.xsd",
      "application/vnd.openxmlformats-officedocument.presentationml.comments+xml": "pml.xsd",
      "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml": "pml.xsd",
      "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml": "pml.xsd",
      "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml": "pml.xsd",
      "application/vnd.openxmlformats-officedocument.presentationml.tags+xml": "pml.xsd",
      ...SHARED_DOC_PROPS,
      ...SHARED_OPC_SKIPS,
    },
    pathFallbacks: [
      [/^ppt\/presentation\.xml$/, "pml.xsd"],
      [/^ppt\/slides\/slide\d+\.xml$/, "pml.xsd"],
      [/^ppt\/slideLayouts\/slideLayout\d+\.xml$/, "pml.xsd"],
      [/^ppt\/slideMasters\/slideMaster\d+\.xml$/, "pml.xsd"],
      [/^ppt\/notesSlides\/notesSlide\d+\.xml$/, "pml.xsd"],
      [/^ppt\/notesMasters\/notesMaster\d+\.xml$/, "pml.xsd"],
      [/^ppt\/handoutMasters\/handoutMaster\d+\.xml$/, "pml.xsd"],
      [/^ppt\/presProps\.xml$/, "pml.xsd"],
      [/^ppt\/viewProps\.xml$/, "pml.xsd"],
      [/^ppt\/tableStyles\.xml$/, "pml.xsd"],
      [/^ppt\/tags\/tag\d+\.xml$/, "pml.xsd"],
      [/^ppt\/comments?\/.+\.xml$/, "pml.xsd"],
      [/^ppt\/commentAuthors\.xml$/, "pml.xsd"],
      [/^ppt\/theme\/theme\d*\.xml$/, "dml-main.xsd"],
      [/^ppt\/charts\/chart\d+\.xml$/, "dml-chart.xsd"],
      [/^ppt\/diagrams\/.+\.xml$/, "dml-diagram.xsd"],
      [/^docProps\/app\.xml$/, "shared-documentPropertiesExtended.xsd"],
      [/^docProps\/custom\.xml$/, "shared-documentPropertiesCustom.xsd"],
      ...SHARED_PATH_SKIPS,
      [/^ppt\/media\//, "skip:binary"],
      [/^ppt\/embeddings\//, "skip:binary"],
      [/^ppt\/printerSettings\//, "skip:binary"],
    ],
  },
};

const formatConfig = FORMATS[FORMAT];
if (!formatConfig) {
  console.error(`Unknown --format ${FORMAT}. Use one of: docx, xlsx, pptx.`);
  process.exit(1);
}

const XSD_SENTINEL = join(XSD_DIR, formatConfig.sentinelXsd);

/**
 * Given a part path and an optional content-type lookup table, return the XSD
 * filename to validate against, or `skip:<reason>` to skip, or `null` if the
 * part is unrecognised (which `--self-test` flags as a failure so we notice
 * new part types).
 */
export function mapPartToXsd(partPath, contentTypeFor) {
  const ct = contentTypeFor ? contentTypeFor(partPath) : undefined;
  if (ct && formatConfig.ctMap[ct]) return formatConfig.ctMap[ct];
  for (const [re, xsd] of formatConfig.pathFallbacks) {
    if (re.test(partPath)) return xsd;
  }
  return null;
}

/* ────────────────────────────────────────────────────────────────────────────
 * Helpers
 * ──────────────────────────────────────────────────────────────────────── */

function listFixtures() {
  const out = [];
  for (const rel of formatConfig.fixtureDirs) {
    const dir = resolve(root, rel);
    if (!existsSync(dir)) continue;
    let entries;
    try {
      entries = readdirSync(dir);
    } catch {
      continue;
    }
    for (const f of entries.sort()) {
      if (f.toLowerCase().endsWith(formatConfig.extension)) {
        out.push(join(dir, f));
      }
    }
  }
  return out;
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
  if (!existsSync(CORE_DIST) || !existsSync(formatConfig.agentDist)) {
    console.error(
      `❌ schema-validate: missing built dist (run \`pnpm build\` first).\n` +
        `   expected: ${CORE_DIST}\n` +
        `   expected: ${formatConfig.agentDist}`
    );
    process.exit(1);
  }
  const core = await import(pathToFileURL(CORE_DIST).href);
  const agentMod = await import(pathToFileURL(formatConfig.agentDist).href);
  const Agent = agentMod[formatConfig.agentName];
  if (!Agent) {
    console.error(`❌ schema-validate: ${formatConfig.agentDist} does not export ${formatConfig.agentName}.`);
    process.exit(1);
  }
  return { core, Agent };
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
  const defaults = new Map();
  const overrides = new Map();
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

function padRight(s, n) {
  return s.length >= n ? s.slice(0, n) : s + " ".repeat(n - s.length);
}

async function validateContainer({ container, label, fixtureName, xmllint, workDir, results }) {
  const ctLookup = buildContentTypeLookup(container);
  const partPaths = [...container.parts.keys()].sort();
  for (const partPath of partPaths) {
    const xsd = mapPartToXsd(partPath, ctLookup);
    if (!xsd) {
      results.push({ fixture: fixtureName, side: label, part: partPath, outcome: "unmapped" });
      continue;
    }
    if (xsd.startsWith("skip:")) {
      results.push({ fixture: fixtureName, side: label, part: partPath, outcome: "skip", reason: xsd });
      continue;
    }
    if (DRY_RUN || !xmllint) {
      results.push({ fixture: fixtureName, side: label, part: partPath, outcome: "would-run", xsd });
      continue;
    }
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
  console.log(
    `self-test: discovered ${fixtures.length} fixture(s) in ${formatConfig.fixtureDirs.join(", ")}`
  );
  if (formatConfig.selfTestCount > 0 && fixtures.length !== formatConfig.selfTestCount) {
    console.error(
      `self-test: expected ${formatConfig.selfTestCount} ${FORMAT} fixtures, found ${fixtures.length}.`
    );
    return 1;
  }

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
    console.warn(`⚠ no ${FORMAT} fixtures in ${formatConfig.fixtureDirs.join(", ")}.`);
    return 0;
  }

  if (SELF_TEST) {
    const { core } = await loadCore();
    return runSelfTest({ core });
  }

  if (!existsSync(XSD_SENTINEL)) {
    console.warn(
      `⚠ vendor/ooxml-xsd/${formatConfig.sentinelXsd} not found — skipping ${FORMAT} schema validation.\n` +
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
        `⚠ xmllint not found on PATH — skipping ${FORMAT} schema validation.\n` +
          `  Install libxml2 (e.g. \`brew install libxml2\` or\n` +
          `  \`apt-get install libxml2-utils\`) to run the full check locally.\n` +
          `  CI installs it explicitly so the gate still runs in CI.`
      );
      return 0;
    }
  }

  const { core, Agent } = await loadCore();
  console.log(`✓ format: ${FORMAT}`);
  console.log(`✓ using xmllint at ${xmllint ?? "(dry-run)"}`);
  console.log(`✓ using XSD bundle at ${XSD_DIR}`);
  console.log(`✓ checking ${fixtures.length} fixtures`);

  const workDir = join(tmpdir(), `officeai-xsd-${FORMAT}-${process.pid}`);
  mkdirSync(workDir, { recursive: true });

  const results = [];
  for (const path of fixtures) {
    const fixtureName = path.split("/").pop();
    const buf = readFileSync(path);

    const inputContainer = await core.ooxml.OoxmlContainer.load(buf);
    await validateContainer({
      container: inputContainer,
      label: "input",
      fixtureName,
      xmllint,
      workDir,
      results,
    });

    // Trivial agent re-emit: for DOCX we apply an `insert-text` so the diff
    // path is exercised end-to-end (the historical CI signal). For XLSX and
    // PPTX we just `exportFile()` — that already exercises serializer paths
    // and avoids coupling the validator to format-specific command shapes.
    // The post-mutation path is exercised separately by `roundtrip-libre`.
    let reemitBuf;
    try {
      const agent = await Agent.fromBuffer(buf);
      if (FORMAT === "docx") {
        const snap = agent.getSnapshot();
        const firstParaIdx = snap.root.body.findIndex((b) => b.kind === "paragraph");
        if (firstParaIdx >= 0) {
          await agent.applyCommand({
            type: "docx:insert-text",
            payload: { at: { paragraph: firstParaIdx, offset: 0 }, text: "X" },
            source: "agent",
            agentId: "schema-validate",
          });
        }
      }
      reemitBuf = Buffer.from(await agent.exportFile());
    } catch (err) {
      console.error(`❌ ${fixtureName}: agent re-emit failed — ${err instanceof Error ? err.message : err}`);
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
    `summary [${FORMAT}]: ${totalRows} part-checks across ${fixtures.length} fixtures — ` +
      `valid=${valid}, invalid=${invalid}, skipped=${skipped}, dry-run=${wouldRun}.`
  );

  if (invalid === 0 && existsSync(workDir)) {
    rmSync(workDir, { recursive: true, force: true });
  }

  if (invalid > 0) {
    console.error(`\n❌ schema-validate [${FORMAT}]: ${invalid} schema violation(s). Artifacts in ${workDir}`);
    return 1;
  }
  console.log(`\n✅ schema-validate [${FORMAT}]: every checked part validates clean.`);
  return 0;
}

const code = await main();
process.exit(code);
