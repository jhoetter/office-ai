#!/usr/bin/env node
/**
 * Matrix-driven roundtrip release gate.
 *
 * Fast, dependency-light gate for every product-relevant fixture in
 * fixtures/MATRIX.json:
 *   import -> project -> no-op or mutation -> export -> validate -> reimport.
 *
 * On failure it writes CI-ready artifacts under .tmp/roundtrip-gate:
 * input, output when available, diagnostics.json and diff-hint.txt.
 * Heavy external compatibility checks stay in run-libreoffice-roundtrip.mjs.
 */

import { createHash } from "node:crypto";
import { existsSync, mkdirSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { basename, dirname, join, resolve } from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";
import JSZip from "jszip";

const here = dirname(fileURLToPath(import.meta.url));
const root = resolve(here, "..");
const matrixPath = resolve(root, "fixtures/MATRIX.json");
const defaultOutDir = resolve(root, ".tmp/roundtrip-gate");

const FORMATS = {
  docx: {
    extension: ".docx",
    agentEntry: "packages/docx/dist/index.js",
    agentName: "DocxAgent",
    mutator: async (agent) =>
      agent.applyCommand({
        type: "docx:insert-text",
        payload: { at: { paragraph: 0, offset: 0 }, text: "ROUNDTRIP_GATE " },
      }),
  },
  xlsx: {
    extension: ".xlsx",
    agentEntry: "packages/xlsx/dist/index.js",
    agentName: "XlsxAgent",
    mutator: async (agent) =>
      agent.applyCommand({
        type: "xlsx:set-cell-value",
        payload: { sheet: 0, ref: "A1", value: "ROUNDTRIP_GATE" },
      }),
  },
  pptx: {
    extension: ".pptx",
    agentEntry: "packages/pptx/dist/index.js",
    agentName: "PptxAgent",
    mutator: async (agent) =>
      agent.applyCommand({
        type: "pptx:set-slide-notes",
        payload: { slideIndex: 0, text: "ROUNDTRIP_GATE speaker notes" },
      }),
  },
  pdf: {
    extension: ".pdf",
    agentEntry: "packages/pdf/dist/index.js",
    agentName: "PdfAgent",
    mutator: async (agent) =>
      agent.applyCommand({
        type: "pdf:rotate-pages",
        payload: { pages: [1], delta: 90 },
      }),
  },
};

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const outDir = resolve(root, args.outDir ?? defaultOutDir);
  rmSync(outDir, { recursive: true, force: true });
  mkdirSync(outDir, { recursive: true });

  const matrix = JSON.parse(readFileSync(matrixPath, "utf8"));
  const formats = args.formats ?? Object.keys(FORMATS);
  const modes = args.mode === "all" ? ["noop", "mutation"] : [args.mode];
  const results = [];

  for (const format of formats) {
    if (!FORMATS[format]) {
      console.error(`roundtrip-gate: unknown format "${format}"`);
      process.exit(2);
    }
    const Agent = await loadAgent(format);
    for (const mode of modes) {
      const behavior = mode === "noop" ? "noop-roundtrip" : "mutation-roundtrip";
      const fixtures = matrix.fixtures
        .filter((fixture) => fixture.format === format && fixture.expectedBehaviors.includes(behavior))
        .slice(0, args.limit ?? undefined);
      for (const fixture of fixtures) {
        const row = await runFixture({ Agent, fixture, format, mode, outDir });
        results.push(row);
        console.log(formatRow(row));
      }
    }
  }

  let negativeSelfTest = null;
  if (args.includeNegativeSelfTest) {
    negativeSelfTest = await runNegativeSelfTest(outDir);
    console.log(
      negativeSelfTest.ok
        ? "✓ negative self-test: corrupt fixture failed as expected"
        : "✗ negative self-test: corrupt fixture unexpectedly passed"
    );
  }

  const summary = {
    schema: "office-ai/roundtrip-gate@1",
    generatedAt: new Date().toISOString(),
    modes,
    formats,
    totals: {
      fixtures: results.length,
      passed: results.filter((row) => row.ok).length,
      failed: results.filter((row) => !row.ok).length,
    },
    negativeSelfTest,
    results: results.map(summarizeResult),
  };
  writeFileSync(join(outDir, "summary.json"), JSON.stringify(summary, null, 2));

  const failed = results.filter((row) => !row.ok);
  if (failed.length > 0 || negativeSelfTest?.ok === false) {
    console.error(
      `roundtrip-gate: FAILED (${failed.length} fixture failure(s)). Artifacts: ${relativeOut(outDir)}`
    );
    process.exit(1);
  }

  console.log(
    `roundtrip-gate: OK (${summary.totals.passed}/${summary.totals.fixtures} fixture checks, artifacts: ${relativeOut(
      outDir
    )})`
  );
}

async function runFixture({ Agent, fixture, format, mode, outDir }) {
  const config = FORMATS[format];
  const absPath = resolve(root, fixture.path);
  const diagnostics = [];
  let input = null;
  let output = null;
  let diffHint = "";
  try {
    input = readFileSync(absPath);
    diagnostics.push(info("import", "input-read", `${fixture.path} (${input.byteLength} bytes)`));
    const before = await Agent.fromBuffer(input);
    diagnostics.push(info("import", "agent-imported", `${format} fixture imported`));
    const projectionBefore = project(before);
    diagnostics.push(
      info(
        "projection",
        "projection-created",
        `projection bytes=${Buffer.byteLength(projectionBefore, "utf8")}`
      )
    );

    if (mode === "mutation") {
      if (typeof config.mutator !== "function") {
        throw new Error(`${format} has mutation-roundtrip fixtures but no mutator`);
      }
      const mutation = await config.mutator(before);
      diagnostics.push(
        info(
          "mutation",
          "mutation-applied",
          `${mutation?.diff?.changes?.length ?? 0} diff change(s) from ${mutation?.command?.type ?? "command"}`
        )
      );
    } else {
      diagnostics.push(info("noop", "noop-export", "no mutation applied before export"));
    }

    output = Buffer.from(await before.exportFile());
    diagnostics.push(info("export", "exported", `${output.byteLength} bytes`));
    await validateExport(format, output);
    diagnostics.push(info("validate", "export-valid", `${format} export validates`));
    const after = await Agent.fromBuffer(output);
    diagnostics.push(info("reimport", "agent-reimported", `${format} export reimported`));
    const projectionAfter = project(after);
    diagnostics.push(
      info(
        "projection",
        "reimport-projection-created",
        `projection bytes=${Buffer.byteLength(projectionAfter, "utf8")}`
      )
    );
    diagnostics.push(...formatDiagnostics(format, after));
    if (fixture.knownRisks.length > 0) {
      diagnostics.push(info("fixture-risk", "known-risk-covered", fixture.knownRisks.join("; ")));
    }
    diffHint = buildDiffHint({ mode, projectionBefore, projectionAfter, input, output });
    return { ok: true, fixture, format, mode, diagnostics, diffHint };
  } catch (err) {
    diagnostics.push(error("gate", "roundtrip-failed", err));
    diffHint =
      diffHint ||
      buildDiffHint({
        mode,
        projectionBefore: "",
        projectionAfter: "",
        input: input ?? Buffer.alloc(0),
        output,
      });
    const artifactDir = writeFailureArtifacts({
      outDir,
      fixture,
      format,
      mode,
      input,
      output,
      diagnostics,
      diffHint,
    });
    return {
      ok: false,
      fixture,
      format,
      mode,
      error: String(err?.message ?? err),
      diagnostics,
      diffHint,
      artifactDir: relativeOut(artifactDir),
    };
  }
}

async function runNegativeSelfTest(outDir) {
  const format = "docx";
  const Agent = await loadAgent(format);
  const fixture = {
    id: "negative.self-test.corrupt-docx",
    format,
    path: "<generated-corrupt-docx>",
    knownRisks: ["negative self-test must fail import"],
  };
  const input = Buffer.from("not a valid office document");
  const diagnostics = [];
  try {
    await Agent.fromBuffer(input);
    const diffHint = "Corrupt DOCX unexpectedly imported. The roundtrip gate is not detecting broken inputs.";
    const artifactDir = writeFailureArtifacts({
      outDir,
      fixture,
      format,
      mode: "negative-self-test",
      input,
      output: null,
      diagnostics: [error("negative-self-test", "unexpected-pass", diffHint)],
      diffHint,
    });
    return { ok: false, expectedFailureObserved: false, artifactDir: relativeOut(artifactDir) };
  } catch (err) {
    diagnostics.push(info("negative-self-test", "expected-failure", String(err?.message ?? err)));
    const artifactDir = writeFailureArtifacts({
      outDir,
      fixture,
      format,
      mode: "negative-self-test",
      input,
      output: null,
      diagnostics,
      diffHint: "Corrupt DOCX failed during import as expected.",
      expectedFailure: true,
    });
    return { ok: true, expectedFailureObserved: true, artifactDir: relativeOut(artifactDir) };
  }
}

async function validateExport(format, output) {
  if (format === "pdf") {
    if (output.length < 5 || output.slice(0, 5).toString("ascii") !== "%PDF-") {
      throw new Error("export does not start with %PDF-");
    }
    return;
  }
  const zip = await JSZip.loadAsync(output);
  if (!zip.file("[Content_Types].xml")) {
    throw new Error("OOXML export is missing [Content_Types].xml");
  }
  const expectedRoot = format === "docx" ? "word/" : format === "xlsx" ? "xl/" : "ppt/";
  if (!Object.keys(zip.files).some((name) => name.startsWith(expectedRoot))) {
    throw new Error(`OOXML export is missing ${expectedRoot} parts`);
  }
}

function project(agent) {
  if (typeof agent.toMarkdown !== "function") {
    throw new Error("agent does not expose toMarkdown()");
  }
  const md = agent.toMarkdown();
  if (typeof md !== "string") {
    throw new Error("projection is not a string");
  }
  return md;
}

function formatDiagnostics(format, agent) {
  const snap = agent.getSnapshot?.();
  if (!snap?.root) return [];
  if (format === "pdf") {
    const pages = snap.root.pages ?? [];
    return [
      info("pdf", "page-count", `${pages.length} page(s)`),
      info(
        "pdf",
        "text-layer-count",
        `${pages.filter((page) => page.hasTextLayer).length} page(s) with text layer`
      ),
      info("pdf", "annotation-count", `${snap.root.annotations?.length ?? 0} annotation(s)`),
      info("pdf", "signature-count", `${snap.root.signatureCount ?? 0} signature(s)`),
    ];
  }
  return [info(format, "snapshot-revision", `revision=${snap.revision ?? "unknown"}`)];
}

function buildDiffHint({ mode, projectionBefore, projectionAfter, input, output }) {
  const beforeHash = hash(projectionBefore);
  const afterHash = hash(projectionAfter);
  return [
    `mode: ${mode}`,
    `input-sha256: ${hash(input)}`,
    `output-sha256: ${output ? hash(output) : "<none>"}`,
    `projection-before-sha256: ${beforeHash}`,
    `projection-after-sha256: ${afterHash}`,
    `projection-changed: ${String(beforeHash !== afterHash)}`,
    `projection-before-bytes: ${Buffer.byteLength(projectionBefore, "utf8")}`,
    `projection-after-bytes: ${Buffer.byteLength(projectionAfter, "utf8")}`,
  ].join("\n");
}

function writeFailureArtifacts({
  outDir,
  fixture,
  format,
  mode,
  input,
  output,
  diagnostics,
  diffHint,
  expectedFailure = false,
}) {
  const dir = join(outDir, "failures", slug(`${mode}-${fixture.id}`));
  mkdirSync(dir, { recursive: true });
  if (input) writeFileSync(join(dir, `input${FORMATS[format]?.extension ?? ".bin"}`), input);
  if (output) writeFileSync(join(dir, `output${FORMATS[format]?.extension ?? ".bin"}`), output);
  writeFileSync(
    join(dir, "diagnostics.json"),
    JSON.stringify(
      {
        schema: "office-ai/roundtrip-diagnostics@1",
        expectedFailure,
        fixture,
        mode,
        diagnostics,
      },
      null,
      2
    )
  );
  writeFileSync(join(dir, "diff-hint.txt"), `${diffHint}\n`);
  return dir;
}

async function loadAgent(format) {
  const config = FORMATS[format];
  const distEntry = resolve(root, config.agentEntry);
  if (!existsSync(distEntry)) {
    throw new Error(`Missing ${format} dist at ${config.agentEntry}; run pnpm build first.`);
  }
  const mod = await import(pathToFileURL(distEntry).href);
  const Agent = mod[config.agentName];
  if (!Agent) throw new Error(`${config.agentEntry} does not export ${config.agentName}`);
  return Agent;
}

function parseArgs(argv) {
  const args = { mode: "all", includeNegativeSelfTest: false };
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if ((arg === "--format" || arg === "--formats") && argv[i + 1]) {
      args.formats = argv[++i]
        .split(",")
        .map((s) => s.trim())
        .filter(Boolean);
    } else if (arg.startsWith("--format=")) {
      args.formats = arg
        .slice("--format=".length)
        .split(",")
        .map((s) => s.trim())
        .filter(Boolean);
    } else if (arg.startsWith("--formats=")) {
      args.formats = arg
        .slice("--formats=".length)
        .split(",")
        .map((s) => s.trim())
        .filter(Boolean);
    } else if (arg === "--mode" && argv[i + 1]) {
      args.mode = argv[++i];
    } else if (arg.startsWith("--mode=")) {
      args.mode = arg.slice("--mode=".length);
    } else if (arg === "--limit" && argv[i + 1]) {
      args.limit = Number(argv[++i]);
    } else if (arg.startsWith("--limit=")) {
      args.limit = Number(arg.slice("--limit=".length));
    } else if (arg === "--out-dir" && argv[i + 1]) {
      args.outDir = argv[++i];
    } else if (arg.startsWith("--out-dir=")) {
      args.outDir = arg.slice("--out-dir=".length);
    } else if (arg === "--include-negative-self-test") {
      args.includeNegativeSelfTest = true;
    } else if (arg === "--help" || arg === "-h") {
      printHelp();
      process.exit(0);
    } else {
      console.error(`roundtrip-gate: unknown argument ${arg}`);
      process.exit(2);
    }
  }
  if (!["noop", "mutation", "all"].includes(args.mode)) {
    console.error("roundtrip-gate: --mode must be noop, mutation or all");
    process.exit(2);
  }
  if (args.limit !== undefined && (!Number.isInteger(args.limit) || args.limit < 1)) {
    console.error("roundtrip-gate: --limit must be a positive integer");
    process.exit(2);
  }
  return args;
}

function printHelp() {
  console.log(`Usage: node scripts/roundtrip-gate.mjs [options]

Options:
  --formats docx,xlsx,pptx,pdf    Formats to check (default: all)
  --mode noop|mutation|all        Roundtrip modes to check (default: all)
  --limit N                       Limit fixtures per format+mode
  --out-dir DIR                   Artifact directory (default: .tmp/roundtrip-gate)
  --include-negative-self-test    Prove corrupt input fails and writes artifacts
`);
}

function summarizeResult(row) {
  return {
    ok: row.ok,
    id: row.fixture.id,
    format: row.format,
    mode: row.mode,
    path: row.fixture.path,
    ...(row.error ? { error: row.error } : {}),
    ...(row.artifactDir ? { artifactDir: row.artifactDir } : {}),
  };
}

function formatRow(row) {
  const rel = row.fixture.path.startsWith("<") ? row.fixture.path : row.fixture.path;
  return row.ok
    ? `✓ ${row.format}:${row.mode} ${row.fixture.id} (${rel})`
    : `✗ ${row.format}:${row.mode} ${row.fixture.id} (${rel}) — ${row.error}`;
}

function info(stage, code, message) {
  return { level: "info", stage, code, message };
}

function error(stage, code, err) {
  return { level: "error", stage, code, message: String(err?.message ?? err) };
}

function hash(value) {
  return createHash("sha256").update(value).digest("hex");
}

function slug(value) {
  return value
    .replace(/[^a-z0-9._-]+/gi, "-")
    .replace(/^-+|-+$/g, "")
    .slice(0, 120);
}

function relativeOut(absPath) {
  return absPath.replace(root + "/", "");
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
