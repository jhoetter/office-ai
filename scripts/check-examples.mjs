#!/usr/bin/env node
import { existsSync, readdirSync, readFileSync, statSync } from "node:fs";
import { extname, join, relative, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = resolve(fileURLToPath(import.meta.url), "..", "..");
const EXAMPLES_DIR = join(ROOT, "examples");
const GENERATED_PREFIX = "examples/_generated/";

const KNOWN_MCP_TOOLS = new Set([
  "approve_change",
  "apply_command",
  "create_deliverable_from_synthesis",
  "create_document",
  "create_session",
  "export_document",
  "get_document",
  "get_document_projection",
  "import_document",
  "list_activity",
  "list_deliverable_templates",
  "list_pending_changes",
  "pdf_document_diagnostics",
  "plan_command",
  "preview_command",
  "reject_change",
]);

const SECRET_PATTERNS = [
  /-----BEGIN (?:RSA |EC |OPENSSH )?PRIVATE KEY-----/,
  /\b(?:sk_live|sk_test|ghp|github_pat|xoxb|AKIA)[A-Za-z0-9_-]{12,}\b/,
  /\b(?:password|secret|token|api[_-]?key)\s*[:=]\s*["']?[A-Za-z0-9_./+=-]{8,}/i,
  /(?:^|["'\s])(?:\/home\/|\/Users\/|C:\\Users\\|~\/\.ssh\/)/,
];

const errors = [];
const coverage = {
  formats: new Set(),
  hasMcpWeb: false,
  hasDiagnostics: false,
  hasDeliverable: false,
  hasCli: false,
};

function readJson(path) {
  try {
    return JSON.parse(readFileSync(path, "utf8"));
  } catch (err) {
    errors.push(`${rel(path)} is not valid JSON: ${err instanceof Error ? err.message : String(err)}`);
    return null;
  }
}

function rel(path) {
  return relative(ROOT, path).replace(/\\/g, "/");
}

function requireFile(path, label) {
  if (!existsSync(path) || !statSync(path).isFile()) {
    errors.push(`${label} missing: ${rel(path)}`);
    return false;
  }
  return true;
}

function isNonEmptyString(value) {
  return typeof value === "string" && value.trim().length > 0;
}

function checkManifest(exampleDir) {
  const manifestPath = join(exampleDir, "manifest.json");
  if (!requireFile(manifestPath, "manifest")) return;
  const manifest = readJson(manifestPath);
  if (!manifest) return;

  const id = exampleDir.split(/[\\/]/).pop();
  if (manifest.schema !== "office-ai/example-manifest@1") {
    errors.push(`${rel(manifestPath)} has unsupported schema`);
  }
  if (manifest.id !== id) errors.push(`${rel(manifestPath)} id must match directory "${id}"`);
  if (!isNonEmptyString(manifest.title)) errors.push(`${rel(manifestPath)} needs a title`);
  if (!Array.isArray(manifest.whatThisProves) || manifest.whatThisProves.length === 0) {
    errors.push(`${rel(manifestPath)} needs whatThisProves entries`);
  }
  if (!Array.isArray(manifest.inputs) || manifest.inputs.length === 0) {
    errors.push(`${rel(manifestPath)} needs at least one input`);
  }

  for (const input of manifest.inputs ?? []) {
    if (!isNonEmptyString(input.path)) {
      errors.push(`${rel(manifestPath)} has an input without path`);
      continue;
    }
    if (input.format) coverage.formats.add(input.format);
    requireFile(join(ROOT, input.path), `input for ${manifest.id}`);
  }

  if (manifest.mcpTranscript) {
    checkMcpTranscript(join(exampleDir, manifest.mcpTranscript), manifest.id);
  }
  if (manifest.cliTranscript) {
    coverage.hasCli = true;
    const cliPath = join(exampleDir, manifest.cliTranscript);
    if (requireFile(cliPath, `CLI transcript for ${manifest.id}`)) {
      const text = readFileSync(cliPath, "utf8");
      if (!text.includes("office-agent") && !text.includes("packages/agent/dist/cli.js")) {
        errors.push(`${rel(cliPath)} does not invoke the office-agent CLI`);
      }
    }
  }

  if (manifest.web) {
    coverage.hasMcpWeb = coverage.hasMcpWeb || Boolean(manifest.mcpTranscript);
    if (!manifest.web.screenshotPlan) {
      errors.push(`${rel(manifestPath)} web example needs screenshotPlan`);
    } else {
      checkScreenshotPlan(join(exampleDir, manifest.web.screenshotPlan), manifest.id);
    }
  }

  for (const output of manifest.expectedOutputs ?? []) {
    if (!isNonEmptyString(output.path)) {
      errors.push(`${rel(manifestPath)} has an expected output without path`);
      continue;
    }
    if (!output.path.startsWith(GENERATED_PREFIX)) {
      errors.push(`${rel(manifestPath)} expected output must stay under ${GENERATED_PREFIX}: ${output.path}`);
    }
    if (output.format) coverage.formats.add(output.format);
  }

  const tags = new Set(manifest.tags ?? []);
  if (tags.has("diagnostics")) coverage.hasDiagnostics = true;
  if (tags.has("deliverable") || tags.has("sonaloop")) coverage.hasDeliverable = true;
}

function checkMcpTranscript(path, exampleId) {
  if (!requireFile(path, `MCP transcript for ${exampleId}`)) return;
  const transcript = readJson(path);
  if (!transcript) return;
  if (transcript.schema !== "office-ai/example-mcp-transcript@1") {
    errors.push(`${rel(path)} has unsupported schema`);
  }
  if (!Array.isArray(transcript.steps) || transcript.steps.length === 0) {
    errors.push(`${rel(path)} needs steps`);
    return;
  }
  for (const [index, step] of transcript.steps.entries()) {
    if (!KNOWN_MCP_TOOLS.has(step.tool)) {
      errors.push(`${rel(path)} step ${index + 1} uses unknown MCP tool "${step.tool}"`);
    }
    if (!isPlainRecord(step.arguments)) {
      errors.push(`${rel(path)} step ${index + 1} needs object arguments`);
    }
    if (JSON.stringify(step).includes("unsupported-operation")) coverage.hasDiagnostics = true;
  }
}

function checkScreenshotPlan(path, exampleId) {
  if (!requireFile(path, `screenshot plan for ${exampleId}`)) return;
  const plan = readJson(path);
  if (!plan) return;
  if (plan.schema !== "office-ai/example-screenshot-plan@1") {
    errors.push(`${rel(path)} has unsupported schema`);
  }
  if (!Array.isArray(plan.targets) || plan.targets.length === 0) {
    errors.push(`${rel(path)} needs at least one screenshot target`);
  }
  for (const target of plan.targets ?? []) {
    if (!isNonEmptyString(target.path) || !isNonEmptyString(target.output)) {
      errors.push(`${rel(path)} has a target without path/output`);
    }
  }
}

function isPlainRecord(value) {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function scanForSecrets(dir) {
  for (const entry of readdirSync(dir)) {
    if (entry === "_generated" || entry === "node_modules" || entry === ".DS_Store") continue;
    const path = join(dir, entry);
    const st = statSync(path);
    if (st.isDirectory()) {
      scanForSecrets(path);
      continue;
    }
    if (![".md", ".json", ".sh", ".txt"].includes(extname(entry))) continue;
    const text = readFileSync(path, "utf8");
    for (const pattern of SECRET_PATTERNS) {
      if (pattern.test(text)) errors.push(`${rel(path)} matches a sensitive path/secret pattern`);
    }
  }
}

function main() {
  requireFile(join(EXAMPLES_DIR, "README.md"), "examples README");
  const exampleDirs = readdirSync(EXAMPLES_DIR)
    .map((entry) => join(EXAMPLES_DIR, entry))
    .filter((path) => statSync(path).isDirectory())
    .filter((path) => !path.endsWith(`${resolve(EXAMPLES_DIR, "_generated")}`))
    .sort();

  if (exampleDirs.length === 0) errors.push("examples/ must contain at least one example directory");
  for (const exampleDir of exampleDirs) {
    requireFile(join(exampleDir, "README.md"), `README for ${rel(exampleDir)}`);
    checkManifest(exampleDir);
  }
  scanForSecrets(EXAMPLES_DIR);

  for (const format of ["docx", "xlsx", "pptx", "pdf"]) {
    if (!coverage.formats.has(format)) errors.push(`examples must cover ${format}`);
  }
  if (!coverage.hasMcpWeb) errors.push("at least one example must combine MCP and web review");
  if (!coverage.hasDiagnostics) errors.push("at least one example must demonstrate diagnostics");
  if (!coverage.hasDeliverable) errors.push("at least one example must demonstrate synthesis deliverables");
  if (!coverage.hasCli) errors.push("at least one example must include a CLI transcript");

  if (errors.length > 0) {
    console.error("examples-check: FAILED\n");
    for (const error of errors) console.error(`  - ${error}`);
    process.exit(1);
  }

  console.log("examples-check: OK");
  console.log(`  examples checked: ${exampleDirs.map((dir) => rel(dir)).join(", ")}`);
}

main();
