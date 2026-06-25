#!/usr/bin/env node
import { spawn } from "node:child_process";
import { existsSync, mkdirSync, rmSync, statSync } from "node:fs";
import { mkdtemp } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = fileURLToPath(new URL("..", import.meta.url));
const AGENT_CLI = join(ROOT, "packages/agent/dist/cli.js");
const DOCX_FIXTURE = join(ROOT, "fixtures/docx/real-world/01-styled-letter.docx");
const WEB_PORT = Number(process.env.OAI_LOCAL_SMOKE_PORT ?? "43100");
const WEB_URL = `http://127.0.0.1:${WEB_PORT}/`;
const SCREENSHOT_PATH =
  process.env.OAI_LOCAL_SMOKE_SCREENSHOT ?? join(ROOT, "apps/web/test-results/local-install-home.png");
const REQUIRED_MCP_TOOLS = [
  "create_session",
  "import_document",
  "get_document_projection",
  "plan_command",
  "preview_command",
  "apply_command",
  "export_document",
];

let webProcess;
let mcpProcess;

try {
  assertFile(AGENT_CLI, "Missing built agent CLI. Run `pnpm --filter @officeai/agent build` first.");
  assertFile(DOCX_FIXTURE, "Missing DOCX smoke fixture.");

  const dataDir = await mkdtemp(join(tmpdir(), "office-ai-local-install-smoke-"));
  try {
    const cliResult = await smokeCliSession(dataDir);
    const mcpResult = await smokeMcpStdio(dataDir);
    const webResult = await smokeWeb(dataDir);

    console.log("office-ai local install smoke ok");
    console.log(`- CLI session flow: ${cliResult.documentId} exported ${cliResult.exportedBytes} bytes`);
    console.log(`- MCP stdio: ${mcpResult.toolCount} tools, canonical tools present`);
    console.log(`- Web editor: HTTP ${webResult.status} at ${WEB_URL}`);
    console.log(`- Screenshot: ${SCREENSHOT_PATH}`);
  } finally {
    rmSync(dataDir, { recursive: true, force: true });
  }
} catch (error) {
  stopProcess(webProcess);
  stopProcess(mcpProcess);
  console.error(error instanceof Error ? error.message : String(error));
  process.exitCode = 1;
}

async function smokeCliSession(dataDir) {
  const imported = await runJson(process.execPath, [
    AGENT_CLI,
    "sessions",
    "import",
    "--file",
    DOCX_FIXTURE,
    "--json",
    "--data-dir",
    dataDir,
  ]);
  assertEqual(imported.schema, "office-ai/import-document@1", "Unexpected import schema.");
  const documentId = imported.document?.documentId;
  assertString(documentId, "Import did not return document.documentId.");

  const projection = await runJson(process.execPath, [
    AGENT_CLI,
    "sessions",
    "projection",
    "--document-id",
    documentId,
    "--projection",
    "markdown",
    "--json",
    "--data-dir",
    dataDir,
  ]);
  assertEqual(projection.schema, "office-ai/document-projection@1", "Unexpected projection schema.");
  assertString(projection.content, "Projection did not return markdown content.");

  const outPath = join(dataDir, "exported.docx");
  const exported = await runJson(process.execPath, [
    AGENT_CLI,
    "sessions",
    "export",
    "--document-id",
    documentId,
    "--out",
    outPath,
    "--json",
    "--data-dir",
    dataDir,
  ]);
  assertEqual(exported.schema, "office-ai/export-document@1", "Unexpected export schema.");
  assertFile(outPath, "Export did not write a DOCX file.");
  return { documentId, exportedBytes: statSync(outPath).size };
}

async function smokeMcpStdio(dataDir) {
  mcpProcess = spawn(process.execPath, [AGENT_CLI, "mcp"], {
    cwd: ROOT,
    env: { ...process.env, OFFICEAI_DATA_DIR: dataDir },
    stdio: ["pipe", "pipe", "pipe"],
  });

  const stderr = [];
  const pending = new Map();
  let stdoutBuffer = "";
  let exited = false;
  mcpProcess.stderr.setEncoding("utf8");
  mcpProcess.stderr.on("data", (chunk) => stderr.push(chunk));
  mcpProcess.stdout.setEncoding("utf8");
  mcpProcess.stdout.on("data", (chunk) => {
    stdoutBuffer += chunk;
    while (stdoutBuffer.includes("\n")) {
      const index = stdoutBuffer.indexOf("\n");
      const line = stdoutBuffer.slice(0, index).trim();
      stdoutBuffer = stdoutBuffer.slice(index + 1);
      if (!line) continue;
      const message = JSON.parse(line);
      if (message.id !== undefined && pending.has(message.id)) {
        pending.get(message.id)(message);
        pending.delete(message.id);
      }
    }
  });
  mcpProcess.on("exit", () => {
    exited = true;
  });

  const initialize = await requestMcp(1, "initialize", {
    protocolVersion: "2024-11-05",
    capabilities: {},
    clientInfo: { name: "office-ai-local-install-smoke", version: "0.0.0" },
  });
  assert(!initialize.error, `MCP initialize failed: ${JSON.stringify(initialize.error)}`);
  sendMcp({ jsonrpc: "2.0", method: "notifications/initialized", params: {} });
  const listed = await requestMcp(2, "tools/list", {});
  assert(!listed.error, `MCP tools/list failed: ${JSON.stringify(listed.error)}`);
  const tools = listed.result?.tools;
  assert(Array.isArray(tools), "MCP tools/list did not return a tools array.");
  const names = new Set(tools.map((tool) => tool.name));
  for (const expected of REQUIRED_MCP_TOOLS) {
    assert(names.has(expected), `MCP tools/list is missing ${expected}.`);
  }

  stopProcess(mcpProcess);
  mcpProcess = undefined;
  return { toolCount: tools.length };

  function sendMcp(message) {
    if (!mcpProcess || exited) {
      throw new Error(`MCP process exited early.\n${stderr.join("")}`);
    }
    mcpProcess.stdin.write(`${JSON.stringify(message)}\n`);
  }

  function requestMcp(id, method, params) {
    return new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        reject(new Error(`Timed out waiting for MCP ${method}.\n${stderr.join("")}`));
      }, 10_000);
      pending.set(id, (message) => {
        clearTimeout(timeout);
        resolve(message);
      });
      sendMcp({ jsonrpc: "2.0", id, method, params });
    });
  }
}

async function smokeWeb(dataDir) {
  mkdirSync(join(ROOT, "apps/web/test-results"), { recursive: true });
  webProcess = spawn(
    "pnpm",
    ["--filter", "@officeai/web", "exec", "next", "start", "--port", String(WEB_PORT)],
    {
      cwd: ROOT,
      detached: process.platform !== "win32",
      env: { ...process.env, OFFICEAI_DATA_DIR: dataDir, PORT: String(WEB_PORT) },
      stdio: ["ignore", "pipe", "pipe"],
    }
  );

  const stderr = [];
  let exitCode;
  webProcess.stderr.setEncoding("utf8");
  webProcess.stderr.on("data", (chunk) => stderr.push(chunk));
  webProcess.on("exit", (code) => {
    exitCode = code;
  });

  const response = await waitForHttp(WEB_URL, () => exitCode, stderr);
  await run("pnpm", [
    "--filter",
    "@officeai/web",
    "exec",
    "playwright",
    "screenshot",
    "--wait-for-timeout=1000",
    WEB_URL,
    SCREENSHOT_PATH,
  ]);
  stopProcess(webProcess);
  webProcess = undefined;
  return { status: response.status };
}

async function waitForHttp(url, exitCode, stderr) {
  const deadline = Date.now() + 45_000;
  let lastError;
  while (Date.now() < deadline) {
    if (exitCode() !== undefined) {
      throw new Error(`Web server exited before HTTP smoke completed.\n${stderr.join("")}`);
    }
    try {
      const response = await fetch(url, { signal: AbortSignal.timeout(2_000) });
      if (response.ok) return response;
      lastError = new Error(`HTTP ${response.status}`);
    } catch (error) {
      lastError = error;
    }
    await sleep(500);
  }
  throw new Error(`Timed out waiting for ${url}: ${lastError?.message ?? lastError}`);
}

async function runJson(command, args) {
  const result = await run(command, args);
  try {
    return JSON.parse(result.stdout);
  } catch (error) {
    throw new Error(
      `Command did not emit JSON: ${command} ${args.join(" ")}\n${result.stdout}\n${result.stderr}`
    );
  }
}

function run(command, args) {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { cwd: ROOT, env: process.env, stdio: ["ignore", "pipe", "pipe"] });
    let stdout = "";
    let stderr = "";
    child.stdout.setEncoding("utf8");
    child.stderr.setEncoding("utf8");
    child.stdout.on("data", (chunk) => {
      stdout += chunk;
    });
    child.stderr.on("data", (chunk) => {
      stderr += chunk;
    });
    child.on("error", reject);
    child.on("exit", (code) => {
      if (code === 0) {
        resolve({ stdout, stderr });
      } else {
        reject(new Error(`Command failed (${code}): ${command} ${args.join(" ")}\n${stdout}\n${stderr}`));
      }
    });
  });
}

function stopProcess(child) {
  if (!child || child.killed) return;
  if (process.platform !== "win32" && child.pid) {
    try {
      process.kill(-child.pid, "SIGTERM");
      return;
    } catch {
      // Fall through to killing the direct child.
    }
  }
  child.kill("SIGTERM");
}

function assertFile(path, message) {
  assert(existsSync(path), message);
}

function assertEqual(actual, expected, message) {
  assert(actual === expected, `${message} Expected ${expected}, got ${actual}.`);
}

function assertString(value, message) {
  assert(typeof value === "string" && value.length > 0, message);
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
