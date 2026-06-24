import { execFile } from "node:child_process";
import { createRequire } from "node:module";
import { createServer } from "node:net";
import { access, mkdir, rm, writeFile } from "node:fs/promises";
import { constants, existsSync } from "node:fs";
import { join, resolve } from "node:path";
import { promisify } from "node:util";
import { createLocalSessionStore, resolveOfficeAiDataDir } from "./session-store.js";

const execFileAsync = promisify(execFile);
const requireFromHere = createRequire(import.meta.url);

export type DoctorStatus = "ok" | "warning" | "error" | "optional";

export interface DoctorCheck {
  readonly code: string;
  readonly label: string;
  readonly status: DoctorStatus;
  readonly message: string;
  readonly hint?: string;
  readonly details?: Record<string, unknown>;
}

export interface DoctorReport {
  readonly schema: "office-ai/doctor@1";
  readonly ok: boolean;
  readonly summary: Record<DoctorStatus, number>;
  readonly environment: {
    readonly cwd: string;
    readonly node: string;
    readonly dataDir: string;
    readonly webPort: number;
    readonly realtimePort: number;
  };
  readonly checks: ReadonlyArray<DoctorCheck>;
}

export interface DoctorOptions {
  readonly cwd?: string;
  readonly env?: NodeJS.ProcessEnv;
  readonly nodeVersion?: string;
  readonly dataDir?: string;
  readonly commandRunner?: CommandRunner;
  readonly portChecker?: PortChecker;
}

export interface CommandResult {
  readonly ok: boolean;
  readonly stdout?: string;
  readonly stderr?: string;
  readonly error?: string;
}

export type CommandRunner = (command: string, args?: ReadonlyArray<string>) => Promise<CommandResult>;
export type PortChecker = (port: number, host: string) => Promise<PortCheckResult>;

export interface PortCheckResult {
  readonly available: boolean;
  readonly error?: string;
}

export async function runDoctor(opts: DoctorOptions = {}): Promise<DoctorReport> {
  const cwd = resolve(opts.cwd ?? process.cwd());
  const env = opts.env ?? process.env;
  const dataDir = resolveOfficeAiDataDir(opts.dataDir);
  const nodeVersion = opts.nodeVersion ?? process.version;
  const commandRunner = opts.commandRunner ?? defaultCommandRunner;
  const portChecker = opts.portChecker ?? defaultPortChecker;
  const webPort = numberFromEnv(env.PORT, 3100);
  const realtimePort = numberFromEnv(env.OAI_RT_PORT ?? env.RT_PORT, 1234);

  const checks: DoctorCheck[] = [];
  checks.push(checkNode(nodeVersion));
  checks.push(
    await checkCommand(
      "package-manager",
      "pnpm",
      ["--version"],
      commandRunner,
      "Install pnpm 9.x and run `pnpm install`."
    )
  );
  checks.push(checkBuildArtifacts(cwd));
  checks.push(await checkPort("web-port", "Web editor port", webPort, "127.0.0.1", portChecker));
  checks.push(
    await checkPort("realtime-port", "Realtime websocket port", realtimePort, "127.0.0.1", portChecker)
  );
  checks.push(
    await checkAnyCommand(
      "libreoffice",
      "LibreOffice roundtrip tooling",
      [
        ["libreoffice", ["--version"]],
        ["soffice", ["--version"]],
      ],
      commandRunner,
      "Optional for heavy OOXML roundtrip validation. Install LibreOffice to run LibreOffice gates."
    )
  );
  checks.push(await checkPlaywright(commandRunner));
  checks.push(
    checkResolvableAnyPackage(
      "pdf-core",
      ["@officeai/pdf", "pdf-lib"],
      "PDF core tooling",
      "Run `pnpm install` to restore PDF dependencies."
    )
  );
  checks.push(
    await checkCommand(
      "ocr",
      "tesseract",
      ["--version"],
      commandRunner,
      "Optional OCR path only. Install tesseract if OCR text extraction is needed.",
      "optional"
    )
  );
  checks.push(await checkDataDir(dataDir));
  checks.push(await checkSessionStore(dataDir));

  const summary: Record<DoctorStatus, number> = { ok: 0, warning: 0, error: 0, optional: 0 };
  for (const check of checks) summary[check.status] += 1;
  return {
    schema: "office-ai/doctor@1",
    ok: summary.error === 0,
    summary,
    environment: {
      cwd,
      node: nodeVersion,
      dataDir,
      webPort,
      realtimePort,
    },
    checks,
  };
}

export function renderDoctorReport(report: DoctorReport): string {
  const lines: string[] = [];
  lines.push("office-ai doctor");
  lines.push(`status: ${report.ok ? "ok" : "error"}`);
  lines.push(`cwd: ${report.environment.cwd}`);
  lines.push(`data-dir: ${report.environment.dataDir}`);
  lines.push("");
  for (const check of report.checks) {
    lines.push(`${check.status.toUpperCase().padEnd(8)} ${check.code} - ${check.message}`);
    if (check.hint) lines.push(`         hint: ${check.hint}`);
  }
  lines.push("");
  lines.push(
    `summary: ok=${report.summary.ok} warning=${report.summary.warning} optional=${report.summary.optional} error=${report.summary.error}`
  );
  return `${lines.join("\n")}\n`;
}

function checkNode(nodeVersion: string): DoctorCheck {
  const major = Number.parseInt(nodeVersion.replace(/^v/, "").split(".")[0] ?? "0", 10);
  if (major >= 20) {
    return {
      code: "node",
      label: "Node.js",
      status: "ok",
      message: `Node ${nodeVersion} satisfies >=20.`,
      details: { version: nodeVersion },
    };
  }
  return {
    code: "node",
    label: "Node.js",
    status: "error",
    message: `Node ${nodeVersion} is too old; office-ai requires Node >=20.`,
    hint: "Install Node 20 or newer, then rerun doctor.",
    details: { version: nodeVersion },
  };
}

async function checkCommand(
  code: string,
  command: string,
  args: ReadonlyArray<string>,
  runner: CommandRunner,
  hint: string,
  missingStatus: DoctorStatus = "error"
): Promise<DoctorCheck> {
  const result = await runner(command, args);
  if (result.ok) {
    const firstLine = (result.stdout ?? result.stderr ?? "").split(/\r?\n/).find(Boolean);
    return {
      code,
      label: command,
      status: "ok",
      message: firstLine ? `${command} available (${firstLine.trim()}).` : `${command} available.`,
    };
  }
  return {
    code,
    label: command,
    status: missingStatus,
    message: `${command} is not available.`,
    hint,
    details: { error: result.error ?? result.stderr },
  };
}

async function checkAnyCommand(
  code: string,
  label: string,
  candidates: ReadonlyArray<readonly [string, ReadonlyArray<string>]>,
  runner: CommandRunner,
  hint: string
): Promise<DoctorCheck> {
  const failures: string[] = [];
  for (const [command, args] of candidates) {
    const result = await runner(command, args);
    if (result.ok) {
      const firstLine = (result.stdout ?? result.stderr ?? "").split(/\r?\n/).find(Boolean);
      return {
        code,
        label,
        status: "ok",
        message: firstLine ? `${command} available (${firstLine.trim()}).` : `${command} available.`,
      };
    }
    failures.push(`${command}: ${result.error ?? "not found"}`);
  }
  return {
    code,
    label,
    status: "warning",
    message: `${label} is not available.`,
    hint,
    details: { failures },
  };
}

function checkBuildArtifacts(cwd: string): DoctorCheck {
  const required = ["packages/agent/dist/cli.js", "packages/core/dist/index.js"];
  const missing = required.filter((path) => !fileExists(join(cwd, path)));
  if (missing.length === 0) {
    return {
      code: "build-artifacts",
      label: "Build artifacts",
      status: "ok",
      message: "Core CLI build artifacts are present.",
      details: { checked: required },
    };
  }
  return {
    code: "build-artifacts",
    label: "Build artifacts",
    status: "warning",
    message: `Missing build artifacts: ${missing.join(", ")}.`,
    hint: "Run `pnpm build` before using packaged CLI/MCP entrypoints.",
    details: { missing },
  };
}

async function checkPort(
  code: string,
  label: string,
  port: number,
  host: string,
  checker: PortChecker
): Promise<DoctorCheck> {
  const result = await checker(port, host);
  if (result.available) {
    return {
      code,
      label,
      status: "ok",
      message: `${label} ${host}:${port} is available.`,
      details: { host, port },
    };
  }
  return {
    code,
    label,
    status: "warning",
    message: `${label} ${host}:${port} is already in use.`,
    hint: "Use PORT/OAI_RT_PORT overrides or stop the process holding the port before `make dev`.",
    details: { host, port, error: result.error },
  };
}

async function checkPlaywright(runner: CommandRunner): Promise<DoctorCheck> {
  const result = await runner("pnpm", ["--filter", "@officeai/web", "exec", "playwright", "--version"]);
  if (result.ok) {
    const firstLine = (result.stdout ?? "").split(/\r?\n/).find(Boolean);
    return {
      code: "playwright",
      label: "Playwright",
      status: "ok",
      message: firstLine ? `Playwright available (${firstLine.trim()}).` : "Playwright available.",
    };
  }
  return {
    code: "playwright",
    label: "Playwright",
    status: "warning",
    message: "Playwright is not available to the web package.",
    hint: "Run `pnpm --filter @officeai/web e2e:install` before screenshot/E2E gates.",
    details: { error: result.error ?? result.stderr },
  };
}

function checkResolvableAnyPackage(
  code: string,
  packages: ReadonlyArray<string>,
  label: string,
  hint: string
): DoctorCheck {
  const failures: string[] = [];
  for (const pkg of packages) {
    try {
      requireFromHere.resolve(pkg);
      return {
        code,
        label,
        status: "ok",
        message: `${label} is resolvable via ${pkg}.`,
        details: { package: pkg },
      };
    } catch (err) {
      failures.push(`${pkg}: ${err instanceof Error ? err.message : String(err)}`);
    }
  }
  return {
    code,
    label,
    status: "error",
    message: `${label} is not resolvable from @officeai/agent.`,
    hint,
    details: { packages, failures },
  };
}

async function checkDataDir(dataDir: string): Promise<DoctorCheck> {
  const probe = join(dataDir, `.doctor-${process.pid}-${Date.now()}.tmp`);
  try {
    await mkdir(dataDir, { recursive: true });
    await access(dataDir, constants.R_OK | constants.W_OK);
    await writeFile(probe, "ok\n", "utf8");
    await rm(probe, { force: true });
    return {
      code: "data-dir",
      label: "Data directory",
      status: "ok",
      message: `Data directory is readable and writable: ${dataDir}.`,
      details: { dataDir },
    };
  } catch (err) {
    return {
      code: "data-dir",
      label: "Data directory",
      status: "error",
      message: `Data directory is not writable: ${dataDir}.`,
      hint: "Set OFFICEAI_DATA_DIR to a writable directory and rerun doctor.",
      details: { dataDir, error: err instanceof Error ? err.message : String(err) },
    };
  } finally {
    await rm(probe, { force: true }).catch(() => undefined);
  }
}

async function checkSessionStore(dataDir: string): Promise<DoctorCheck> {
  try {
    const inspection = await createLocalSessionStore({ dataDir }).inspectDataDir();
    const errors = inspection.diagnostics.filter((diagnostic) => diagnostic.level === "error");
    if (errors.length > 0) {
      return {
        code: "session-store",
        label: "Session store",
        status: "error",
        message: `Session store has ${errors.length} metadata error(s).`,
        hint: "Inspect the data-dir or restore from backup before using persisted sessions.",
        details: { diagnostics: inspection.diagnostics },
      };
    }
    if (inspection.needsMigration) {
      return {
        code: "session-store",
        label: "Session store",
        status: "warning",
        message: "Session store metadata needs migration.",
        hint: "Run `office-agent sessions migrate` to back up and migrate old metadata.",
        details: { diagnostics: inspection.diagnostics },
      };
    }
    return {
      code: "session-store",
      label: "Session store",
      status: "ok",
      message: "Session store metadata is current.",
      details: { dataDir },
    };
  } catch (err) {
    return {
      code: "session-store",
      label: "Session store",
      status: "error",
      message: "Session store inspection failed.",
      hint: "Check OFFICEAI_DATA_DIR permissions and rerun doctor.",
      details: { error: err instanceof Error ? err.message : String(err) },
    };
  }
}

async function defaultCommandRunner(
  command: string,
  args: ReadonlyArray<string> = []
): Promise<CommandResult> {
  try {
    const result = await execFileAsync(command, [...args], {
      timeout: 5_000,
      maxBuffer: 1024 * 1024,
    });
    return { ok: true, stdout: result.stdout, stderr: result.stderr };
  } catch (err) {
    const e = err as NodeJS.ErrnoException & { stdout?: string; stderr?: string };
    return {
      ok: false,
      stdout: e.stdout,
      stderr: e.stderr,
      error: e.code ?? e.message,
    };
  }
}

async function defaultPortChecker(port: number, host: string): Promise<PortCheckResult> {
  return new Promise((resolvePort) => {
    const server = createServer();
    server.once("error", (err: NodeJS.ErrnoException) => {
      resolvePort({ available: false, error: err.code ?? err.message });
    });
    server.listen(port, host, () => {
      server.close(() => resolvePort({ available: true }));
    });
  });
}

function numberFromEnv(value: string | undefined, fallback: number): number {
  if (!value) return fallback;
  const parsed = Number.parseInt(value, 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallback;
}

function fileExists(path: string): boolean {
  return existsSync(path);
}
