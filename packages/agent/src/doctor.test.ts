import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, describe, expect, it } from "vitest";
import { renderDoctorReport, runDoctor, type CommandRunner, type PortChecker } from "./doctor.js";

const tempDirs: string[] = [];

afterEach(() => {
  for (const dir of tempDirs.splice(0)) {
    rmSync(dir, { recursive: true, force: true });
  }
});

function tempDir(): string {
  const dir = mkdtempSync(join(tmpdir(), "officeai-doctor-"));
  tempDirs.push(dir);
  return dir;
}

describe("doctor", () => {
  it("reports hard errors for missing core runtime prerequisites", async () => {
    const runner: CommandRunner = async (command) => {
      if (command === "pnpm") return { ok: false, error: "ENOENT" };
      if (command === "libreoffice" || command === "soffice") return { ok: false, error: "ENOENT" };
      if (command === "tesseract") return { ok: false, error: "ENOENT" };
      return { ok: false, error: "unexpected" };
    };
    const ports: PortChecker = async () => ({ available: true });

    const report = await runDoctor({
      cwd: tempDir(),
      dataDir: join(tempDir(), "data"),
      nodeVersion: "v18.19.0",
      commandRunner: runner,
      portChecker: ports,
    });

    expect(report.ok).toBe(false);
    expect(report.checks.find((check) => check.code === "node")).toMatchObject({ status: "error" });
    expect(report.checks.find((check) => check.code === "package-manager")).toMatchObject({
      status: "error",
    });
    expect(report.checks.find((check) => check.code === "libreoffice")).toMatchObject({
      status: "warning",
    });
    expect(report.checks.find((check) => check.code === "ocr")).toMatchObject({ status: "optional" });
  });

  it("keeps optional and heavy tooling non-blocking", async () => {
    const runner: CommandRunner = async (command, args = []) => {
      if (command === "pnpm" && args.includes("playwright")) {
        return { ok: false, error: "playwright-missing" };
      }
      if (command === "pnpm") return { ok: true, stdout: "9.15.4\n" };
      if (command === "libreoffice") return { ok: true, stdout: "LibreOffice 24.8\n" };
      if (command === "tesseract") return { ok: false, error: "ENOENT" };
      return { ok: false, error: "unexpected" };
    };
    const ports: PortChecker = async (port) => ({ available: port !== 3100, error: "EADDRINUSE" });

    const report = await runDoctor({
      cwd: process.cwd(),
      dataDir: join(tempDir(), "data"),
      nodeVersion: "v20.11.1",
      commandRunner: runner,
      portChecker: ports,
    });

    expect(report.ok).toBe(true);
    expect(report.checks.find((check) => check.code === "web-port")).toMatchObject({
      status: "warning",
    });
    expect(report.checks.find((check) => check.code === "playwright")).toMatchObject({
      status: "warning",
    });
    expect(report.checks.find((check) => check.code === "ocr")).toMatchObject({ status: "optional" });
  });

  it("warns when persisted sessions need migration", async () => {
    const dataDir = tempDir();
    const sessionDir = join(dataDir, "sessions", "session_legacy");
    mkdirSync(sessionDir, { recursive: true });
    writeFileSync(
      join(sessionDir, "session.json"),
      JSON.stringify({
        id: "session_legacy",
        title: "Legacy",
        createdAt: "2026-06-24T09:00:00.000Z",
        updatedAt: "2026-06-24T09:00:00.000Z",
        documentIds: [],
      })
    );

    const report = await runDoctor({
      cwd: process.cwd(),
      dataDir,
      nodeVersion: "v20.11.1",
      commandRunner: async () => ({ ok: true, stdout: "ok\n" }),
      portChecker: async () => ({ available: true }),
    });

    expect(report.ok).toBe(true);
    expect(report.checks.find((check) => check.code === "session-store")).toMatchObject({
      status: "warning",
    });
  });

  it("renders a compact human report", async () => {
    const report = await runDoctor({
      cwd: process.cwd(),
      dataDir: join(tempDir(), "data"),
      nodeVersion: "v20.11.1",
      commandRunner: async () => ({ ok: true, stdout: "ok\n" }),
      portChecker: async () => ({ available: true }),
    });

    const text = renderDoctorReport(report);
    expect(text).toContain("office-ai doctor");
    expect(text).toContain("summary:");
    expect(text).toContain("data-dir:");
  });
});
