import { spawnSync } from "node:child_process";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";

const ROOT = resolve(__dirname, "..", "..");
const SCRIPT = resolve(ROOT, "scripts/license-scan.mjs");

function runScript(args: string[] = []): { code: number; stdout: string; stderr: string } {
  const r = spawnSync("node", [SCRIPT, ...args], { encoding: "utf8", cwd: ROOT });
  return { code: r.status ?? -1, stdout: r.stdout ?? "", stderr: r.stderr ?? "" };
}

describe("scripts/license-scan.mjs", () => {
  it("exits 0 on the current lockfile (no banned licenses today)", () => {
    const r = runScript();
    expect(r.code, `stderr: ${r.stderr}\nstdout: ${r.stdout}`).toBe(0);
    expect(r.stdout).toContain("license-scan: scanned");
    expect(r.stdout).toContain("no banned licenses");
  });

  it("exits non-zero when a synthetic AGPL entry is injected (--inject-agpl)", () => {
    const r = runScript(["--inject-agpl"]);
    expect(r.code).not.toBe(0);
    expect(r.stdout).toContain("BANNED license");
    expect(r.stdout).toContain("AGPL-3.0-only");
  });
});
