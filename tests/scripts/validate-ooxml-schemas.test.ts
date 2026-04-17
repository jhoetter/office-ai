import { spawnSync } from "node:child_process";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";

/**
 * Hermetic unit test for `scripts/validate-ooxml-schemas.mjs`.
 *
 * The validate script's heavy path (xmllint + the full XSD bundle) is opt-in
 * and lives behind `make schema-validate` / its CI job. This test stays
 * lightweight: it invokes the script in `--self-test` mode (which deliberately
 * skips xmllint) and asserts:
 *
 *   (a) the script discovers all 6 real-world fixtures;
 *   (b) every part type observed across the fixtures maps to either a known
 *       XSD filename (`wml.xsd`, `dml-main.xsd`, …) or an explicit
 *       `skip:<reason>` bucket — i.e. there are no unrecognised parts;
 *   (c) when invoked with `--inject-broken`, the well-formedness probe
 *       correctly flags a synthetic malformed XML blob and the script
 *       propagates a non-zero exit code.
 *
 * No xmllint, no XSD bundle, no DocxAgent build is required to run this test;
 * the script is happy to run on a fresh CI runner with only Node + the docx
 * fixture directory present. The packaged dist IS imported by the underlying
 * loader though, so `pnpm build` (which `pnpm test` runs ahead of the test
 * suite via turbo) must have produced `packages/{core,docx}/dist`.
 */

const ROOT = resolve(__dirname, "..", "..");
const SCRIPT = resolve(ROOT, "scripts/validate-ooxml-schemas.mjs");

function runScript(args: string[]): { code: number; stdout: string; stderr: string } {
  const r = spawnSync("node", [SCRIPT, ...args], {
    encoding: "utf8",
    cwd: ROOT,
    timeout: 30_000,
  });
  return { code: r.status ?? -1, stdout: r.stdout ?? "", stderr: r.stderr ?? "" };
}

describe("scripts/validate-ooxml-schemas.mjs", () => {
  it("--self-test: discovers all 6 real-world fixtures and exits 0", () => {
    const r = runScript(["--self-test"]);
    expect(r.code, `stderr: ${r.stderr}\nstdout: ${r.stdout}`).toBe(0);
    expect(r.stdout).toMatch(/discovered 6 fixture\(s\)/);
    expect(r.stdout).toContain("self-test: ✅ all checks passed.");
  });

  it("--self-test: every observed part maps to a known XSD or an explicit 'skip' bucket", () => {
    const r = runScript(["--self-test"]);
    expect(r.code).toBe(0);
    // The script prints `unmapped=N`; assert N is zero. The line shape is
    // stable: `self-test: scanned <N> part(s) — <X> would-validate, <Y> skipped, <Z> unmapped`.
    const m = r.stdout.match(
      /scanned\s+(\d+)\s+part\(s\)\s+—\s+(\d+)\s+would-validate,\s+(\d+)\s+skipped,\s+(\d+)\s+unmapped/
    );
    expect(m, `expected summary line, got:\n${r.stdout}`).not.toBeNull();
    expect(Number(m![4]), "unmapped count").toBe(0);
    // Sanity: at least some parts are validated and some are skipped.
    expect(Number(m![2])).toBeGreaterThan(0);
    expect(Number(m![3])).toBeGreaterThan(0);
  });

  it("--self-test --inject-broken: failure path raises non-zero on a synthetically-broken XML blob", () => {
    const r = runScript(["--self-test", "--inject-broken"]);
    expect(r.code).not.toBe(0);
    expect(r.stdout).toContain("--inject-broken correctly flagged the synthetic blob");
  });
});
