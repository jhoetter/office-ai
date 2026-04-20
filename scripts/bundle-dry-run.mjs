#!/usr/bin/env node
/**
 * Smoke-test the GitHub Releases bundle without uploading anything.
 *
 * Mirrors the `Build release bundle` step in
 * .github/workflows/auto-release.yml: runs `pnpm --filter @officeai/agent
 * --prod deploy <tmp>` to materialise every `workspace:*` dep into a
 * self-contained directory, sanity-checks the resulting bundle, prints
 * its size, and cleans up.
 *
 * Used by `pnpm verify` so a broken `pnpm deploy` (e.g. a workspace dep
 * cycle, missing build output, broken CLI shebang) is caught locally
 * before it reaches CI.
 *
 * Pure dry-run: nothing is uploaded, nothing is left on disk.
 */
import { execFileSync } from "node:child_process";
import { mkdtempSync, rmSync, statSync, readdirSync, readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";

const PKG = "@officeai/agent";

function run(cmd, args, opts = {}) {
  return execFileSync(cmd, args, { stdio: "inherit", ...opts });
}

function bytes(n) {
  const units = ["B", "KB", "MB", "GB"];
  let i = 0;
  let v = n;
  while (v >= 1024 && i < units.length - 1) {
    v /= 1024;
    i += 1;
  }
  return `${v.toFixed(1)} ${units[i]}`;
}

function dirSize(dir) {
  let total = 0;
  for (const entry of readdirSync(dir, { withFileTypes: true })) {
    const full = join(dir, entry.name);
    if (entry.isDirectory()) total += dirSize(full);
    else if (entry.isFile()) total += statSync(full).size;
  }
  return total;
}

const out = mkdtempSync(join(tmpdir(), "officeai-bundle-dryrun-"));
let exitCode = 0;
try {
  console.log(`bundle:dry-run -> ${PKG} into ${out}`);
  run("pnpm", ["--filter", PKG, "--prod", "deploy", out]);

  const cliPath = join(out, "dist", "cli.js");
  let stat;
  try {
    stat = statSync(cliPath);
  } catch {
    throw new Error(`bundle missing ${cliPath} — did @officeai/agent build run first?`);
  }
  if (!stat.isFile()) {
    throw new Error(`bundle ${cliPath} is not a regular file`);
  }
  const firstLine = readFileSync(cliPath, "utf8").split("\n", 1)[0];
  if (!/^#!.*\bnode\b/.test(firstLine)) {
    throw new Error(`bundle cli.js missing node shebang (got: ${JSON.stringify(firstLine)})`);
  }

  const nm = join(out, "node_modules");
  let nmStat;
  try {
    nmStat = statSync(nm);
  } catch {
    throw new Error("bundle missing node_modules/ — workspace deps did not get inlined");
  }
  if (!nmStat.isDirectory()) {
    throw new Error("bundle node_modules is not a directory");
  }

  const total = dirSize(out);
  console.log(`bundle:dry-run OK — ${PKG}, total ${bytes(total)} (cli.js + node_modules verified)`);
} catch (err) {
  console.error(`bundle:dry-run FAILED: ${err.message}`);
  exitCode = 1;
} finally {
  rmSync(out, { recursive: true, force: true });
}
process.exit(exitCode);
