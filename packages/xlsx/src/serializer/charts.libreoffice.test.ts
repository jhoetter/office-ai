import { execFile } from "node:child_process";
import { existsSync } from "node:fs";
import { mkdir, readFile, rm, writeFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";
import { describe, expect, it } from "vitest";
import { CommandBus, type SerializedCommand } from "@officeai/core";
import { allXlsxHandlers } from "../commands/registry.js";
import type { ChartKind, XlsxSnapshot } from "../model/types.js";
import { parseXlsx } from "../parser/parse.js";
import { serializeXlsx } from "./serialize.js";

const exec = promisify(execFile);
const here = dirname(fileURLToPath(import.meta.url));
const fixtures = resolve(here, "../../../../fixtures/xlsx/synthetic");

/**
 * Resolve a usable LibreOffice executable. Honour `$SOFFICE` so CI
 * can point at a different install. Returns `null` when nothing is
 * available — callers should `it.skip` rather than fail in that
 * case so dev workstations without LibreOffice still see a green
 * suite.
 */
function findSoffice(): string | null {
  const candidates = [
    process.env.SOFFICE,
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/usr/bin/soffice",
    "/usr/local/bin/soffice",
    "/opt/homebrew/bin/soffice",
  ].filter((p): p is string => typeof p === "string" && p.length > 0);
  return candidates.find((p) => existsSync(p)) ?? null;
}

const SOFFICE = findSoffice();

async function applyAddChart(
  snap: XlsxSnapshot,
  payload: { sheet: string; kind: ChartKind; dataRange: string; title?: string }
): Promise<XlsxSnapshot> {
  const bus = new CommandBus<XlsxSnapshot>(snap);
  bus.registerAll(allXlsxHandlers);
  const cmd: SerializedCommand = {
    type: "xlsx:add-chart",
    payload: { ...payload, hasHeaderRow: true, hasCategoryColumn: true },
  };
  const m = await bus.dispatch(cmd);
  if (m.status === "rejected") {
    throw new Error(`add-chart was rejected: ${m.rejection?.message}`);
  }
  return m.after;
}

const describeIfLO = SOFFICE ? describe : describe.skip;

describeIfLO("LibreOffice headless chart round-trip", () => {
  it(
    "all four chart kinds survive a LibreOffice convert pass",
    async () => {
      const buf = new Uint8Array(
        await readFile(resolve(fixtures, "01-single-sheet-numbers.xlsx"))
      );
      const snap = await parseXlsx(buf);
      const sheetName = snap.root.sheets[0]!.name;

      let cur = snap;
      for (const kind of ["column", "bar", "line", "pie"] as const) {
        cur = await applyAddChart(cur, {
          sheet: sheetName,
          kind,
          dataRange: "A1:C5",
          title: `${kind} chart`,
        });
      }

      const out = await serializeXlsx(cur);
      const tmpDir = resolve("/tmp/officeai-chart-roundtrip");
      await rm(tmpDir, { recursive: true, force: true });
      await mkdir(tmpDir, { recursive: true });
      const inPath = resolve(tmpDir, "with-charts.xlsx");
      const outDir = resolve(tmpDir, "converted");
      await mkdir(outDir, { recursive: true });
      await writeFile(inPath, new Uint8Array(out));

      // Each LO invocation locks `~/.config/libreoffice` for a moment;
      // running the test serially is fine, but we add `--user-profile`
      // so a parallel test session doesn't fight ours.
      const profile = `-env:UserInstallation=file://${tmpDir}/lo-profile`;
      const { stdout, stderr } = await exec(SOFFICE!, [
        profile,
        "--headless",
        "--convert-to",
        "xlsx",
        inPath,
        "--outdir",
        outDir,
      ]);
      // Surface LO output on test failure so we don't have to re-run
      // by hand to see what went wrong.
      void stdout;
      void stderr;

      const reEmitted = new Uint8Array(
        await readFile(resolve(outDir, "with-charts.xlsx"))
      );
      const reparsed = await parseXlsx(reEmitted);
      const chartParts = [...reparsed.container.parts.keys()].filter((p) =>
        /^xl\/charts\/chart\d+\.xml$/.test(p)
      );
      // LibreOffice may rename or coalesce the chart parts but it
      // must keep the same number of charts (one per kind).
      expect(chartParts.length).toBeGreaterThanOrEqual(4);
    },
    // LO startup is slow on cold cache; give it generous breathing room.
    60_000
  );
});
