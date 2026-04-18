import { readFile, readdir } from "node:fs/promises";
import { resolve } from "node:path";
import { describe, expect, it } from "vitest";
import { deterministicIdMinter } from "@officeai/core";
import { DocxAgent, docToPM } from "@officeai/docx";

/**
 * Render-snapshot suite for the real-world fixture corpus (Phase 0 of
 * docx-fidelity-overhaul). For every fixture we project the parsed
 * `DocxSnapshot` through `docToPM` and capture the resulting ProseMirror
 * document JSON.
 *
 * The captured JSON is a structural fingerprint of the typed projection
 * pipeline (parser → model → renderer's `doc-to-pm`). Phases 2-5 will
 * intentionally change these snapshots as new typed fields (cell
 * shading, TOC entries, textbox body, multi-column layout) are surfaced
 * to the editor; an unintended diff fails CI and forces the author to
 * justify the regression.
 *
 * `toDOM` output (the actual HTML the editor mounts) is exercised by the
 * unit tests in `packages/docx/src/renderer/renderer.test.ts` and the
 * Playwright visual regression suite (`apps/web/e2e/visual.spec.ts`).
 */

const FIXTURE_DIR = resolve(__dirname, "../../../fixtures/docx/real-world");

async function listFixtures(): Promise<string[]> {
  try {
    const entries = await readdir(FIXTURE_DIR);
    return entries.filter((f) => f.endsWith(".docx")).sort();
  } catch {
    return [];
  }
}

describe("DOCX renderer snapshots", async () => {
  const fixtures = await listFixtures();

  if (fixtures.length === 0) {
    it("(no real-world fixtures present — run `pnpm fixtures-real`)", () => {
      expect(true).toBe(true);
    });
    return;
  }

  for (const name of fixtures) {
    it(`${name}: doc-to-pm projection is stable`, async () => {
      const buf = await readFile(resolve(FIXTURE_DIR, name));
      // Deterministic ID minter so the snapshot is byte-stable across
      // runs (parser mints node IDs on every parse otherwise).
      const agent = await DocxAgent.fromBuffer(buf, { idMinter: deterministicIdMinter() });
      const pm = docToPM(agent.getSnapshot());
      // toJSON() is deterministic given a stable parser + projection;
      // any change must be reflected by intentionally updating the
      // snapshot file under `__snapshots__/`.
      expect(pm.toJSON()).toMatchSnapshot();
    });
  }
});
