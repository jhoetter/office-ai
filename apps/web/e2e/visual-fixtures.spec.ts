import { resolve } from "node:path";
import { readdirSync } from "node:fs";
import { expect, test } from "@playwright/test";
import { gotoEditor } from "./_helpers";

/**
 * Visual regression scaffold for the docx-fidelity-overhaul (Phase 0).
 *
 * For every fixture under `fixtures/docx/real-world/` we open the file
 * via the toolbar's hidden file input and screenshot the editor surface.
 * The first run records baselines under
 * `e2e/__screenshots__/visual-fixtures.spec.ts/<fixture>-<browser>.png`;
 * later phases intentionally update the baselines as typed projections
 * (cell shading, TOC leader dots, textboxes, multi-column layout) become
 * pixel-correct.
 *
 * Note: this suite is heavier than the existing smoke specs (it opens N
 * fixtures) and is intentionally separated from the bundled-sample
 * `open-fixture.spec.ts` so a flake in one fixture does not block the
 * default smoke loop. Set `OFFICEAI_VISUAL_FIXTURES=skip` to opt out of
 * the suite entirely (e.g. on a developer machine with no baselines yet
 * who does not want to land snapshot-update churn).
 */

const FIXTURE_ROOT = resolve(__dirname, "../../../fixtures/docx/real-world");

function listFixtures(): string[] {
  try {
    return readdirSync(FIXTURE_ROOT)
      .filter((f) => f.endsWith(".docx"))
      .sort();
  } catch {
    return [];
  }
}

const SHOULD_SKIP = process.env.OFFICEAI_VISUAL_FIXTURES === "skip";

test.describe("visual regression: real-world fixtures", () => {
  test.skip(SHOULD_SKIP, "OFFICEAI_VISUAL_FIXTURES=skip");

  const fixtures = listFixtures();
  if (fixtures.length === 0) {
    test("(no real-world fixtures present — run `pnpm fixtures-real`)", async () => {
      expect(true).toBe(true);
    });
    return;
  }

  for (const name of fixtures) {
    test(`${name}: editor surface matches baseline`, async ({ page }) => {
      await gotoEditor(page);

      // The toolbar exposes the docx file input via a hidden <input
      // type="file" />. Resolve it by the accept attribute so we don't
      // collide with the image-file input that lives in the same form.
      const fileInput = page.locator('input[type="file"][accept*="wordprocessingml"]').first();
      await fileInput.setInputFiles(resolve(FIXTURE_ROOT, name));

      // Wait for the editor to re-mount with the new doc: the toolbar's
      // toast confirms the open. We also wait for the PM surface to
      // become editable again.
      await expect(page.locator(".ProseMirror")).toBeVisible({ timeout: 15_000 });
      // Allow the page-decorations plugin a tick to lay out chunks.
      await page.waitForTimeout(500);

      const surface = page.locator('[data-testid="docx-editor-surface"], .ProseMirror').first();
      await expect(surface).toHaveScreenshot(`${name}.png`, {
        // The editor pulls in fonts asynchronously; allow small render
        // diffs so a font fallback flicker doesn't fail CI.
        maxDiffPixelRatio: 0.02,
        animations: "disabled",
      });
    });
  }
});
