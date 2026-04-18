import { resolve } from "node:path";
import { test, expect } from "@playwright/test";
import { gotoEditor } from "./_helpers";

/**
 * Debug spec for the masterthesis fixture (Phase B verification of the
 * docx-fidelity-overhaul). Loads the doc, scrolls through the first
 * ~12 pages, and snapshots the editor surface so we can eyeball the
 * TOC layout, page-break behaviour, and quadrant-table rendering after
 * the wrapper-marker / `lastRenderedPageBreak` fixes. Skipped on CI;
 * run locally with:
 *
 *   pnpm --filter @officeai/web exec playwright test masterthesis-debug
 */

const FIXTURE = resolve(
  __dirname,
  "../../../Masterthesis_Rohfassung_Verification_Asymmetry Kopie.docx"
);

test.describe("masterthesis fixture", () => {
  test("opens, paginates, and screenshots the first pages", async ({ page }) => {
    test.setTimeout(120_000);
    await gotoEditor(page);
    const fileInput = page
      .locator('input[type="file"][accept*="wordprocessingml"]')
      .first();
    await fileInput.setInputFiles(FIXTURE);

    await expect(page.locator(".ProseMirror")).toBeVisible({ timeout: 30_000 });
    // Allow the page-decorations plugin two RAFs to measure block heights
    // and lay out chunks once the masterthesis (414 paragraphs) projects.
    await page.waitForTimeout(4_000);

    const pageCount = await page
      .locator(".pm-page-block")
      .count()
      .catch(() => 0);
    test.info().annotations.push({ type: "page-count", description: String(pageCount) });
    // eslint-disable-next-line no-console
    console.log(`[masterthesis-debug] pm-page-block count=${pageCount}`);

    // Page-block-level shots are unreliable when each block paints into a
    // CSS grid track that the viewport rounds to 0px during scroll;
    // capture the entire scrolling viewport instead so we can see the
    // first 8 sheets in 8 separate snapshots.
    const sheetScroller = page.locator('[data-testid="docx-editor-surface"]').first();
    const exists = await sheetScroller.count();
    const scroller = exists > 0 ? sheetScroller : page.locator(".ProseMirror").first();

    // Reset to the top.
    await scroller.evaluate((el) => {
      (el as HTMLElement).scrollTop = 0;
    });
    await page.waitForTimeout(300);

    for (let i = 0; i < 12; i++) {
      await page.screenshot({
        path: `test-results/masterthesis-vp-${String(i + 1).padStart(2, "0")}.png`,
        fullPage: false,
      });
      await scroller.evaluate((el, step) => {
        (el as HTMLElement).scrollBy({ top: step, behavior: "instant" as ScrollBehavior });
      }, 760);
      await page.waitForTimeout(250);
    }
  });
});
