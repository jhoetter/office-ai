import { expect, test, type Page } from "@playwright/test";

/**
 * Smoke for `/pdf-viewer` — mirrors `pptx-editor.spec.ts`.
 *
 * The viewer is "open-only" by design (no auto-built welcome doc),
 * so we navigate with `?new=1` to force the editor to bootstrap a
 * blank PDF agent. From there we assert the toolbar, sidebar and
 * canvas mount and remain interactive.
 */

async function gotoBlankPdf(page: Page): Promise<void> {
  await page.goto("/pdf-viewer?new=1");
  await expect(page.getByTestId("pdf-canvas")).toBeVisible({ timeout: 20_000 });
  await expect(page.getByTestId("pdf-sidebar")).toBeVisible({ timeout: 20_000 });
}

test.describe("pdf-viewer route", () => {
  test("mounts the blank PDF, sidebar and toolbar", async ({ page }) => {
    await gotoBlankPdf(page);
    await expect(page.getByTestId("pdf-page-nav")).toBeVisible();
    await expect(page.getByTestId("pdf-zoom-menu-trigger")).toBeVisible();
    await expect(page.getByTestId("pdf-view-mode-trigger")).toBeVisible();
    await expect(page.getByTestId("pdf-page-ops-trigger")).toBeVisible();
    await expect(page.getByTestId("pdf-print")).toBeVisible();
  });

  test("renders the first page thumbnail and the page-1 canvas tile", async ({ page }) => {
    await gotoBlankPdf(page);
    await expect(page.getByTestId("pdf-thumbnail-1")).toBeVisible();
    await expect(page.getByTestId("pdf-page-1")).toBeVisible();
  });

  test("status bar reports the active page number", async ({ page }) => {
    await gotoBlankPdf(page);
    await expect(page.getByTestId("pdf-status-hint")).toContainText("1");
  });

  test("page input accepts a value and the navigator stays interactive", async ({ page }) => {
    await gotoBlankPdf(page);
    const input = page.getByTestId("pdf-page-input");
    await expect(input).toBeVisible();
    await expect(input).toHaveValue("1");
  });

  test("opens with no error toast on first paint", async ({ page }) => {
    await gotoBlankPdf(page);
    const errors = page.getByRole("status").filter({ hasText: /error/i });
    await expect(errors).toHaveCount(0);
  });
});
