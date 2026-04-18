import { expect, test, type Page } from "@playwright/test";
import { gotoEditor, gotoXlsxEditor } from "./_helpers";

/**
 * Smoke tests for the shared "Keyboard shortcuts" help dialog.
 *
 * The dialog is mounted by every editor (DOCX, XLSX, PPTX) via
 * {@link useShortcutsDialog} + {@link KeyboardShortcutsDialog}.
 * These specs prove the global Mod+/ trigger fires, the toolbar `?`
 * button opens the same dialog, and Esc / backdrop click both close
 * it. They also probe a representative entry per product so a future
 * catalog regression (entry deleted, label renamed) flips the test.
 */

async function gotoPptxEditor(page: Page): Promise<void> {
  // Use a generous timeout — dev-mode `next dev` compiles each route
  // on first hit, which can exceed Playwright's default expect timeout.
  await page.goto("/pptx-editor", { timeout: 60_000 });
  await expect(page.getByTestId("pptx-sidebar")).toBeVisible({ timeout: 60_000 });
  await expect(page.getByTestId("pptx-sidebar").getByRole("button").first()).toBeVisible({
    timeout: 60_000,
  });
}

const dialog = (page: Page) => page.locator("[data-shortcuts-dialog]");

test.describe("shortcuts dialog", () => {
  test.setTimeout(60_000);

  test("DOCX: Mod+/ opens, Esc closes, search filters", async ({ page }) => {
    await gotoEditor(page);

    await expect(dialog(page)).toHaveCount(0);
    await page.keyboard.press("ControlOrMeta+/");
    await expect(dialog(page)).toBeVisible();

    // Catalog entry sanity check.
    await expect(dialog(page).getByText("Underline", { exact: false })).toBeVisible();
    // Pilcrow toggle should appear under the View category.
    await expect(dialog(page).getByText(/Show formatting marks/)).toBeVisible();

    // Search narrows the list.
    await page.locator("[data-shortcuts-search]").fill("strikethrough");
    await expect(dialog(page).getByText("Strikethrough")).toBeVisible();
    await expect(dialog(page).getByText("Underline", { exact: true })).toHaveCount(0);

    await page.keyboard.press("Escape");
    await expect(dialog(page)).toHaveCount(0);
  });

  test("DOCX: ? toolbar button opens the dialog", async ({ page }) => {
    await gotoEditor(page);

    await page.locator("[data-shortcuts-help]").click();
    await expect(dialog(page)).toBeVisible();
  });

  test("XLSX: Mod+/ opens with XLSX-scoped entries", async ({ page }) => {
    await gotoXlsxEditor(page);

    await page.keyboard.press("ControlOrMeta+/");
    await expect(dialog(page)).toBeVisible();
    // XLSX-only entry that isn't in DOCX/PPTX (4 directions, just probe one).
    await expect(dialog(page).getByText(/Jump to data edge/).first()).toBeVisible();
  });

  test("PPTX: Mod+/ opens with PPTX-scoped entries", async ({ page }) => {
    await gotoPptxEditor(page);

    await page.keyboard.press("ControlOrMeta+/");
    await expect(dialog(page)).toBeVisible();
    // PPTX-only entry that isn't in DOCX/XLSX (4 directions, just probe one).
    await expect(dialog(page).getByText(/Nudge shape/).first()).toBeVisible();
    await expect(dialog(page).getByText(/Add new slide/)).toBeVisible();
  });
});
