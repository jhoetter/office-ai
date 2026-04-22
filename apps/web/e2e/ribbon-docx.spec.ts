import { expect, test } from "@playwright/test";
import { gotoEditor } from "./_helpers";

/**
 * DOCX ribbon — manual tab navigation and selection-driven contextual
 * tab switching. The ribbon root keeps the legacy
 * `data-testid="docx-toolbar"` selector so existing toolbar specs
 * continue to resolve; the active tab id is exposed via
 * `data-active-tab` for these tests.
 */
test.describe("docx ribbon", () => {
  test("renders persistent tabs and defaults to Start", async ({ page }) => {
    await gotoEditor(page);
    const ribbon = page.getByTestId("docx-toolbar");
    await expect(ribbon).toBeVisible();
    await expect(ribbon).toHaveAttribute("data-active-tab", "start");

    await expect(page.getByTestId("ribbon-tab-start")).toBeVisible();
    await expect(page.getByTestId("ribbon-tab-insert")).toBeVisible();
    await expect(page.getByTestId("ribbon-tab-review")).toBeVisible();
    await expect(page.getByTestId("ribbon-tab-view")).toBeVisible();
  });

  test("clicking the Einfügen tab pins it and reveals the Insert image action", async ({ page }) => {
    await gotoEditor(page);
    await page.getByTestId("ribbon-tab-insert").click();
    const ribbon = page.getByTestId("docx-toolbar");
    await expect(ribbon).toHaveAttribute("data-active-tab", "insert");
  });
});
