import { expect, test } from "@playwright/test";
import { gotoXlsxEditor } from "./_helpers";

/**
 * XLSX ribbon — manual tab navigation. The ribbon root keeps the
 * `data-testid="xlsx-toolbar"` selector so existing XLSX specs
 * continue to resolve. We exercise tab switching plus the Insert tab
 * which surfaces the chart / pivot / image quick actions.
 */

test.describe("xlsx ribbon", () => {
  test("renders persistent tabs and defaults to Start", async ({ page }) => {
    await gotoXlsxEditor(page);
    const ribbon = page.getByTestId("xlsx-toolbar");
    await expect(ribbon).toBeVisible();
    await expect(ribbon).toHaveAttribute("data-active-tab", "start");

    await expect(page.getByTestId("ribbon-tab-start")).toBeVisible();
    await expect(page.getByTestId("ribbon-tab-insert")).toBeVisible();
    await expect(page.getByTestId("ribbon-tab-data")).toBeVisible();
    await expect(page.getByTestId("ribbon-tab-view")).toBeVisible();
  });

  test("Einfügen tab surfaces chart / pivot / image / comment buttons", async ({ page }) => {
    await gotoXlsxEditor(page);
    await page.getByTestId("ribbon-tab-insert").click();
    const ribbon = page.getByTestId("xlsx-toolbar");
    await expect(ribbon).toHaveAttribute("data-active-tab", "insert");

    await expect(page.getByTestId("action-insert-chart")).toBeVisible();
    await expect(page.getByTestId("action-insert-pivot")).toBeVisible();
    await expect(page.getByTestId("action-insert-image")).toBeVisible();
    await expect(page.getByTestId("action-add-comment")).toBeVisible();
  });

  test("Daten tab surfaces text-to-columns and AutoFilter toggle", async ({ page }) => {
    await gotoXlsxEditor(page);
    await page.getByTestId("ribbon-tab-data").click();
    const ribbon = page.getByTestId("xlsx-toolbar");
    await expect(ribbon).toHaveAttribute("data-active-tab", "data");

    await expect(page.getByTestId("data-text-to-columns")).toBeVisible();
    await expect(page.getByTestId("data-filter-toggle")).toBeVisible();
  });
});
