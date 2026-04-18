import { expect, test } from "@playwright/test";
import { gotoXlsxEditor } from "./_helpers";

/**
 * AutoFilter UX: toggle the filter band from the toolbar, open the
 * column dropdown via the chevron, and apply a values filter to hide
 * a row. Mirrors the persistence story tested at the model level
 * (see tests/roundtrip/xlsx/commands-roundtrip.test.ts).
 */
test.describe("xlsx editor: data filters", () => {
  test("apply autoFilter, open dropdown, hide a row by value", async ({ page }) => {
    await gotoXlsxEditor(page);

    // Sanity: seeded body rows render.
    await expect(page.getByTestId("cell-A2")).toContainText("Alex");
    await expect(page.getByTestId("cell-A3")).toContainText("Sam");

    // Toggle the AutoFilter on. Detection picks the workbook's used range.
    await page.getByTestId("data-filter-toggle").click();

    // The column-A header now sprouts a filter chevron.
    const chevronA = page.getByTestId("col-filter-A");
    await expect(chevronA).toBeVisible();
    await expect(page.getByTestId("col-filter-B")).toBeVisible();

    // Open the dropdown for column A.
    await chevronA.click();
    const dropdown = page.getByTestId("filter-dropdown");
    await expect(dropdown).toBeVisible();

    // Toggle off "Alex" only — leaves Sam, Total visible (Total isn't part
    // of the body but IS in the autoFilter range so it shows in the list).
    await page.getByTestId("filter-select-all").click(); // unchecks all
    await page.getByTestId("filter-value-Sam").click();
    await page.getByTestId("filter-value-Total").click();
    await page.getByTestId("filter-apply").click();

    // Dropdown closes after apply.
    await expect(dropdown).toHaveCount(0);

    // The "Alex" row is now hidden — row 2 should not render.
    await expect(page.getByTestId("cell-A2")).toHaveCount(0);
    // Sam still visible.
    await expect(page.getByTestId("cell-A3")).toContainText("Sam");

    // Revision bumped through the bus (set-auto-filter + set-filter-column).
    await expect(page.getByTestId("revision-badge")).not.toContainText("rev 0");
  });

  test("toggling the toolbar button off removes the chevrons", async ({ page }) => {
    await gotoXlsxEditor(page);

    await page.getByTestId("data-filter-toggle").click();
    await expect(page.getByTestId("col-filter-A")).toBeVisible();

    await page.getByTestId("data-filter-toggle").click();
    await expect(page.getByTestId("col-filter-A")).toHaveCount(0);
  });
});
