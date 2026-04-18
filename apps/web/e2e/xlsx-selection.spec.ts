import { expect, test } from "@playwright/test";
import { gotoXlsxEditor } from "./_helpers";

/**
 * Multi-cell selection: shift-click should extend the active range.
 * The cell-ref pill is the canonical UI for "what's selected": it
 * shows a single A1 for one cell and `A1:B3` for a range.
 */
test.describe("xlsx editor: multi-cell selection", () => {
  test("shift-click extends the selection from anchor to focus", async ({ page }) => {
    await gotoXlsxEditor(page);

    // Click A1 → single-cell selection.
    await page.getByTestId("cell-A1").click();
    await expect(page.getByTestId("cell-ref")).toHaveText("A1");

    // Shift-click B3 → bounding range covers A1:B3.
    await page.getByTestId("cell-B3").click({ modifiers: ["Shift"] });
    await expect(page.getByTestId("cell-ref")).toHaveText("A1:B3");

    // The bounding-box marquee renders for ranges.
    await expect(page.getByTestId("grid-marquee")).toBeVisible();
  });

  test("plain click collapses back to a single cell", async ({ page }) => {
    await gotoXlsxEditor(page);

    await page.getByTestId("cell-A1").click();
    await page.getByTestId("cell-B3").click({ modifiers: ["Shift"] });
    await expect(page.getByTestId("cell-ref")).toHaveText("A1:B3");

    await page.getByTestId("cell-B2").click();
    await expect(page.getByTestId("cell-ref")).toHaveText("B2");
  });
});
