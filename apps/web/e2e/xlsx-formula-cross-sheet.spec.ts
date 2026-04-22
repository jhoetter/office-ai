import { expect, test } from "@playwright/test";
import { gotoXlsxEditor } from "./_helpers";

test.describe("xlsx editor: cross-sheet formula reference (Excel point mode)", () => {
  test("clicks a sheet tab mid-formula and picks a cell on the other sheet", async ({ page }) => {
    await gotoXlsxEditor(page);

    // Workbook bootstraps with `Sheet1`. Add a second sheet so the
    // "click a tab while editing a formula" path has somewhere to go.
    await page.getByTestId("sheet-tab-add").click();
    await expect(page.getByTestId("sheet-tab-Sheet2")).toBeVisible();

    // Return to Sheet1 and put a value in B2 so we can assert the
    // formula resolves something concrete after submission.
    await page.getByTestId("sheet-tab-Sheet1").click();
    await page.getByTestId("cell-A1").click();

    // Begin editing a formula on Sheet1!A1.
    const formula = page.getByTestId("formula-input");
    await formula.click();
    await formula.fill("=");

    // Switch to Sheet2 — focus must stay on the formula bar so a
    // subsequent cell click inserts a sheet-qualified reference.
    await page.getByTestId("sheet-tab-Sheet2").click();
    await expect(formula).toBeFocused();

    // Pick B2 on Sheet2 → formula bar receives `Sheet2!B2`.
    await page.getByTestId("cell-B2").click();
    await expect(formula).toHaveValue("=Sheet2!B2");

    // Submit. We should snap back to Sheet1 with the new formula
    // committed against the original anchor (A1).
    await formula.press("Enter");
    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveAttribute("class", /bg-background/);
    await page.getByTestId("cell-A1").click();
    await expect(formula).toHaveValue("=Sheet2!B2");
  });

  test("Escape returns to the origin sheet without committing", async ({ page }) => {
    await gotoXlsxEditor(page);

    await page.getByTestId("sheet-tab-add").click();
    await page.getByTestId("sheet-tab-Sheet1").click();
    await page.getByTestId("cell-A2").click();
    const original = await page.getByTestId("formula-input").inputValue();

    const formula = page.getByTestId("formula-input");
    await formula.click();
    await formula.fill("=");
    await page.getByTestId("sheet-tab-Sheet2").click();
    await page.getByTestId("cell-C3").click();
    await expect(formula).toHaveValue("=Sheet2!C3");

    await formula.press("Escape");

    // Esc snaps us back to Sheet1 and leaves the original anchor cell
    // untouched (formula bar reverts to the cell's original value).
    await expect(page.getByTestId("sheet-tab-Sheet1")).toHaveAttribute("class", /bg-background/);
    await expect(page.getByTestId("cell-A2")).toBeVisible();
    await page.getByTestId("cell-A2").click();
    await expect(formula).toHaveValue(original);
  });
});
