import { expect, test } from "@playwright/test";
import { gotoXlsxEditor } from "./_helpers";

test.describe("xlsx editor: edit a cell via the formula bar", () => {
  test("loads the sample workbook and edits A2", async ({ page }) => {
    await gotoXlsxEditor(page);

    await expect(page.getByText("XLSX editor")).toBeVisible();
    // Sanity: a couple of seeded cells should render through the
    // virtualized grid right out of the gate.
    await expect(page.getByTestId("cell-A1")).toContainText("Name");
    await expect(page.getByTestId("cell-A2")).toContainText("Alex");

    // Initial revision should be 0 (no edits yet).
    await expect(page.getByTestId("revision-badge")).toContainText("rev 0");

    // Click cell A2 → it becomes the focused cell + populates the formula bar.
    await page.getByTestId("cell-A2").click();
    await expect(page.getByTestId("cell-ref")).toHaveText("A2");
    await expect(page.getByTestId("formula-input")).toHaveValue("Alex");

    // Edit via the formula bar: clear + type Bob + Enter.
    const formula = page.getByTestId("formula-input");
    await formula.click();
    await formula.fill("Bob");
    await formula.press("Enter");

    // The grid cell should now show Bob.
    await expect(page.getByTestId("cell-A2")).toContainText("Bob");
    await expect(page.getByTestId("cell-A2")).not.toContainText("Alex");

    // Revision badge should have ticked from 0 to 1.
    await expect(page.getByTestId("revision-badge")).toContainText("rev 1");
  });
});
