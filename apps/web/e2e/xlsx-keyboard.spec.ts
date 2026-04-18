import { expect, test } from "@playwright/test";
import { gotoXlsxEditor } from "./_helpers";

test.describe("xlsx editor: keyboard navigation (P12d-g)", () => {
  test("arrow keys move the single-cell selection", async ({ page }) => {
    await gotoXlsxEditor(page);
    await page.getByTestId("cell-A1").click();

    // Pull focus back to the surface (click inside the grid is fine).
    await page.keyboard.press("ArrowDown");
    await expect(page.getByTestId("cell-ref")).toHaveText("A2");
    await page.keyboard.press("ArrowRight");
    await expect(page.getByTestId("cell-ref")).toHaveText("B2");
    await page.keyboard.press("ArrowUp");
    await expect(page.getByTestId("cell-ref")).toHaveText("B1");
    await page.keyboard.press("ArrowLeft");
    await expect(page.getByTestId("cell-ref")).toHaveText("A1");
  });

  test("Shift+Arrow extends the selection from the existing anchor", async ({ page }) => {
    await gotoXlsxEditor(page);
    await page.getByTestId("cell-A1").click();

    await page.keyboard.press("Shift+ArrowDown");
    await page.keyboard.press("Shift+ArrowRight");
    await expect(page.getByTestId("cell-ref")).toHaveText("A1:B2");
  });

  test("Tab and Enter commit the formula bar and move selection", async ({ page }) => {
    await gotoXlsxEditor(page);

    // Type into A1 → Tab → A1 commits, selection moves to B1.
    await page.getByTestId("cell-A1").click();
    await page.keyboard.press("Backspace"); // clear A1
    await page.keyboard.type("hello");
    await page.keyboard.press("Tab");
    await expect(page.getByTestId("cell-A1")).toContainText("hello");
    await expect(page.getByTestId("cell-ref")).toHaveText("B1");

    // Now type in B1 → Enter → commits and moves selection down to B2.
    await page.keyboard.type("world");
    await page.keyboard.press("Enter");
    await expect(page.getByTestId("cell-B1")).toContainText("world");
    await expect(page.getByTestId("cell-ref")).toHaveText("B2");
  });

  test("Home goes to column A; Ctrl+Home jumps to A1", async ({ page }) => {
    await gotoXlsxEditor(page);

    await page.getByTestId("cell-C3").click();
    await page.keyboard.press("Home");
    await expect(page.getByTestId("cell-ref")).toHaveText("A3");

    await page.getByTestId("cell-C3").click();
    await page.keyboard.press("ControlOrMeta+Home");
    await expect(page.getByTestId("cell-ref")).toHaveText("A1");
  });

  test("F2 enters edit mode focusing the formula bar", async ({ page }) => {
    await gotoXlsxEditor(page);

    await page.getByTestId("cell-A1").click();
    await page.keyboard.press("F2");

    const formula = page.getByTestId("formula-input");
    await expect(formula).toBeFocused();
  });

  test("clicking a row header selects the entire row", async ({ page }) => {
    await gotoXlsxEditor(page);

    await page.getByTestId("row-header-2").click();
    // Whole-row selection — expect the cell-ref reflects A2:Z2.
    await expect(page.getByTestId("cell-ref")).toHaveText("A2:Z2");
  });

  test("clicking a column header selects the entire column", async ({ page }) => {
    await gotoXlsxEditor(page);

    await page.getByTestId("col-header-B").click();
    // Whole-column selection — total rows is 1000 → B1:B1000.
    await expect(page.getByTestId("cell-ref")).toHaveText("B1:B1000");
  });

  test("Delete on a whole-row selection deletes the row", async ({ page }) => {
    await gotoXlsxEditor(page);

    // A2 currently shows "Alex". Deleting row 2 should bump A3
    // (currently "Sam") up into A2.
    await expect(page.getByTestId("cell-A2")).toContainText("Alex");
    await page.getByTestId("row-header-2").click();
    await expect(page.getByTestId("xlsx-surface")).toHaveAttribute("data-whole-row", "1");
    await page.keyboard.press("Delete");

    await expect(page.getByTestId("cell-A2")).toContainText("Sam");
    await expect(page.getByTestId("revision-badge")).toContainText("rev 1");
  });

  test("Delete on a whole-column selection deletes the column", async ({ page }) => {
    await gotoXlsxEditor(page);

    // Column B currently has the "Score" header. Deleting it should
    // shift the remaining columns left.
    await expect(page.getByTestId("cell-B1")).toContainText("Score");
    await page.getByTestId("col-header-B").click();
    await page.keyboard.press("Delete");

    await expect(page.getByTestId("cell-B1")).not.toContainText("Score");
    await expect(page.getByTestId("revision-badge")).toContainText("rev 1");
  });

  test("Delete on a single cell only clears its content", async ({ page }) => {
    await gotoXlsxEditor(page);

    await page.getByTestId("cell-A2").click();
    await page.keyboard.press("Delete");

    await expect(page.getByTestId("cell-A2")).not.toContainText("Alex");
    // A3 should still hold its old value — single-cell delete must
    // not shift other cells around.
    await expect(page.getByTestId("cell-A3")).toContainText("Sam");
  });
});
