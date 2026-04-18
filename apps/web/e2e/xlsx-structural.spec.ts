import { expect, test } from "@playwright/test";
import { gotoXlsxEditor } from "./_helpers";

/**
 * Toolbar structural ops → `xlsx:merge-cells` /
 * `xlsx:unmerge-cells` / `xlsx:insert-row` / `xlsx:delete-row`.
 *
 * These hit the same handlers the CLI / MCP tools use, so the
 * round-trip integrity guarantee carries over to the web surface.
 */
test.describe("xlsx editor: structural operations", () => {
  test("Merge collapses A1:B2 into a single rendered cell", async ({ page }) => {
    await gotoXlsxEditor(page);

    await page.getByTestId("cell-A1").click();
    await page.getByTestId("cell-B2").click({ modifiers: ["Shift"] });
    await expect(page.getByTestId("cell-ref")).toHaveText("A1:B2");

    const mergeBtn = page.getByTestId("format-merge");
    await expect(mergeBtn).toBeEnabled();
    await mergeBtn.click();

    // The non-anchor cells inside the merge are no longer rendered.
    await expect(page.getByTestId("cell-A2")).toHaveCount(0);
    await expect(page.getByTestId("cell-B1")).toHaveCount(0);
    await expect(page.getByTestId("cell-B2")).toHaveCount(0);
    // The anchor still exists and now spans the merge.
    await expect(page.getByTestId("cell-A1")).toBeVisible();

    // Clicking the merged surface (now an oversized A1) enables
    // Unmerge — the editor resolves the matching merge from the
    // single-cell selection.
    await page.getByTestId("cell-A1").click();
    const unmergeBtn = page.getByTestId("format-unmerge");
    await expect(unmergeBtn).toBeEnabled();
    await unmergeBtn.click();

    // The covered cells reappear after unmerge.
    await expect(page.getByTestId("cell-A2")).toBeVisible();
    await expect(page.getByTestId("cell-B1")).toBeVisible();
    await expect(page.getByTestId("cell-B2")).toBeVisible();
  });

  test("Insert row above pushes existing content down by one row", async ({ page }) => {
    await gotoXlsxEditor(page);

    // Header row currently has Name in A1; after Insert-Row-Above on
    // A1, Name moves to A2, leaving A1 blank.
    await page.getByTestId("cell-A1").click();
    await page.getByTestId("row-insert-above").click();

    await expect(page.getByTestId("cell-A1")).toHaveText("");
    await expect(page.getByTestId("cell-A2")).toContainText("Name");
    await expect(page.getByTestId("revision-badge")).toContainText("rev 1");
  });

  test("Delete column removes that column and shifts data left", async ({ page }) => {
    await gotoXlsxEditor(page);

    // Sample sheet headers: Name (A) / Score (B). Deleting column A
    // moves Score to column A.
    await page.getByTestId("cell-A1").click();
    await page.getByTestId("col-delete").click();

    await expect(page.getByTestId("cell-A1")).toContainText("Score");
    await expect(page.getByTestId("revision-badge")).toContainText("rev 1");
  });
});
