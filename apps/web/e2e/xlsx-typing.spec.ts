import { expect, test } from "@playwright/test";
import { gotoXlsxEditor } from "./_helpers";

/**
 * "Type-to-edit" + "click-to-insert-ref" — the two interactions that
 * make the editor actually feel like Excel. Both go through the
 * existing command bus (`xlsx:set-cell-value` /
 * `xlsx:set-cell-formula`).
 */
test.describe("xlsx editor: Excel-like keyboard editing", () => {
  test("type-to-edit: a printable key on a single cell starts editing", async ({ page }) => {
    await gotoXlsxEditor(page);

    await page.getByTestId("cell-B2").click();
    await expect(page.getByTestId("cell-ref")).toHaveText("B2");

    // Press "9" then "9" then Enter — the surface-level handler should
    // pre-fill the formula bar with the first key, then subsequent
    // keys land in the now-focused input.
    await page.keyboard.press("9");
    await page.keyboard.press("9");
    await page.keyboard.press("Enter");

    await expect(page.getByTestId("cell-B2")).toContainText("99");
    // Revision badge ticked from 0 to 1 — the value went through the
    // command bus, not just the local input state.
    await expect(page.getByTestId("revision-badge")).toContainText("rev 1");
  });

  test("Backspace clears the active cell", async ({ page }) => {
    await gotoXlsxEditor(page);

    await page.getByTestId("cell-A2").click();
    await expect(page.getByTestId("formula-input")).toHaveValue("Alex");

    await page.keyboard.press("Backspace");

    // Cell renders as empty.
    await expect(page.getByTestId("cell-A2")).toHaveText("");
  });

  test("click-to-insert-ref: clicking a cell while editing a formula appends its ref", async ({ page }) => {
    await gotoXlsxEditor(page);

    // Start editing C1 with `=` — that puts us in formula point mode.
    await page.getByTestId("cell-C1").click();
    await page.keyboard.type("=");
    // Click A1 → the formula bar should become `=A1`.
    await page.getByTestId("cell-A1").click();
    await expect(page.getByTestId("formula-input")).toHaveValue("=A1");

    // Submit and assert C1 mirrors A1's value.
    await page.keyboard.press("Enter");
    await expect(page.getByTestId("cell-C1")).toContainText("Name");
  });
});
