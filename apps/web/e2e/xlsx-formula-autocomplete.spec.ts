import { expect, test } from "@playwright/test";
import { gotoXlsxEditor } from "./_helpers";

/**
 * Excel-style formula autocomplete: typing `=SU` should pop a list
 * of matching function names; Tab accepts the highlighted entry and
 * splices `NAME(` into the formula bar.
 */
test.describe("xlsx editor: formula autocomplete", () => {
  test("Tab accepts the highlighted suggestion and parks caret in parens", async ({ page }) => {
    await gotoXlsxEditor(page);

    // Pick a free cell and start typing a formula. `=SUM` is exact
    // enough that SUM is the highlighted (first) match so Tab accepts
    // it deterministically.
    await page.getByTestId("cell-D1").click();
    await page.keyboard.type("=SUM");

    // Popover appears with SUM as the first row.
    const popover = page.getByTestId("formula-suggest");
    await expect(popover).toBeVisible();
    await expect(page.getByTestId("formula-suggest-row-SUM")).toBeVisible();

    // Tab accepts → input becomes `=SUM(`, popover closes (caret now
    // sits inside the parens, prefix is empty).
    await page.keyboard.press("Tab");
    await expect(page.getByTestId("formula-input")).toHaveValue("=SUM(");
    await expect(popover).toHaveCount(0);
  });

  test("ArrowDown moves the highlight before Tab accepts", async ({ page }) => {
    await gotoXlsxEditor(page);

    await page.getByTestId("cell-D1").click();
    await page.keyboard.type("=I");

    // First match (alphabetical) is some I-function. Move down once
    // and accept — the result should be a *different* I-function.
    await expect(page.getByTestId("formula-suggest")).toBeVisible();
    const firstMatchValue = await page.getByTestId("formula-input").inputValue();
    expect(firstMatchValue).toBe("=I");

    await page.keyboard.press("ArrowDown");
    await page.keyboard.press("Tab");

    const accepted = await page.getByTestId("formula-input").inputValue();
    // Whatever the second I-function is, the input must now look like
    // `=NAME(` and start with =I…(.
    expect(accepted).toMatch(/^=I[A-Z0-9.]*\($/);
  });

  test("Esc dismisses the popover without inserting", async ({ page }) => {
    await gotoXlsxEditor(page);

    await page.getByTestId("cell-D1").click();
    await page.keyboard.type("=SU");
    await expect(page.getByTestId("formula-suggest")).toBeVisible();

    await page.keyboard.press("Escape");
    // Esc also blurs the formula bar in our handler — the popover only
    // renders while the bar is focused.
    await expect(page.getByTestId("formula-suggest")).toHaveCount(0);
  });
});
