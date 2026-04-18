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

  test("type-to-edit: typed keystrokes appear inside the cell, not just the formula bar", async ({
    page,
  }) => {
    // Excel parity — when the user types into a selected cell, the
    // characters must visibly fill the cell as they're typed. The
    // <input> still lives in the formula bar (so click-to-insert-ref
    // and autocomplete keep working), but the cell mirrors the draft
    // through the `liveEditDraft` channel.
    await gotoXlsxEditor(page);

    await page.getByTestId("cell-B2").click();
    await page.keyboard.press("4");
    await page.keyboard.press("2");

    const liveDraft = page.getByTestId("cell-B2-live-draft");
    await expect(liveDraft).toBeVisible();
    await expect(liveDraft).toHaveText("42");
    // Formula bar still sees the same draft (single source of truth).
    await expect(page.getByTestId("formula-input")).toHaveValue("42");

    // Pressing Escape cancels the edit; the cell snaps back to its
    // original value and the live-draft node is gone.
    await page.keyboard.press("Escape");
    await expect(page.getByTestId("cell-B2-live-draft")).toHaveCount(0);
    await expect(page.getByTestId("cell-B2")).toContainText("42"); // pre-existing seed
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
