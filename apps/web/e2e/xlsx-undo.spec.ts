import { expect, test } from "@playwright/test";
import { gotoXlsxEditor } from "./_helpers";

/**
 * Phase 13h — Undo / Redo. Verifies that Cmd+Z / Cmd+Shift+Z and the
 * toolbar buttons round-trip through the CommandBus.
 */
test.describe("xlsx editor: undo + redo", () => {
  test("Cmd+Z reverts a typed value, Cmd+Shift+Z reapplies it", async ({ page }) => {
    await gotoXlsxEditor(page);

    await page.getByTestId("cell-A2").click();
    await page.keyboard.type("123");
    await page.keyboard.press("Enter");
    await expect(page.getByTestId("cell-A2")).toContainText("123");

    // Click back on the surface so the keydown isn't swallowed by an
    // input.
    await page.getByTestId("cell-A2").click();
    await page.keyboard.press("ControlOrMeta+z");
    await expect(page.getByTestId("cell-A2")).not.toContainText("123");

    await page.keyboard.press("ControlOrMeta+Shift+z");
    await expect(page.getByTestId("cell-A2")).toContainText("123");
  });

  test("toolbar Undo / Redo buttons enable + disable correctly", async ({ page }) => {
    await gotoXlsxEditor(page);

    const undoBtn = page.getByTestId("action-undo");
    const redoBtn = page.getByTestId("action-redo");

    // Fresh sample workbook → nothing to undo / redo yet.
    await expect(undoBtn).toBeDisabled();
    await expect(redoBtn).toBeDisabled();

    // Make an edit so undo wakes up.
    await page.getByTestId("cell-A2").click();
    await page.keyboard.type("abc");
    await page.keyboard.press("Enter");

    await expect(undoBtn).toBeEnabled();
    await undoBtn.click();
    await expect(redoBtn).toBeEnabled();

    await redoBtn.click();
    await expect(page.getByTestId("cell-A2")).toContainText("abc");
  });

  test("a fresh edit kills the redo trail", async ({ page }) => {
    await gotoXlsxEditor(page);

    await page.getByTestId("cell-A2").click();
    await page.keyboard.type("first");
    await page.keyboard.press("Enter");

    await page.getByTestId("cell-A2").click();
    await page.keyboard.press("ControlOrMeta+z");
    await expect(page.getByTestId("action-redo")).toBeEnabled();

    // Branch the timeline → redo stack must clear.
    await page.getByTestId("cell-A3").click();
    await page.keyboard.type("branch");
    await page.keyboard.press("Enter");

    await expect(page.getByTestId("action-redo")).toBeDisabled();
  });
});
