import { expect, test } from "@playwright/test";
import { gotoXlsxEditor } from "./_helpers";

/**
 * Excel keyboard-shortcut smoke tests. The new bindings live in
 * `XlsxEditor.tsx`'s `onSurfaceKeyDown` and dispatch through the
 * existing `xlsx:set-cell-format` plumbing, so we assert via the
 * rendered cell CSS — same signal the toolbar specs use.
 */
test.describe("xlsx editor: keyboard shortcuts", () => {
  test.setTimeout(60_000);

  test("Mod+B toggles bold on the active cell", async ({ page }) => {
    await gotoXlsxEditor(page);

    const a1 = page.getByTestId("cell-A1");
    await a1.click();
    await expect(a1).toHaveCSS("font-weight", "400");

    await page.keyboard.press("ControlOrMeta+b");
    await expect(a1).toHaveCSS("font-weight", "700");

    await page.keyboard.press("ControlOrMeta+b");
    await expect(a1).toHaveCSS("font-weight", "400");
  });

  test("Mod+I + Mod+U stack on the active cell", async ({ page }) => {
    await gotoXlsxEditor(page);

    const b2 = page.getByTestId("cell-B2");
    await b2.click();
    await page.keyboard.press("ControlOrMeta+i");
    await page.keyboard.press("ControlOrMeta+u");

    await expect(b2).toHaveCSS("font-style", "italic");
    await expect(b2).toHaveCSS("text-decoration-line", "underline");
  });

  test("Mod+Shift+5 applies percent format to the active cell", async ({ page }) => {
    await gotoXlsxEditor(page);

    // Type a numeric literal first so the percent format has
    // something to render on.
    const a1 = page.getByTestId("cell-A1");
    await a1.click();
    await page.keyboard.type("0.42");
    await page.keyboard.press("Enter");

    await a1.click();
    await page.keyboard.press("ControlOrMeta+Shift+5");

    // 0.42 → "42%" once the percent numFmt is applied. We allow the
    // suffix to vary slightly across locales by matching just the
    // "%" in the rendered text.
    await expect(a1).toContainText("%");
  });
});
