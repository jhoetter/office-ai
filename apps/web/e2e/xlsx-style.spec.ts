import { expect, test } from "@playwright/test";
import { gotoXlsxEditor } from "./_helpers";

/**
 * Toolbar formatting → `xlsx:set-cell-format`. Verifies that
 * Bold/Italic/Underline ripple through the command bus into the
 * rendered cell, and that the toolbar's pressed state mirrors the
 * active cell's effective style.
 */
test.describe("xlsx editor: formatting toolbar", () => {
  test("Bold toggles font weight on the selected cell", async ({ page }) => {
    await gotoXlsxEditor(page);

    await page.getByTestId("cell-A1").click();
    await expect(page.getByTestId("cell-ref")).toHaveText("A1");

    const a1 = page.getByTestId("cell-A1");
    await expect(a1).toHaveCSS("font-weight", "400");

    const boldBtn = page.getByTestId("format-bold");
    await expect(boldBtn).toHaveAttribute("aria-pressed", "false");
    await boldBtn.click();

    await expect(boldBtn).toHaveAttribute("aria-pressed", "true");
    await expect(a1).toHaveCSS("font-weight", "700");
    // Roundtrips through the command bus → revision ticked.
    await expect(page.getByTestId("revision-badge")).toContainText("rev 1");

    // Toggling off should restore.
    await boldBtn.click();
    await expect(boldBtn).toHaveAttribute("aria-pressed", "false");
    await expect(a1).toHaveCSS("font-weight", "400");
  });

  test("Italic + Underline stack on the selected cell", async ({ page }) => {
    await gotoXlsxEditor(page);

    await page.getByTestId("cell-B2").click();
    await page.getByTestId("format-italic").click();
    await page.getByTestId("format-underline").click();

    const b2 = page.getByTestId("cell-B2");
    await expect(b2).toHaveCSS("font-style", "italic");
    await expect(b2).toHaveCSS("text-decoration-line", "underline");
  });

  test("Align right pushes content to the cell's right edge", async ({ page }) => {
    await gotoXlsxEditor(page);

    await page.getByTestId("cell-A2").click();
    await page.getByTestId("format-align-right").click();

    const a2 = page.getByTestId("cell-A2");
    await expect(a2).toHaveCSS("text-align", "right");
    await expect(page.getByTestId("format-align-right")).toHaveAttribute("aria-pressed", "true");
  });

  test("Format applies across a multi-cell selection", async ({ page }) => {
    await gotoXlsxEditor(page);

    await page.getByTestId("cell-A1").click();
    await page.getByTestId("cell-B2").click({ modifiers: ["Shift"] });
    await expect(page.getByTestId("cell-ref")).toHaveText("A1:B2");

    await page.getByTestId("format-bold").click();

    for (const id of ["cell-A1", "cell-A2", "cell-B1", "cell-B2"]) {
      await expect(page.getByTestId(id)).toHaveCSS("font-weight", "700");
    }
  });
});
