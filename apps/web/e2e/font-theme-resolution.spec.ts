import { expect, test } from "@playwright/test";
import { focusEditor, gotoEditor, selectParagraphContaining } from "./_helpers";

/**
 * P3.9 — theme-aware font resolution.
 *
 * The bundled welcome doc carries `<w:rFonts w:asciiTheme="majorHAnsi"/>`
 * on every Heading style. Word 2024+ resolves that ref through
 * `word/theme/theme1.xml` to "Aptos Display"; before this batch the
 * editor's resolver had no theme awareness and silently fell through
 * to `docDefaults.rPrDefault` (Calibri), so the toolbar's font picker
 * disagreed with what Word actually rendered on export.
 *
 * The tests park the caret on a target paragraph and assert the
 * toolbar font dropdown reads the typeface Word would render. We use
 * keyboard navigation rather than click() to move the caret because
 * PM's hit-test on the page-decorated surface is occasionally off by
 * a paragraph in headless Chromium — keys that PM has installed
 * `keymap` plugins for (`ArrowDown`, `End`, `Home`) are deterministic.
 */
test.describe("editor: theme-aware font resolution", () => {
  test("Heading 1 toolbar font reads 'Aptos Display' (theme-resolved)", async ({ page }) => {
    await gotoEditor(page);
    await focusEditor(page);
    // Caret starts at doc start → first paragraph (Heading 1).
    await page.keyboard.press("ControlOrMeta+Home");

    const fontSelect = page.locator('select[aria-label="Font family"]');
    await expect(fontSelect).toBeVisible();
    await expect(fontSelect).toHaveValue("Aptos Display");

    // Sanity: paragraph style picker confirms we are actually on H1.
    await expect(page.locator('select[aria-label="Paragraph style"]')).toHaveValue("Heading1");
  });

  test("Plain body text toolbar font reads 'Calibri' (docDefaults literal)", async ({ page }) => {
    await gotoEditor(page);
    await focusEditor(page);
    // Walk caret down to the second paragraph (the first body text).
    // Welcome doc layout: H1, body, body.
    await page.keyboard.press("ControlOrMeta+Home");
    await page.keyboard.press("ArrowDown");

    // Confirm we landed on a Normal paragraph before asserting font.
    await expect(page.locator('select[aria-label="Paragraph style"]')).toHaveValue("Normal");
    const fontSelect = page.locator('select[aria-label="Font family"]');
    await expect(fontSelect).toHaveValue("Calibri");
  });

  test("Heading 2 also resolves through majorHAnsi → 'Aptos Display'", async ({ page }) => {
    await gotoEditor(page);
    await focusEditor(page);
    // Welcome doc no longer ships an H2 by default. Park the caret on
    // a body paragraph and convert it to Heading 2 via the toolbar's
    // paragraph style picker so the test still exercises the
    // majorHAnsi → 'Aptos Display' resolver path.
    await selectParagraphContaining(page, "tiny synthetic .docx generated");
    const styleSelect = page.locator('select[aria-label="Paragraph style"]');
    await styleSelect.selectOption("Heading2");

    await expect(styleSelect).toHaveValue("Heading2");
    const fontSelect = page.locator('select[aria-label="Font family"]');
    await expect(fontSelect).toHaveValue("Aptos Display");
  });
});
