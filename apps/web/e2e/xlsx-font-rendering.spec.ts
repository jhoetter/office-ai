import { expect, test } from "@playwright/test";
import { gotoXlsxEditor } from "./_helpers";

/**
 * Font fallback aliasing: when a workbook stores `Calibri` (or any
 * other Microsoft-only family) as the cell font, the cell must still
 * render in something that *looks* like Calibri on a system without
 * Office installed.
 *
 * Two layers are exercised here:
 *
 *   1. The renderer wrap (`apps/web/app/xlsx-editor/styles.ts` →
 *      `wrapFontFamily()`) appends `system-ui, sans-serif` so the
 *      computed `font-family` becomes `"Calibri, system-ui, sans-serif"`.
 *   2. The CSS `@font-face` aliases in `apps/web/app/globals.css`
 *      redefine `Calibri` to load the bundled metric-equivalent twin
 *      (Carlito) when the local system doesn't ship Calibri itself.
 *
 * Pre-fix, layer (1) was missing entirely (the cell got bare
 * `font-family: Calibri`) and layer (2) didn't exist, so non-Office
 * users saw the UA's default Times-style serif inside cells while the
 * toolbar misleadingly claimed "Calibri".
 */
test.describe("xlsx editor: font fallback aliasing", () => {
  test("Calibri-styled cells receive the wrapped fallback chain", async ({ page }) => {
    await gotoXlsxEditor(page);

    // The bundled sample workbook applies the styleSheet's default
    // `<font><name val="Calibri"/></font>` to every populated cell.
    // The grid's per-cell <div> takes that name through `styleForCell`
    // and emits an inline `font-family` style.
    const cell = page.getByTestId("cell-A2");
    await expect(cell).toBeVisible();

    const computedFamily = await cell.evaluate((el) => getComputedStyle(el).fontFamily);
    // The exact whitespace and quoting can vary by browser; what we
    // care about is that the family chain starts with Calibri (so the
    // @font-face alias takes over) and falls through to the generic
    // tail (so unknown families never land on UA serif).
    expect(computedFamily).toMatch(/^["']?Calibri["']?\s*,/);
    expect(computedFamily).toMatch(/sans-serif\s*$/);
  });

  test("the bundled Carlito woff2 actually loads when Calibri is requested", async ({ page }) => {
    await gotoXlsxEditor(page);

    // Allow a couple of frames for the @font-face load to kick off
    // once at least one cell has been painted with `font-family: Calibri`.
    await page.waitForFunction(() => document.fonts.ready.then(() => true), null, {
      timeout: 5_000,
    });

    // FontFaceSet.check returns true iff the browser can resolve a
    // glyph for the requested family/size combination — either via a
    // locally installed font OR via a successfully-loaded @font-face.
    // With our `@font-face { font-family: "Calibri"; src: ... url(carlito-...woff2) }`
    // declaration, this must always return true; pre-fix on a Mac
    // without Office it returned false.
    const calibriResolves = await page.evaluate(() => document.fonts.check('11pt "Calibri"'));
    expect(calibriResolves).toBe(true);
  });
});
