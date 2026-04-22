import { expect, test } from "@playwright/test";
import { gotoEditor } from "./_helpers";

/**
 * Word's "double-click the header band to start typing" affordance.
 *
 * The bundled sample DOCX ships without a `<w:headerReference>` /
 * `<w:footerReference>` of its own, so each page renders an empty
 * "no-part" zone with the dashed `Doppelklicken zum Bearbeiten…`
 * placeholder. A dblclick on either band must:
 *
 *   1. dispatch `docx:create-header-footer-part` (mint a fresh part),
 *   2. swap the inert no-part shell for an editable contenteditable,
 *   3. drop the caret into the freshly minted zone so the user can
 *      type immediately — exactly like Word.
 *
 * Regression coverage for the bug where the page-decorations widget
 * key didn't include the resolved part's identity, so PM reused the
 * old "no-part" widget DOM and the dblclick silently no-op'd.
 */
test.describe("docx header/footer double-click to author", () => {
  test("dblclick on empty header zone mints a header part and focuses caret", async ({ page }) => {
    await gotoEditor(page);

    const headerZone = page.locator(".pm-page-header-band .pm-page-zone-header").first();
    await expect(headerZone).toBeVisible({ timeout: 10_000 });
    await expect(headerZone.locator(".pm-page-zone-no-part")).toBeVisible();

    await headerZone.dblclick();

    const editable = headerZone.locator('[contenteditable="true"]').first();
    await expect(editable).toBeVisible({ timeout: 5_000 });
    await expect(editable).toBeFocused({ timeout: 5_000 });

    await page.keyboard.type("Hello header");
    await expect(editable).toContainText("Hello header");
  });

  test("dblclick on empty footer zone mints a footer part and focuses caret", async ({ page }) => {
    await gotoEditor(page);

    const footerZone = page.locator(".pm-page-footer-band .pm-page-zone-footer").first();
    await expect(footerZone).toBeVisible({ timeout: 10_000 });
    await expect(footerZone.locator(".pm-page-zone-no-part")).toBeVisible();

    await footerZone.dblclick();

    const editable = footerZone.locator('[contenteditable="true"]').first();
    await expect(editable).toBeVisible({ timeout: 5_000 });
    await expect(editable).toBeFocused({ timeout: 5_000 });

    await page.keyboard.type("Hello footer");
    await expect(editable).toContainText("Hello footer");
  });
});
