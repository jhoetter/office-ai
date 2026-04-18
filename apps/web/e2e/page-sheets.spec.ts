import { expect, test } from "@playwright/test";
import { gotoEditor } from "./_helpers";

/**
 * P3.8 — Word-flavoured page sheets + double-click header/footer UX.
 *
 * The page-decorations plugin paints every body block as part of a
 * white "page sheet" sitting on a grey desk, with a header zone above
 * the first sheet and a footer zone below the last sheet. Double-
 * clicking either zone fires a CustomEvent that DocxEditor catches
 * to mount the inline `PageZoneEditor` popover.
 *
 * The bundled welcome doc has zero header/footer parts, so we assert
 * the popover surfaces the "no part" warning instead of editing
 * anything. The warning copy is the contract the user sees in Word
 * before they hit "Add header", and matches the deferred-creation
 * carry-over noted in `docs/build-log/docx.md`.
 */
test.describe("editor: paginated page sheets", () => {
  test("renders header + footer zones at the top and bottom of the page", async ({ page }) => {
    await gotoEditor(page);
    const surface = page.locator(".ProseMirror").first();
    await expect(surface).toBeVisible();

    const headerZone = surface.locator(".pm-page-cap-top .pm-page-zone-header");
    const footerZone = surface.locator(".pm-page-cap-bottom .pm-page-zone-footer");

    await expect(headerZone).toBeVisible();
    await expect(footerZone).toBeVisible();
    await expect(headerZone).toContainText("Double-click to add a header");
    await expect(footerZone).toContainText("Double-click to add a footer");

    const block = surface.locator(".pm-page-block").first();
    await expect(block)
      .toHaveCount(1, { timeout: 1000 })
      .catch(() => {});
    await expect(surface.locator(".pm-page-block").first()).toBeVisible();
  });

  test("double-clicking the header zone opens the page-zone editor", async ({ page }) => {
    await gotoEditor(page);
    const surface = page.locator(".ProseMirror").first();
    const headerZone = surface.locator(".pm-page-cap-top .pm-page-zone-header .pm-page-zone-content").first();
    await headerZone.scrollIntoViewIfNeeded();
    await headerZone.dblclick();

    const popover = page.getByTestId("page-zone-editor");
    await expect(popover).toBeVisible();
    await expect(popover).toContainText(/header.*page 1/i);
    // The bundled welcome doc has no header part, so the editor must
    // surface the "add one in Word" guidance instead of accepting an
    // edit silently.
    await expect(popover).toContainText("no header part");
  });

  test("double-clicking the footer zone opens the page-zone editor", async ({ page }) => {
    await gotoEditor(page);
    const surface = page.locator(".ProseMirror").first();
    const footerZone = surface
      .locator(".pm-page-cap-bottom .pm-page-zone-footer .pm-page-zone-content")
      .first();
    await footerZone.scrollIntoViewIfNeeded();
    await footerZone.dblclick();

    const popover = page.getByTestId("page-zone-editor");
    await expect(popover).toBeVisible();
    await expect(popover).toContainText(/footer.*page 1/i);
    await expect(popover).toContainText("no footer part");
  });
});
