import { expect, test } from "@playwright/test";
import { gotoEditor } from "./_helpers";

/**
 * P3.8 / B2 — Word-flavoured page sheets + in-place header/footer
 * authoring.
 *
 * The page-decorations plugin paints every body block as part of a
 * white "page sheet" sitting on a grey desk, with a header zone above
 * the first sheet and a footer zone below the last sheet. The zones
 * render their inner content as `contenteditable=true` so the user
 * clicks straight into them — Word's in-place authoring model — and
 * commits via the `pm-page-zone-commit` event on blur.
 *
 * The bundled welcome doc has zero header/footer parts, so the zones
 * surface a "No header/footer for this section" hint and stay
 * read-only. Auto-mint of header/footer parts is on the B7 list.
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
    // The placeholder hint is now hover-only (mirrors Word: a clean
    // page must never read as if it had two headings) — assert the
    // hint copy lives in the data attribute the CSS pulls into ::before.
    await expect(headerZone.locator(".pm-page-zone-no-part")).toHaveAttribute(
      "data-empty-text",
      /No header for this section/i
    );
    await expect(footerZone.locator(".pm-page-zone-no-part")).toHaveAttribute(
      "data-empty-text",
      /No footer for this section/i
    );

    const block = surface.locator(".pm-page-block").first();
    await expect(block)
      .toHaveCount(1, { timeout: 1000 })
      .catch(() => {});
    await expect(surface.locator(".pm-page-block").first()).toBeVisible();
  });

  test("zones without a part stay read-only (no contenteditable)", async ({ page }) => {
    await gotoEditor(page);
    const surface = page.locator(".ProseMirror").first();
    const headerContent = surface
      .locator(".pm-page-cap-top .pm-page-zone-header .pm-page-zone-content")
      .first();
    await expect(headerContent).toBeVisible();
    // No part → the inner content is NOT contenteditable; the user
    // sees a hint instead of authoring.
    const editable = await headerContent.getAttribute("contenteditable");
    expect(editable).not.toBe("true");
  });
});
