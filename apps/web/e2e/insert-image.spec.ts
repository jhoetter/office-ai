import { expect, test } from "@playwright/test";
import { gotoEditor } from "./_helpers";

/**
 * P2.4 / W23: smoke-test the image insertion pipeline end-to-end.
 *
 * The dedicated hidden `<input data-testid="image-file-input">` lets us
 * drive the UX without simulating the OS picker. We feed it a tiny but
 * valid PNG, then assert that:
 *
 *  - a real `<img class="pm-image">` lands in the editor (proving the
 *    parser → media resolver → schema → toDOM chain is wired), and
 *  - its `src` is a `data:image/png;base64,...` URL (proving the
 *    resolver actually populated the attribute, not the placeholder
 *    fallback).
 */

// 1×1 transparent PNG, identical to the one used in the docx unit tests.
const PNG_1x1 = Buffer.from([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52, 0x00, 0x00,
  0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4, 0x89, 0x00, 0x00, 0x00,
  0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x62, 0x00, 0x01, 0x00, 0x00, 0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d,
  0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82,
]);

test.describe("editor: image insertion", () => {
  test("uploading a PNG inserts a real <img> with a data: URL", async ({ page }) => {
    await gotoEditor(page);

    const initial = await page.locator(".ProseMirror img.pm-image").count();

    await page.locator('[data-testid="image-file-input"]').setInputFiles({
      name: "smoke.png",
      mimeType: "image/png",
      buffer: PNG_1x1,
    });

    const inserted = page.locator(".ProseMirror img.pm-image").nth(initial);
    await expect(inserted).toBeVisible({ timeout: 5_000 });
    const src = await inserted.getAttribute("src");
    expect(src).toMatch(/^data:image\/png;base64,/);
  });
});
