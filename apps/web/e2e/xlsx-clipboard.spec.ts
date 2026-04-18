import { expect, test } from "@playwright/test";
import { gotoXlsxEditor } from "./_helpers";

/**
 * Phase 13d — system-clipboard bridge.
 *
 * Granting `clipboard-read` + `clipboard-write` lets us round-trip
 * through `navigator.clipboard.read()` / `.write()`, which is what
 * the editor's fallback path uses. Browsers without those grants
 * still get the synchronous `clipboardData` path, but we exercise
 * the async one here because it's the harder one to keep working.
 */
test.use({
  permissions: ["clipboard-read", "clipboard-write"],
});

test.describe("xlsx editor: clipboard", () => {
  test("Cmd+C then Cmd+V duplicates a value into the destination", async ({ page }) => {
    await gotoXlsxEditor(page);

    // Seed a known value at A2.
    await page.getByTestId("cell-A2").click();
    await page.keyboard.type("hello");
    await page.keyboard.press("Enter");
    await expect(page.getByTestId("cell-A2")).toContainText("hello");

    // Re-select A2 and copy.
    await page.getByTestId("cell-A2").click();
    await page.keyboard.press("ControlOrMeta+c");

    // Marching ants overlay confirms the copy registered.
    await expect(page.getByTestId("grid-marching-ants")).toBeVisible();

    // Move to D5 and paste.
    await page.getByTestId("cell-D5").click();
    await page.keyboard.press("ControlOrMeta+v");

    await expect(page.getByTestId("cell-D5")).toContainText("hello");
  });

  test("Escape clears the marching-ants overlay", async ({ page }) => {
    await gotoXlsxEditor(page);

    await page.getByTestId("cell-A2").click();
    await page.keyboard.type("ants");
    await page.keyboard.press("Enter");
    await page.getByTestId("cell-A2").click();
    await page.keyboard.press("ControlOrMeta+c");

    await expect(page.getByTestId("grid-marching-ants")).toBeVisible();
    await page.keyboard.press("Escape");
    await expect(page.getByTestId("grid-marching-ants")).toHaveCount(0);
  });
});
