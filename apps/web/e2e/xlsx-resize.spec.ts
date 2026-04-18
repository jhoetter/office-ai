import { expect, test } from "@playwright/test";
import { gotoXlsxEditor } from "./_helpers";

/**
 * Drag the column / row resize handles → `xlsx:set-column-width` /
 * `xlsx:set-row-height`. Verifies the grid renders the variable
 * geometry and the change is dispatched through the command bus.
 */
test.describe("xlsx editor: column / row resizing", () => {
  test("Dragging the A-column resize handle widens the column", async ({ page }) => {
    await gotoXlsxEditor(page);

    const handle = page.getByTestId("col-resize-A");
    const before = await page.getByTestId("col-header-A").boundingBox();
    expect(before).not.toBeNull();
    const startX = (before!.x + before!.width) - 1;
    const y = before!.y + before!.height / 2;

    // Drag 80px to the right.
    await page.mouse.move(startX, y);
    await page.mouse.down();
    await page.mouse.move(startX + 80, y);
    await page.mouse.up();

    // Column header now wider; revision ticked.
    await expect(async () => {
      const after = await page.getByTestId("col-header-A").boundingBox();
      expect(after).not.toBeNull();
      expect(after!.width).toBeGreaterThan(before!.width + 50);
    }).toPass({ timeout: 4000 });
    await expect(page.getByTestId("revision-badge")).toContainText("rev 1");

    // The handle remains usable for the next drag.
    await expect(handle).toBeAttached();
  });

  test("Dragging row 1's resize handle taller increases its height", async ({ page }) => {
    await gotoXlsxEditor(page);

    const before = await page.getByTestId("row-header-1").boundingBox();
    expect(before).not.toBeNull();
    const x = before!.x + before!.width / 2;
    const startY = (before!.y + before!.height) - 1;

    await page.mouse.move(x, startY);
    await page.mouse.down();
    await page.mouse.move(x, startY + 40);
    await page.mouse.up();

    await expect(async () => {
      const after = await page.getByTestId("row-header-1").boundingBox();
      expect(after).not.toBeNull();
      expect(after!.height).toBeGreaterThan(before!.height + 20);
    }).toPass({ timeout: 4000 });
    await expect(page.getByTestId("revision-badge")).toContainText("rev 1");
  });
});
