import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";
import { gotoXlsxEditor } from "./_helpers";

/**
 * Phase 13g — smart fill handle. The handle is a 7×7 violet square
 * pinned to the bottom-right corner of the marquee. Dragging it past
 * the source rectangle should:
 *
 *   1. Render the live `grid-fill-preview` overlay.
 *   2. Dispatch `xlsx:fill-range` on mouse-up.
 *   3. Extrapolate the series into the new cells.
 */
test.describe("xlsx editor: smart fill handle", () => {
  /**
   * Drag the fill handle from its current position by the given
   * pixel deltas. Uses raw mouse events so we exercise the same
   * mousedown / mousemove / mouseup pipeline the production code
   * listens on.
   */
  async function dragFillHandle(page: Page, dx: number, dy: number): Promise<void> {
    const handle = page.getByTestId("grid-fill-handle");
    const box = await handle.boundingBox();
    if (!box) throw new Error("fill handle not visible");
    const startX = box.x + box.width / 2;
    const startY = box.y + box.height / 2;
    await page.mouse.move(startX, startY);
    await page.mouse.down();
    // Step in halves to give React a chance to throttle properly.
    await page.mouse.move(startX + dx / 2, startY + dy / 2, { steps: 6 });
    await page.mouse.move(startX + dx, startY + dy, { steps: 12 });
    await page.mouse.up();
  }

  test("drag-down extrapolates a numeric series", async ({ page }) => {
    await gotoXlsxEditor(page);

    // Seed a 1, 2 source so the numeric detector picks step 1.
    await page.getByTestId("cell-D2").click();
    await page.keyboard.type("1");
    await page.keyboard.press("Enter");
    await page.keyboard.type("2");
    await page.keyboard.press("Enter");

    // Re-select D2:D3 as the source.
    await page.getByTestId("cell-D2").click();
    await page.getByTestId("cell-D3").click({ modifiers: ["Shift"] });

    await expect(page.getByTestId("grid-fill-handle")).toBeVisible();
    // Each row is ROW_HEIGHT = 22px; drag down ~3 rows.
    await dragFillHandle(page, 0, 22 * 3);

    await expect(page.getByTestId("cell-D4")).toContainText("3");
    await expect(page.getByTestId("cell-D5")).toContainText("4");
    await expect(page.getByTestId("cell-D6")).toContainText("5");
  });

  test("preview overlay is visible during the drag", async ({ page }) => {
    await gotoXlsxEditor(page);

    await page.getByTestId("cell-E2").click();
    await page.keyboard.type("10");
    await page.keyboard.press("Enter");
    await page.getByTestId("cell-E2").click();

    const handle = page.getByTestId("grid-fill-handle");
    const box = await handle.boundingBox();
    if (!box) throw new Error("fill handle not visible");
    const startX = box.x + box.width / 2;
    const startY = box.y + box.height / 2;
    await page.mouse.move(startX, startY);
    await page.mouse.down();
    await page.mouse.move(startX, startY + 60, { steps: 8 });

    await expect(page.getByTestId("grid-fill-preview")).toBeVisible();

    await page.mouse.up();
  });
});
