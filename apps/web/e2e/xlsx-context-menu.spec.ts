import { expect, test } from "@playwright/test";
import { gotoXlsxEditor } from "./_helpers";

/**
 * Phase 13b — right-click context menus on cells, row headers, and
 * column headers. The menu has to render with the right items for
 * each surface and dispatch a real command when a destructive item
 * is picked.
 */
test.describe("xlsx editor: context menus", () => {
  test("right-click on a body cell opens the cell menu", async ({ page }) => {
    await gotoXlsxEditor(page);

    await page.getByTestId("cell-A2").click({ button: "right" });
    const menu = page.getByTestId("context-menu");
    await expect(menu).toBeVisible();

    // Spot-check items that should show up on a cell menu.
    await expect(page.getByTestId("menu-item-copy")).toBeVisible();
    await expect(page.getByTestId("menu-item-paste")).toBeVisible();
    await expect(page.getByTestId("menu-item-clear-contents")).toBeVisible();
    await expect(page.getByTestId("menu-item-text-to-columns")).toBeVisible();
  });

  test("Escape dismisses the menu", async ({ page }) => {
    await gotoXlsxEditor(page);

    await page.getByTestId("cell-A2").click({ button: "right" });
    await expect(page.getByTestId("context-menu")).toBeVisible();
    await page.keyboard.press("Escape");
    await expect(page.getByTestId("context-menu")).toHaveCount(0);
  });

  test("Clear contents wipes the cell value through the bus", async ({ page }) => {
    await gotoXlsxEditor(page);

    await page.getByTestId("cell-B1").click();
    await expect(page.getByTestId("cell-B1")).toContainText("Score");

    await page.getByTestId("cell-B1").click({ button: "right" });
    await page.getByTestId("menu-item-clear-contents").click();

    await expect(page.getByTestId("cell-B1")).not.toContainText("Score");
    // Bus revision bumped → command actually dispatched.
    await expect(page.getByTestId("revision-badge")).not.toContainText("rev 0");
  });
});
