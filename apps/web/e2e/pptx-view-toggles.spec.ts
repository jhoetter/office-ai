import { expect, test, type Page } from "@playwright/test";

/**
 * View-tab toggles in the PPTX editor.
 *
 * Covers the two PowerPoint-style preferences added by the rulers /
 * gridlines feature:
 *
 * - "Rulers" — horizontal + vertical scale around the slide card.
 *   Default ON (matches PowerPoint), persisted to localStorage so a
 *   reload restores the user's choice.
 * - "Gridlines" — translucent grid overlay across the slide rectangle.
 *   Default OFF, persisted the same way.
 *
 * Both are visual UI prefs, not part of the deck — so we don't go
 * through the agent / save flow, just assert DOM presence and the
 * `aria-pressed` indicator on the toolbar buttons.
 */

async function gotoPptxEditor(page: Page): Promise<void> {
  await page.goto("/pptx-editor");
  await expect(page.getByTestId("pptx-sidebar")).toBeVisible({ timeout: 15_000 });
  await expect(page.getByTestId("pptx-sidebar").getByRole("button").first()).toBeVisible({
    timeout: 15_000,
  });
}

async function openViewTab(page: Page): Promise<void> {
  await page.getByTestId("ribbon-tab-view").click();
}

test.describe("pptx view-tab toggles", () => {
  test("rulers visible by default and can be hidden / re-shown", async ({ page }) => {
    await gotoPptxEditor(page);
    await expect(page.getByTestId("pptx-ruler-h")).toBeVisible();
    await expect(page.getByTestId("pptx-ruler-v")).toBeVisible();

    await openViewTab(page);
    const rulersToggle = page.getByTestId("pptx-rulers-toggle");
    await expect(rulersToggle).toHaveAttribute("aria-pressed", "true");

    await rulersToggle.click();
    await expect(rulersToggle).toHaveAttribute("aria-pressed", "false");
    await expect(page.getByTestId("pptx-ruler-h")).toHaveCount(0);
    await expect(page.getByTestId("pptx-ruler-v")).toHaveCount(0);

    await rulersToggle.click();
    await expect(page.getByTestId("pptx-ruler-h")).toBeVisible();
    await expect(page.getByTestId("pptx-ruler-v")).toBeVisible();
  });

  test("gridlines hidden by default and can be toggled on", async ({ page }) => {
    await gotoPptxEditor(page);
    await expect(page.getByTestId("pptx-grid-overlay")).toHaveCount(0);

    await openViewTab(page);
    const gridToggle = page.getByTestId("pptx-grid-toggle");
    await expect(gridToggle).toHaveAttribute("aria-pressed", "false");

    await gridToggle.click();
    await expect(gridToggle).toHaveAttribute("aria-pressed", "true");
    // The grid overlay lives inside the dangerously-set slide SVG so
    // it's painted as a `<g>` element with a stable test id.
    await expect(page.locator('[data-testid="pptx-grid-overlay"]')).toHaveCount(1);

    await gridToggle.click();
    await expect(page.locator('[data-testid="pptx-grid-overlay"]')).toHaveCount(0);
  });

  test("toggle state survives a reload (localStorage persistence)", async ({ page }) => {
    await gotoPptxEditor(page);
    await openViewTab(page);

    // Hide rulers, show grid — neither matches the default — then reload.
    await page.getByTestId("pptx-rulers-toggle").click();
    await page.getByTestId("pptx-grid-toggle").click();

    await page.reload();
    await expect(page.getByTestId("pptx-sidebar")).toBeVisible({ timeout: 15_000 });

    await expect(page.getByTestId("pptx-ruler-h")).toHaveCount(0);
    await expect(page.locator('[data-testid="pptx-grid-overlay"]')).toHaveCount(1);

    await openViewTab(page);
    await expect(page.getByTestId("pptx-rulers-toggle")).toHaveAttribute("aria-pressed", "false");
    await expect(page.getByTestId("pptx-grid-toggle")).toHaveAttribute("aria-pressed", "true");
  });
});
