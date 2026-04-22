import { expect, test, type Page } from "@playwright/test";

/**
 * PPTX ribbon — manual tab navigation. The ribbon root keeps the
 * `data-testid="pptx-toolbar"` selector so existing PPTX specs
 * continue to resolve; the active tab id is exposed via
 * `data-active-tab`. We exercise tab switching plus the contextual
 * Übergänge tab where the TransitionMenu lives.
 */

async function gotoPptxEditor(page: Page): Promise<void> {
  await page.goto("/pptx-editor");
  await expect(page.getByTestId("pptx-sidebar")).toBeVisible({ timeout: 15_000 });
  await expect(page.getByTestId("pptx-sidebar").getByRole("button").first()).toBeVisible({
    timeout: 15_000,
  });
}

test.describe("pptx ribbon", () => {
  test("renders the persistent tab strip with Start as default", async ({ page }) => {
    await gotoPptxEditor(page);
    const ribbon = page.getByTestId("pptx-toolbar");
    await expect(ribbon).toBeVisible();
    await expect(ribbon).toHaveAttribute("data-active-tab", "start");

    await expect(page.getByTestId("ribbon-tab-start")).toBeVisible();
    await expect(page.getByTestId("ribbon-tab-insert")).toBeVisible();
    await expect(page.getByTestId("ribbon-tab-transitions")).toBeVisible();
    await expect(page.getByTestId("ribbon-tab-animations")).toBeVisible();
    await expect(page.getByTestId("ribbon-tab-slideshow")).toBeVisible();
    await expect(page.getByTestId("ribbon-tab-view")).toBeVisible();
  });

  test("switching to Übergänge surfaces the TransitionMenu trigger", async ({ page }) => {
    await gotoPptxEditor(page);
    await page.getByTestId("ribbon-tab-transitions").click();
    const ribbon = page.getByTestId("pptx-toolbar");
    await expect(ribbon).toHaveAttribute("data-active-tab", "transitions");
    await expect(page.getByTestId("pptx-transition-menu-trigger")).toBeVisible();
  });

  test("switching to Bildschirmpräsentation reveals the Present action", async ({ page }) => {
    await gotoPptxEditor(page);
    await page.getByTestId("ribbon-tab-slideshow").click();
    await expect(page.getByTestId("pptx-present-from-ribbon")).toBeVisible();
  });

  test("manual tab activation overrides an auto-active contextual tab", async ({ page }) => {
    // Office-style "I want to navigate elsewhere even though there's
    // a contextual tab open": adding a shape auto-activates
    // Formformat (shape-tools), but clicking Animationen must switch
    // to it and stay there as long as the shape selection is
    // unchanged. Regression for the case where the contextual tab
    // re-stole the active state on every render.
    await gotoPptxEditor(page);
    const ribbon = page.getByTestId("pptx-toolbar");

    // Insert a rounded rectangle from the Einfügen → Shape menu so the
    // shape is selected on creation and Formformat auto-activates.
    await page.getByTestId("ribbon-tab-insert").click();
    await page.getByTestId("pptx-shape-menu-trigger").click();
    await page.getByRole("menuitem", { name: "Rounded Rectangle" }).click();
    await expect(ribbon).toHaveAttribute("data-active-tab", "shape-tools");
    await expect(page.getByTestId("ribbon-tab-shape-tools")).toBeVisible();

    // Click the persistent Animationen tab. The contextual tab stays
    // visible but the user's pinned choice wins.
    await page.getByTestId("ribbon-tab-animations").click();
    await expect(ribbon).toHaveAttribute("data-active-tab", "animations");
    await expect(page.getByTestId("ribbon-tab-shape-tools")).toBeVisible();
    await expect(page.getByTestId("pptx-open-animations-rail")).toBeVisible();

    // And we can navigate to other persistent tabs too — Start —
    // proving the override isn't tab-specific.
    await page.getByTestId("ribbon-tab-start").click();
    await expect(ribbon).toHaveAttribute("data-active-tab", "start");
    await expect(page.getByTestId("ribbon-tab-shape-tools")).toBeVisible();
  });
});
