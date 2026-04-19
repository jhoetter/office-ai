import { expect, test, type Page } from "@playwright/test";

/**
 * PPTX keyboard-shortcut smoke tests. Backed by the
 * {@link usePptxShortcuts} hook on the slide-canvas surface. We
 * exercise the two most observable bindings — `Mod+M` (add slide)
 * and arrow-key nudge — because both produce a clean DOM delta we
 * can assert on without leaking into the agent.
 */
async function gotoPptxEditor(page: Page): Promise<void> {
  // Dev-mode `next dev` compiles each route on first hit, which can
  // exceed Playwright's default expect timeout — give it generous head
  // room so test ordering doesn't determine pass/fail.
  await page.goto("/pptx-editor", { timeout: 60_000 });
  await expect(page.getByTestId("pptx-sidebar")).toBeVisible({ timeout: 60_000 });
  await expect(page.getByTestId("pptx-sidebar").getByRole("button").first()).toBeVisible({
    timeout: 60_000,
  });
}

const thumbnails = (page: Page) => page.getByTestId("pptx-sidebar").getByRole("button");

test.describe("pptx editor: keyboard shortcuts", () => {
  test.setTimeout(60_000);

  async function focusSurface(page: Page): Promise<void> {
    // The slide-canvas section has tabIndex=-1 so it only takes
    // programmatic focus. A bare click on a non-focusable parent
    // doesn't move activeElement, so we drive focus() directly.
    await page.locator('[data-testid="pptx-slide-surface"]').evaluate((el) => {
      (el as HTMLElement).focus();
    });
  }

  test("Mod+M adds a new slide", async ({ page }) => {
    await gotoPptxEditor(page);

    await expect(thumbnails(page)).toHaveCount(1);
    await focusSurface(page);
    await page.keyboard.press("ControlOrMeta+m");

    await expect(thumbnails(page)).toHaveCount(2, { timeout: 5_000 });
  });

  test("Arrow nudges the selected shape's transform", async ({ page }) => {
    await gotoPptxEditor(page);

    // Pick the first shape on the canvas; the SVG renderer wraps each
    // shape in <g class="shape" data-shape-id=…> with a translate().
    const shape = page.locator('[data-testid="pptx-slide-surface"] svg g.shape').first();
    await expect(shape).toBeVisible();
    const beforeTransform = await shape.getAttribute("transform");

    // Clicking a shape now focuses the slide surface implicitly via
    // its pointer-down handler, so the keyboard shortcut works without
    // an extra `focusSurface()` step. Arrow + Shift = +10 px; the
    // larger delta is easier to assert without rounding noise.
    await shape.click();
    await page.keyboard.press("Shift+ArrowRight");

    await expect
      .poll(async () => shape.getAttribute("transform"), { timeout: 5_000 })
      .not.toBe(beforeTransform);
  });

  test("Delete removes the selected shape after a click", async ({ page }) => {
    await gotoPptxEditor(page);

    const shapes = page.locator('[data-testid="pptx-slide-surface"] svg g.shape');
    await expect(shapes.first()).toBeVisible();
    const before = await shapes.count();
    expect(before).toBeGreaterThan(0);

    const target = shapes.first();
    const targetId = await target.getAttribute("data-shape-id");
    expect(targetId).not.toBeNull();

    // Click → select → Delete should remove the shape without needing
    // a separate `focusSurface()` step. Regression for the issue where
    // focus stayed on <body> after clicking an SVG shape, so the
    // Delete shortcut never reached `usePptxShortcuts`.
    await target.click();
    await page.keyboard.press("Delete");

    await expect
      .poll(
        async () =>
          page
            .locator(`[data-testid="pptx-slide-surface"] svg g.shape[data-shape-id="${targetId}"]`)
            .count(),
        { timeout: 5_000 }
      )
      .toBe(0);
  });
});
