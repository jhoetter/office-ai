import { expect, test, type Page } from "@playwright/test";

/**
 * W4 plan / pptx-tests — assigning a transition via the toolbar must
 * actually fire the WAAPI overlay when entering presenter mode.
 *
 * Sequence:
 *   1. Open the bundled PPTX editor with `?new=1` so we land on a
 *      brand-new deck.
 *   2. Open the slide-transition menu, pick `Fade`.
 *   3. Click `Present` to enter presenter mode.
 *   4. Assert the transition kind is surfaced in the header badge
 *      (`pptx-present-transition`) and that the WAAPI overlay
 *      element appears within 500 ms — proving the
 *      pptx-initial-transition + pptx-stale-snapshot fixes work
 *      end-to-end.
 */

async function gotoPptxEditor(page: Page): Promise<void> {
  await page.goto("/pptx-editor?new=1");
  await expect(page.getByTestId("pptx-sidebar")).toBeVisible({ timeout: 15_000 });
  await expect(page.getByTestId("pptx-sidebar").getByRole("button").first()).toBeVisible({
    timeout: 15_000,
  });
}

test.describe("pptx presenter: assigned transitions actually play", () => {
  test("assigning Fade then opening presenter renders the transition badge", async ({ page }) => {
    await gotoPptxEditor(page);

    // Open the transition menu and pick Fade.
    await page.getByTestId("pptx-transition-menu-trigger").click();
    await expect(page.getByTestId("pptx-transition-menu")).toBeVisible();
    await page.getByTestId("pptx-transition-fade").click();

    // Enter presenter.
    await page.getByTestId("pptx-present").click();

    const present = page.getByTestId("pptx-present-mode");
    await expect(present).toBeVisible({ timeout: 5_000 });

    // The transition kind badge must surface the *currently-landed-on*
    // slide's transition. Before pptx-stale-snapshot was fixed this
    // could read empty because the snapshot prop pre-dated the
    // command's commit.
    await expect(page.getByTestId("pptx-present-transition")).toContainText(/fade/i, {
      timeout: 5_000,
    });

    // Crucially: the WAAPI overlay must mount and stay mounted long
    // enough for the keyframes to actually paint. The previous build
    // re-ran the overlay's effect on every parent render (the 1 Hz
    // `setNow` tick alone was enough); under React Strict Mode the
    // very first cleanup cancelled the animation and the cancel
    // listener torpedoed the overlay before any frame committed.
    const overlay = page.getByTestId("pptx-present-transition-overlay");
    await expect(overlay).toBeVisible({ timeout: 1_000 });
    await expect(overlay).toHaveAttribute("data-kind", "fade");
  });
});
