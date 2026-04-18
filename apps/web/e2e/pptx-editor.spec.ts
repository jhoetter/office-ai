import { expect, test, type Page } from "@playwright/test";

/**
 * F1.3 — Playwright smoke for `/pptx-editor`.
 *
 * Mirrors the DOCX `toolbar.spec.ts` pattern: navigate, wait for the
 * sample deck to mount, then exercise the toolbar / sidebar / agent
 * panel surface that PptxEditor.tsx exposes. We assert observable
 * UI deltas (slide thumbnails, SVG shape count, agent rationale toast)
 * rather than reaching into the agent — keeps the spec resilient to
 * model refactors.
 */

async function gotoPptxEditor(page: Page): Promise<void> {
  await page.goto("/pptx-editor");
  // PptxEditor mounts the agent on first render and only then renders
  // the SVG canvas. Wait for at least one slide thumbnail to appear.
  await expect(page.getByTestId("pptx-sidebar")).toBeVisible({ timeout: 15_000 });
  await expect(page.getByTestId("pptx-sidebar").getByRole("button").first()).toBeVisible({ timeout: 15_000 });
}

function thumbnails(page: Page) {
  return page.getByTestId("pptx-sidebar").getByRole("button");
}

function shapesOnCanvas(page: Page) {
  // Each shape is wrapped in `<g class="shape …" data-shape-id=…>` —
  // see packages/pptx/src/renderer/svg/shapes.ts groupOpen().
  return page.locator(".pptx-editor svg g.shape");
}

test.describe("pptx-editor route", () => {
  test("mounts the sample deck with one slide and renders shapes", async ({ page }) => {
    await gotoPptxEditor(page);

    await expect(thumbnails(page)).toHaveCount(1);
    // Sample deck has at least the title text shape.
    await expect(shapesOnCanvas(page).first()).toBeVisible();
  });

  test("toolbar 'Add slide' adds a thumbnail and 'Delete' becomes enabled", async ({ page }) => {
    await gotoPptxEditor(page);

    const deleteBtn = page.getByRole("button", { name: "Delete" });
    await expect(deleteBtn).toBeDisabled();

    await page.getByRole("button", { name: "Add slide" }).click();
    await expect(thumbnails(page)).toHaveCount(2);
    await expect(deleteBtn).toBeEnabled();
  });

  test("toolbar 'Text box' adds a new shape on the active slide", async ({ page }) => {
    await gotoPptxEditor(page);

    const before = await shapesOnCanvas(page).count();
    await page.getByRole("button", { name: "Text box" }).click();

    await expect
      .poll(async () => shapesOnCanvas(page).count(), {
        timeout: 5_000,
      })
      .toBeGreaterThan(before);
  });

  test("'Bold' toolbar toggle survives without tossing an error toast", async ({ page }) => {
    await gotoPptxEditor(page);

    await page.getByRole("button", { name: "Bold", exact: true }).click();
    // No error toast (the warn/error toasts have role="status" and are
    // mounted in the bottom-center stack). We give the bus a beat to
    // settle so the assertion isn't racy.
    await page.waitForTimeout(300);
    const errorToasts = page.getByRole("status").filter({ hasText: /error/i });
    await expect(errorToasts).toHaveCount(0);
  });

  test("zoom slider rescales the slide canvas and resets to 100%", async ({ page }) => {
    await gotoPptxEditor(page);

    const canvas = page.getByTestId("pptx-slide-canvas");
    await expect(canvas).toBeVisible();
    await expect(canvas).toHaveAttribute("data-zoom", "1.00");

    const slider = page.getByTestId("pptx-zoom-slider");
    // React intercepts native value setters on `<input>` to track changes;
    // calling the prototype setter directly is the supported way to drive
    // a controlled range input from outside the React event system.
    await slider.evaluate((el) => {
      const input = el as HTMLInputElement;
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value")?.set;
      setter?.call(input, "1.5");
      input.dispatchEvent(new Event("input", { bubbles: true }));
      input.dispatchEvent(new Event("change", { bubbles: true }));
    });
    await expect(canvas).toHaveAttribute("data-zoom", "1.50");

    await page.getByTestId("pptx-zoom-reset").click();
    await expect(canvas).toHaveAttribute("data-zoom", "1.00");
  });
});
