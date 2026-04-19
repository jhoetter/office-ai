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

  test("shared text-format bar disables Bold until there's a selection", async ({ page }) => {
    await gotoPptxEditor(page);

    // The shared TextFormatBar (Phase 6) drives `disabled` from
    // `provider.hasSelection()`. With no shape being edited the bar
    // should expose Bold/Italic/Underline as visible-but-disabled so
    // the toolbar remains discoverable, mirroring Word/Excel.
    const bold = page.getByRole("button", { name: "Bold", exact: true });
    await expect(bold).toBeVisible();
    await expect(bold).toBeDisabled();

    const italic = page.getByRole("button", { name: "Italic", exact: true });
    await expect(italic).toBeDisabled();

    // Sanity check: no error toast appeared just by mounting the bar.
    const errorToasts = page.getByRole("status").filter({ hasText: /error/i });
    await expect(errorToasts).toHaveCount(0);
  });

  test("connector menu arms the tool, surfaces a banner, and Esc exits", async ({ page }) => {
    await gotoPptxEditor(page);

    const trigger = page.getByTestId("pptx-connector-menu-trigger");
    await expect(trigger).toHaveAttribute("data-active-type", "");

    await trigger.click();
    await page.getByTestId("pptx-connector-elbow").click();

    // The trigger now reports the armed type and the canvas shows the
    // tool banner. We deliberately don't synthesise a draw gesture
    // here — the snap walker depends on real layout boxes which jsdom
    // / Playwright headless can't faithfully reproduce. The banner
    // presence is the single most reliable observable signal that
    // the tool entered armed mode.
    await expect(trigger).toHaveAttribute("data-active-type", "elbow");
    await expect(page.getByTestId("pptx-connector-tool-banner")).toBeVisible();

    await page.keyboard.press("Escape");
    await expect(trigger).toHaveAttribute("data-active-type", "");
    await expect(page.getByTestId("pptx-connector-tool-banner")).toHaveCount(0);
  });

  test("re-clicking the active connector type exits the tool", async ({ page }) => {
    await gotoPptxEditor(page);

    const trigger = page.getByTestId("pptx-connector-menu-trigger");
    await trigger.click();
    await page.getByTestId("pptx-connector-straight").click();
    await expect(trigger).toHaveAttribute("data-active-type", "straight");

    await trigger.click();
    const item = page.getByTestId("pptx-connector-straight");
    await expect(item).toHaveAttribute("aria-pressed", "true");
    await item.click();
    await expect(trigger).toHaveAttribute("data-active-type", "");
  });

  test("arming the connector tool reveals the connectable-shapes overlay", async ({ page }) => {
    await gotoPptxEditor(page);

    // Overlay is hidden by default — only mounts while the tool is
    // armed (or while an existing endpoint is being edited).
    await expect(page.getByTestId("pptx-connector-affordance-overlay")).toHaveCount(0);

    await page.getByTestId("pptx-connector-menu-trigger").click();
    await page.getByTestId("pptx-connector-elbow").click();

    await expect(page.getByTestId("pptx-connector-affordance-overlay")).toBeVisible();

    await page.keyboard.press("Escape");
    await expect(page.getByTestId("pptx-connector-affordance-overlay")).toHaveCount(0);
  });

  test("the connector menu exposes the three connector types", async ({ page }) => {
    await gotoPptxEditor(page);

    await page.getByTestId("pptx-connector-menu-trigger").click();
    await expect(page.getByTestId("pptx-connector-straight")).toBeVisible();
    await expect(page.getByTestId("pptx-connector-elbow")).toBeVisible();
    await expect(page.getByTestId("pptx-connector-curved")).toBeVisible();
  });

  test("switching connector type from the menu updates the active type", async ({ page }) => {
    await gotoPptxEditor(page);

    const trigger = page.getByTestId("pptx-connector-menu-trigger");

    await trigger.click();
    await page.getByTestId("pptx-connector-straight").click();
    await expect(trigger).toHaveAttribute("data-active-type", "straight");

    await trigger.click();
    await page.getByTestId("pptx-connector-curved").click();
    await expect(trigger).toHaveAttribute("data-active-type", "curved");
  });

  test("text alignment buttons re-render the active text shape with the chosen alignment", async ({
    page,
  }) => {
    await gotoPptxEditor(page);

    // Pick the first text shape on the canvas. The sample deck always
    // ships a title shape with at least one paragraph, so this is a
    // stable target.
    const shape = page.locator(".pptx-editor svg g.shape").first();
    await expect(shape).toBeVisible();
    await shape.click();

    // The cluster gates on `hasTextShapeFocus`; with the text shape
    // selected it should now be enabled. We assert disabled-state
    // inversion via aria-pressed instead of `disabled` to avoid being
    // brittle against the surrounding wrapper layout.
    const alignCenter = page.getByTestId("pptx-text-align-center");
    const alignRight = page.getByTestId("pptx-text-align-right");
    const anchorMiddle = page.getByTestId("pptx-text-anchor-middle");
    const anchorBottom = page.getByTestId("pptx-text-anchor-bottom");
    await expect(alignCenter).toBeEnabled();
    await expect(anchorMiddle).toBeEnabled();

    // Read paragraph element under the active shape's foreignObject
    // for verification. The renderer emits one `<div>` per paragraph
    // with `text-align:<value>` baked into the inline style.
    const paragraph = shape.locator("foreignObject div div").first();
    await expect(paragraph).toBeVisible();

    await alignCenter.click();
    await expect
      .poll(async () => paragraph.getAttribute("style"), { timeout: 5_000 })
      .toMatch(/text-align:\s*center/);
    await expect(alignCenter).toHaveAttribute("aria-pressed", "true");

    await alignRight.click();
    await expect
      .poll(async () => paragraph.getAttribute("style"), { timeout: 5_000 })
      .toMatch(/text-align:\s*right/);
    await expect(alignRight).toHaveAttribute("aria-pressed", "true");

    // The vertical anchor lives on the foreignObject's outer flex
    // container — `justify-content` switches between flex-start (top),
    // center (middle), and flex-end (bottom).
    const anchorContainer = shape.locator("foreignObject > div").first();
    await anchorMiddle.click();
    await expect
      .poll(async () => anchorContainer.getAttribute("style"), { timeout: 5_000 })
      .toMatch(/justify-content:\s*center/);
    await expect(anchorMiddle).toHaveAttribute("aria-pressed", "true");

    await anchorBottom.click();
    await expect
      .poll(async () => anchorContainer.getAttribute("style"), { timeout: 5_000 })
      .toMatch(/justify-content:\s*flex-end/);
    await expect(anchorBottom).toHaveAttribute("aria-pressed", "true");
  });

  test("zoom controls rescale the slide canvas and reset to 100%", async ({ page }) => {
    await gotoPptxEditor(page);

    const canvas = page.getByTestId("pptx-slide-canvas");
    await expect(canvas).toBeVisible();
    await expect(canvas).toHaveAttribute("data-zoom", "1.00");

    // The shared `ZoomControl` lives in the status bar across all
    // products: minus button, percent label (click = reset to 100 %),
    // plus button. Five plus-clicks = 1.0 → 1.5.
    const zoomIn = page.getByTestId("zoom-in");
    for (let i = 0; i < 5; i++) await zoomIn.click();
    await expect(canvas).toHaveAttribute("data-zoom", "1.50");

    await page.getByTestId("zoom-percent").click();
    await expect(canvas).toHaveAttribute("data-zoom", "1.00");
  });
});
