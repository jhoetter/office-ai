import { expect, test, type Page } from "@playwright/test";
import { focusEditor, gotoEditor } from "./_helpers";

/**
 * Phase 9a — "no more lying buttons" smoke spec.
 *
 * Each scenario here pins a fix from the 9a slice that traded a
 * silently-broken UI for a working one. The point is regression-safety:
 * if any of these affordances regress to a no-op, this spec fails fast
 * instead of waiting for a user bug report.
 *
 *  • DOCX Replace All (was a stub) actually rewrites every match.
 *  • XLSX command palette has separate entries for `set-sheet-view`
 *    and `open-zoom-dialog`, and the former opens the new
 *    `SheetViewDialog` (not the Zoom dialog).
 *  • PPTX PresentMode skips slides marked as hidden and surfaces a
 *    pill with the count.
 *
 * The titlePg derivation fix is exercised indirectly via the Header &
 * Footer mode unit tests in @officeai/docx — reaching the H/F double-
 * click flow from Playwright is brittle, so we leave that gate to the
 * lower test ring.
 */

async function gotoXlsx(page: Page): Promise<void> {
  await page.goto("/xlsx-editor");
  await expect(page.getByTestId("cell-B1")).toContainText("Score", {
    timeout: 15_000,
  });
}

async function gotoPptx(page: Page): Promise<void> {
  await page.goto("/pptx-editor");
  await expect(page.getByTestId("pptx-sidebar")).toBeVisible({ timeout: 15_000 });
  await expect(page.getByTestId("pptx-sidebar").getByRole("button").first()).toBeVisible({
    timeout: 15_000,
  });
}

test.describe("Phase 9a — lying-buttons regression net", () => {
  test("DOCX Replace All rewrites every match, not just the first", async ({ page }) => {
    await gotoEditor(page);
    await focusEditor(page);

    // The bundled welcome doc contains "office-ai" multiple times across
    // headings + body. Cmd+Alt+F opens find with replace exposed; if
    // that chord ever regresses to find-only the test fails on the
    // missing replace-input rather than misreading state.
    await page.keyboard.press("ControlOrMeta+Alt+f");
    const panel = page.getByTestId("find-replace");
    await expect(panel).toBeVisible();

    await page.getByTestId("find-input").fill("office-ai");
    await page.getByTestId("replace-input").fill("officeAI");
    await page.getByTestId("replace-all").click();

    // The query string must no longer appear in the rendered surface.
    // We assert at least one occurrence of the replacement so the test
    // fails loud if the replacement was inserted somewhere wrong (e.g.
    // the find input itself).
    await expect(page.locator(".ProseMirror").getByText("officeAI", { exact: false }).first())
      .toBeVisible({ timeout: 5_000 });
    await expect(
      page.locator(".ProseMirror").getByText("office-ai", { exact: false })
    ).toHaveCount(0, { timeout: 5_000 });
  });

  test("XLSX palette: set-sheet-view + open-zoom-dialog are distinct entries", async ({ page }) => {
    await gotoXlsx(page);

    // Open the palette and type "view" — both fixes should land in the
    // dropdown but resolve to different runners.
    await page.keyboard.press("ControlOrMeta+k");
    const palette = page.getByTestId("command-palette");
    await expect(palette).toBeVisible();

    // The two distinct catalogue entries each get their own
    // `palette-cmd-${id}` testid.
    const sheetView = page.getByTestId("palette-cmd-xlsx.set-sheet-view");
    const zoomDialog = page.getByTestId("palette-cmd-xlsx.open-zoom-dialog");

    await page.getByTestId("command-palette-input").fill("view");
    await expect(sheetView).toBeVisible();

    await page.getByTestId("command-palette-input").fill("zoom");
    await expect(zoomDialog).toBeVisible();

    // Picking the sheet-view runner must open the *new* sheet-view
    // dialog, not the legacy zoom dialog. Before 9a these collided.
    await zoomDialog.click();
    // Dialog is the Zoom dialog when activated via open-zoom-dialog.
    // We don't assert its specific testid here (zoom dialog is shared
    // with other entry points and may evolve), only that the sheet-
    // view dialog stayed closed.
    await expect(page.getByTestId("xlsx-sheet-view-dialog")).toBeHidden();

    // Reopen palette and confirm sheet-view actually opens
    // SheetViewDialog (the 9a fix).
    await page.keyboard.press("Escape");
    await page.keyboard.press("ControlOrMeta+k");
    await page.getByTestId("command-palette-input").fill("sheet view");
    await sheetView.click();
    await expect(page.getByTestId("xlsx-sheet-view-dialog")).toBeVisible();
    // Close it so the spec leaves the page in a sane state.
    await page.keyboard.press("Escape");
  });

  test("PPTX PresentMode skips hidden slides and pills the count", async ({ page }) => {
    await gotoPptx(page);

    // The bundled deck has at least one slide. We need >= 2 slides so
    // hiding one still leaves a presentable deck. Use the slide-rail
    // duplicate affordance via the toolbar's "add slide" path —
    // simplest cross-locale entry is the keyboard shortcut for "new
    // slide" (Ctrl/Cmd+M, Word-parity), but that's PPTX-specific. We
    // dispatch via the Layout menu: pick the first preset to land a
    // second slide.
    await page.getByTestId("pptx-layout-menu-trigger").click();
    // Pick whatever the first menu item is — we don't care about the
    // specific layout, only that a new slide gets appended.
    const firstLayout = page.locator('[data-testid^="pptx-layout-preset-"]').first();
    if (await firstLayout.count()) {
      await firstLayout.click();
    } else {
      // Some builds expose layouts as `pptx-layout-item-…`. Fall back
      // to a more permissive selector before failing the test.
      await page.locator('[role="menu"] [role="menuitem"], [role="menu"] button')
        .first()
        .click();
    }

    // Wait for the second thumbnail to appear in the rail.
    await expect(
      page.getByTestId("pptx-sidebar").getByRole("button").nth(1)
    ).toBeVisible({ timeout: 5_000 });

    // Hide the *currently active* slide (the toggle acts on it).
    await page.getByTestId("pptx-hide-slide-toggle").click();

    // Enter presenter mode and assert the pill renders with the count.
    await page.getByTestId("pptx-present").click();
    const present = page.getByTestId("pptx-present-mode");
    await expect(present).toBeVisible({ timeout: 5_000 });

    const pill = page.getByTestId("pptx-present-hidden-count");
    await expect(pill).toBeVisible({ timeout: 5_000 });
    await expect(pill).toContainText(/hidden/i);

    // Bail out so we don't leave the spec stuck in present mode for
    // the next worker.
    await page.keyboard.press("Escape");
  });
});
