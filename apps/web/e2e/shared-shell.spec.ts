import { expect, test, type Page } from "@playwright/test";
import { gotoEditor, gotoXlsxEditor } from "./_helpers";

/**
 * F5 — End-to-end coverage of the cross-product unified shell.
 *
 * The {@link EditorShell} is mounted by all three editors (DOCX, XLSX,
 * PPTX) and ships the same primitives: command palette (`Mod+K`),
 * find/replace (`Mod+F`), shortcuts dialog (`Mod+/`), and the
 * `Open / Save / Export` triplet in the top bar. These specs prove
 * the shell behaves identically on every route — a regression on
 * any one of them flips the test for the offending product only,
 * making it obvious where the divergence happened.
 */

async function gotoPptxEditor(page: Page): Promise<void> {
  await page.goto("/pptx-editor", { timeout: 60_000 });
  await expect(page.getByTestId("pptx-sidebar")).toBeVisible({ timeout: 60_000 });
  await expect(page.getByTestId("pptx-sidebar").getByRole("button").first()).toBeVisible({
    timeout: 60_000,
  });
}

const palette = (page: Page) => page.getByTestId("command-palette");
const paletteInput = (page: Page) => page.getByTestId("command-palette-input");

test.describe("shared shell — command palette", () => {
  test.setTimeout(60_000);

  test("DOCX: Mod+K opens the palette and Esc closes", async ({ page }) => {
    await gotoEditor(page);

    await expect(palette(page)).toHaveCount(0);
    await page.keyboard.press("ControlOrMeta+k");
    await expect(palette(page)).toBeVisible();
    await expect(paletteInput(page)).toBeFocused();

    await page.keyboard.press("Escape");
    await expect(palette(page)).toHaveCount(0);
  });

  test("XLSX: Mod+K opens the palette and Esc closes", async ({ page }) => {
    await gotoXlsxEditor(page);

    await expect(palette(page)).toHaveCount(0);
    await page.keyboard.press("ControlOrMeta+k");
    await expect(palette(page)).toBeVisible();
    await expect(paletteInput(page)).toBeFocused();

    await page.keyboard.press("Escape");
    await expect(palette(page)).toHaveCount(0);
  });

  test("PPTX: Mod+K opens the palette and Esc closes", async ({ page }) => {
    await gotoPptxEditor(page);

    await expect(palette(page)).toHaveCount(0);
    await page.keyboard.press("ControlOrMeta+k");
    await expect(palette(page)).toBeVisible();
    await expect(paletteInput(page)).toBeFocused();

    await page.keyboard.press("Escape");
    await expect(palette(page)).toHaveCount(0);
  });
});

test.describe("shared shell — top bar", () => {
  test.setTimeout(60_000);

  for (const route of [
    { name: "DOCX", goto: gotoEditor },
    { name: "XLSX", goto: gotoXlsxEditor },
    { name: "PPTX", goto: gotoPptxEditor },
  ]) {
    test(`${route.name}: Open / Save / Export buttons are present in the top bar`, async ({ page }) => {
      await route.goto(page);

      // Icon-only buttons exposed via stable testids on the shared shell.
      await expect(page.getByTestId("shell-open")).toBeVisible();
      await expect(page.getByTestId("shell-save")).toBeVisible();
      await expect(page.getByTestId("shell-export")).toBeVisible();
    });
  }
});
