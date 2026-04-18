import { expect, test } from "@playwright/test";
import { focusEditor, gotoEditor } from "./_helpers";

/**
 * Pilcrow ("Show formatting marks") toggle. Asserted via the
 * widget-decoration DOM the {@link formattingMarksPlugin} produces:
 * a `.fmt-mark-pilcrow` span at the end of every paragraph when the
 * toggle is on, none when it's off.
 *
 * The keyboard binding (`Mod+Shift+8`) and the toolbar button drive
 * the same handler; both pathways are covered here.
 */
test.describe("formatting marks (pilcrow)", () => {
  // Allow generous head room for `next dev`'s first-hit compile when
  // these specs run after others in the same Playwright invocation.
  test.setTimeout(60_000);

  test("toolbar button toggles ¶ widgets on and off", async ({ page }) => {
    await gotoEditor(page);

    const pilcrows = page.locator(".ProseMirror .fmt-mark-pilcrow");
    await expect(pilcrows).toHaveCount(0);

    await page.getByTitle("Show formatting marks").click();
    await expect(pilcrows.first()).toBeVisible({ timeout: 5_000 });
    expect(await pilcrows.count()).toBeGreaterThan(0);

    await page.getByTitle("Show formatting marks").click();
    await expect(pilcrows).toHaveCount(0);
  });

  test("Mod+Shift+8 toggles the same state", async ({ page }) => {
    await gotoEditor(page);
    await focusEditor(page);

    const pilcrows = page.locator(".ProseMirror .fmt-mark-pilcrow");
    await expect(pilcrows).toHaveCount(0);

    await page.keyboard.press("ControlOrMeta+Shift+8");
    await expect(pilcrows.first()).toBeVisible({ timeout: 5_000 });

    await page.keyboard.press("ControlOrMeta+Shift+8");
    await expect(pilcrows).toHaveCount(0);
  });
});
