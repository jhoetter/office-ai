import { expect, test } from "@playwright/test";
import { gotoEditor, focusEditor } from "./_helpers";

/**
 * Regression coverage for the "Suggesting mode looks like Editing mode" bug.
 *
 * Before the fix in `packages/docx/src/renderer/mount.ts`, typing while
 * the toolbar's edit-mode selector was set to "Suggesting" produced a
 * `<w:ins>` revision wrapper in the snapshot but PM's local view still
 * reflected the optimistic apply (plain text). The subscribe handler
 * was skipped via `pendingFunnelCount` so the `revision_mark` never
 * landed in the DOM — visually indistinguishable from Editing mode.
 *
 * The fix bypasses the suppression counter for tracked-revision
 * commands so the snapshot is re-projected into PM, the
 * `.pm-revision-ins` span shows up, the `TrackedChangesMargin` paints
 * a balloon, and the toolbar status reports the new revision.
 */
test.describe("editor: suggesting mode", () => {
  test("typing in Suggesting mode renders a tracked insertion", async ({ page }) => {
    await gotoEditor(page);

    // Switch the edit-mode selector from "Editing" to "Suggesting".
    // The button label includes the active mode so we target the
    // current value before clicking.
    await page.getByRole("button", { name: /Edit mode: Editing/ }).click();
    await page.getByRole("menuitemradio", { name: /^Suggesting/ }).click();
    await expect(page.getByRole("button", { name: /Edit mode: Suggesting/ })).toBeVisible();

    await focusEditor(page);
    // Land at the end of the line so we don't disturb existing
    // headings (which would also exercise tracked deletions).
    await page.keyboard.press("End");
    await page.keyboard.type(" SUGGESTED", { delay: 50 });

    // Each keystroke produces its own `<w:ins>` revision wrapper —
    // that matches Word's behaviour ("one revision per insertion
    // event") and keeps the snapshot round-trippable. Assert that
    // the concatenated text of every .pm-revision-ins span equals
    // what the user typed so a regression that drops the marks
    // (the original "Suggesting mode looks like Editing mode" bug)
    // or reverses character order (the
    // positionFromPM-pinned-to-run-0 regression) fails this test.
    await expect(page.locator(".pm-revision-ins").first()).toBeVisible({
      timeout: 5_000,
    });
    const insertedText = await page.evaluate(() => {
      const spans = Array.from(document.querySelectorAll(".pm-revision-ins"));
      return spans.map((s) => s.textContent ?? "").join("");
    });
    expect(insertedText).toBe(" SUGGESTED");

    // The Word-style margin balloon should land in the right gutter
    // with the German "hat eingefügt" label and an Accept button.
    const balloon = page.getByTestId("tracked-change-balloon").last();
    await expect(balloon).toBeVisible();
    await expect(balloon).toContainText("hat eingefügt");
    await expect(balloon.getByRole("button", { name: /^Accept change/ })).toBeVisible();
  });
});
