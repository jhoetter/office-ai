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

    page.on("console", (msg) => {
      const text = msg.text();
      if (text.includes("[docx-mount]")) {
        console.log("BROWSER:", text);
      }
    });

    await focusEditor(page);
    // Land at the end of the document so we don't disturb existing
    // headings (which would also exercise tracked deletions).
    await page.keyboard.press("End");
    await page.keyboard.type(" SUGGESTED", { delay: 50 });

    // Debug: dump the editor DOM so we can see what landed.
    const dump = await page.evaluate(() => {
      const pm = document.querySelector(".ProseMirror");
      return {
        text: pm?.textContent ?? "<no .ProseMirror>",
        revisionSpans: pm?.querySelectorAll(".pm-revision-ins").length ?? 0,
        delSpans: pm?.querySelectorAll(".pm-revision-del").length ?? 0,
        innerHTML: pm?.innerHTML?.slice(0, 1500) ?? "",
      };
    });
    console.log("DEBUG editor state:", JSON.stringify(dump, null, 2));

    // The renderer should re-project from the snapshot, surfacing a
    // .pm-revision-ins span carrying the freshly typed text. We
    // assert against the underline-styled span rather than the raw
    // page text so a regression that drops the mark fails this test.
    const insertionSpan = page.locator(".pm-revision-ins").last();
    await expect(insertionSpan).toBeVisible({ timeout: 5_000 });
    await expect(insertionSpan).toContainText("SUGGESTED");

    // The Word-style margin balloon should land in the right gutter
    // with the German "hat eingefügt" label and an Accept button.
    const balloon = page.getByTestId("tracked-change-balloon").last();
    await expect(balloon).toBeVisible();
    await expect(balloon).toContainText("hat eingefügt");
    await expect(balloon.getByRole("button", { name: /^Accept change/ })).toBeVisible();
  });
});
