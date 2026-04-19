import { expect, test } from "@playwright/test";
import { focusEditor, gotoEditor, selectAll, selectParagraphContaining } from "./_helpers";

/**
 * Word-parity keyboard shortcut smoke tests. These cover the
 * `wordShortcutsKeymapPlugin` end-to-end: a real keystroke goes
 * through ProseMirror, the funnel dispatches a command into the
 * bus, and the resulting DOM mutation is observable.
 *
 * We intentionally test the "muscle-memory" subset and rely on the
 * unit-tested handlers in `@officeai/docx` for correctness in
 * detail. The Playwright specs only need to prove that the
 * keystroke is wired up and lands in the right command.
 */
test.describe("editor: Word keyboard shortcuts", () => {
  test("Mod+B on a selection wraps it in <strong> (bold)", async ({ page }) => {
    await gotoEditor(page);
    await selectAll(page);

    await page.keyboard.press("ControlOrMeta+b");

    // The PM `bold` mark serialises to <strong> in the renderer.
    const strongs = page.locator(".ProseMirror strong");
    await expect(strongs.first()).toBeVisible();
  });

  /**
   * End-to-end regression for the user-reported "Cmd+B does nothing"
   * symptom (post-undo/redo unification). The lighter test above
   * proves the chord wraps a selection at all; this one pins the
   * full muscle-memory contract:
   *
   *   1. After the splash detaches, the PM surface accepts focus.
   *   2. Selecting a paragraph + Cmd+B immediately marks it bold AND
   *      flips the toolbar Bold button to `aria-pressed=true` (the
   *      from-bus projection mustn't clobber selection or pressed
   *      state).
   *   3. A SECOND Cmd+B toggles bold off — that's where the previous
   *      bug lived (wrong `(paragraph, offset)` probe re-applied
   *      `bold:true` instead of clearing). If the toolbar `aria-
   *      pressed` and the DOM `<strong>` count agree on the off
   *      state, the round-trip is healthy.
   */
  test("Cmd+B applies AND a second press toggles it back off", async ({ page }) => {
    await gotoEditor(page);
    // Splash uses `data-testid="loading-screen"`; wait for it to be
    // gone before we rely on focus reaching the PM surface. The
    // `gotoEditor` helper already waits on the sample text, but
    // the splash can still be in its fade-out frame so we belt-
    // and-braces it.
    await page.waitForSelector('[data-testid="loading-screen"]', { state: "detached" });
    await selectParagraphContaining(page, "Welcome to officeAI");

    const boldBtn = page.getByTitle("Bold");
    const strongs = page.locator(".ProseMirror strong");

    // Apply.
    await page.keyboard.press("ControlOrMeta+b");
    await expect(strongs.first()).toBeVisible();
    await expect(boldBtn).toHaveAttribute("aria-pressed", "true");

    // Re-select the same paragraph (the from-bus projection clamps
    // selection to the previous from/to, but a triple-click is the
    // most faithful "user repeats the gesture" simulation).
    await selectParagraphContaining(page, "Welcome to officeAI");

    // Toggle off.
    await page.keyboard.press("ControlOrMeta+b");
    await expect(boldBtn).toHaveAttribute("aria-pressed", "false");
    // Bold has been cleared — the title paragraph should now have
    // zero <strong> wrappers. (Other paragraphs in the sample doc
    // don't carry bold, so this is a global zero-count assertion.)
    await expect(strongs).toHaveCount(0);
  });

  test("Mod+I on a selection wraps it in <em> (italic)", async ({ page }) => {
    await gotoEditor(page);
    await selectAll(page);

    await page.keyboard.press("ControlOrMeta+i");

    const ems = page.locator(".ProseMirror em");
    await expect(ems.first()).toBeVisible();
  });

  test("Mod+U on a selection wraps it in <u> (underline)", async ({ page }) => {
    await gotoEditor(page);
    await selectAll(page);

    await page.keyboard.press("ControlOrMeta+u");

    // The PM `underline` mark serialises to <u> in the renderer.
    const underlines = page.locator(".ProseMirror u");
    await expect(underlines.first()).toBeVisible();
  });

  test("Mod+Shift+X on a selection wraps it in <s> (strikethrough)", async ({ page }) => {
    await gotoEditor(page);
    await selectAll(page);

    await page.keyboard.press("ControlOrMeta+Shift+x");

    const strikes = page.locator(".ProseMirror s");
    await expect(strikes.first()).toBeVisible();
  });

  test("Mod+Shift+E centres the paragraph containing the caret", async ({ page }) => {
    await gotoEditor(page);
    await selectParagraphContaining(page, "Welcome to officeAI");

    await page.keyboard.press("ControlOrMeta+Shift+e");

    // The Align-center toolbar button reflects the active alignment
    // via aria-pressed; this is the most stable signal regardless
    // of how the renderer materialises text-align.
    const centerBtn = page.getByTitle("Align center");
    await expect(centerBtn).toHaveAttribute("aria-pressed", "true");
  });

  test("Mod+Shift+L re-aligns left after a centre toggle", async ({ page }) => {
    await gotoEditor(page);
    await selectParagraphContaining(page, "Welcome to officeAI");
    await page.keyboard.press("ControlOrMeta+Shift+e");
    await expect(page.getByTitle("Align center")).toHaveAttribute("aria-pressed", "true");

    await page.keyboard.press("ControlOrMeta+Shift+l");
    await expect(page.getByTitle("Align left")).toHaveAttribute("aria-pressed", "true");
  });

  /**
   * The bare `Cmd+R` chord must NOT trigger right-alignment — that
   * was the pre-Google-Docs-parity bug where the Word-style binding
   * stole the browser's reload shortcut. The aligned paragraph
   * stays as it was; we assert via the toolbar's `aria-pressed`
   * which mirrors `currentParagraphAlignment`.
   *
   * We can't actually let `Cmd+R` reload the page mid-test (that
   * tears down Playwright's state), so we dispatch the keystroke
   * via PM's `handleKeyDown` slot directly to verify the binding
   * is gone.
   */
  test("Mod+R alone does NOT right-align (browser keeps the reload chord)", async ({ page }) => {
    await gotoEditor(page);
    await selectParagraphContaining(page, "Welcome to officeAI");

    // Sanity: paragraph starts left-aligned (the default).
    await expect(page.getByTitle("Align right")).toHaveAttribute("aria-pressed", "false");

    // Dispatch directly into PM so we don't actually reload the
    // page. If the bare Mod+R binding were still in place, this
    // would flip the toolbar to `aria-pressed=true`.
    await page.evaluate(() => {
      const editor = document.querySelector(".ProseMirror");
      if (!editor) return;
      const ev = new KeyboardEvent("keydown", {
        key: "r",
        code: "KeyR",
        metaKey: true,
        ctrlKey: true,
        bubbles: true,
        cancelable: true,
      });
      editor.dispatchEvent(ev);
    });

    // Give the bus a microtask to settle, then assert no change.
    await page.waitForTimeout(50);
    await expect(page.getByTitle("Align right")).toHaveAttribute("aria-pressed", "false");
  });

  test("Mod+Shift+R right-aligns the paragraph (the new chord)", async ({ page }) => {
    await gotoEditor(page);
    await selectParagraphContaining(page, "Welcome to officeAI");

    await page.keyboard.press("ControlOrMeta+Shift+r");

    await expect(page.getByTitle("Align right")).toHaveAttribute("aria-pressed", "true");
  });

  test("Mod+M increases the indent readout", async ({ page }) => {
    await gotoEditor(page);
    await selectParagraphContaining(page, "Welcome to officeAI");

    await page.keyboard.press("ControlOrMeta+m");

    // The toolbar shows the left indent in inches (¼" per step).
    const indentReadout = page.getByTitle("Left indent");
    await expect(indentReadout).toContainText('0.25"');
  });

  test("Mod+Shift+M decreases the indent back to zero", async ({ page }) => {
    await gotoEditor(page);
    await selectParagraphContaining(page, "Welcome to officeAI");
    await page.keyboard.press("ControlOrMeta+m");
    await expect(page.getByTitle("Left indent")).toContainText('0.25"');

    await page.keyboard.press("ControlOrMeta+Shift+m");

    // At zero indent the readout disappears entirely (the toolbar
    // only renders it when activeIndentLeft > 0).
    await expect(page.getByTitle("Left indent")).toHaveCount(0);
  });

  test("Mod+Enter inserts a page break (delegates to existing page-keymap)", async ({ page }) => {
    await gotoEditor(page);
    await focusEditor(page);

    // Sample doc starts as a single page; the page-decorations
    // plugin re-chunks on every snapshot revision, so a hard break
    // immediately produces a second `[data-page-number="2"]`
    // wrapper.
    await expect(page.locator('[data-page-number="2"]')).toHaveCount(0);
    await page.keyboard.press("ControlOrMeta+Enter");
    await expect(page.locator('[data-page-number="2"]').first()).toBeVisible({ timeout: 5_000 });
  });

  /**
   * Toolbar Undo/Redo and the keyboard chord MUST share the same
   * history. Pre-fix the DOCX editor had two stacks (PM-history
   * for the keyboard, the bus for the toolbar), so it was
   * possible for `Cmd+Z` and the toolbar Undo button to undo
   * different things — or for the toolbar to be disabled while
   * `Cmd+Z` still worked. After unification the buttons reflect
   * `agent.canUndo() / canRedo()` and the keyboard chord routes
   * through `agent.undo()`. This test pins that contract.
   */
  test("toolbar Undo and Cmd+Z agree after a typing/undo/redo cycle", async ({ page }) => {
    await gotoEditor(page);
    await focusEditor(page);

    const undoButton = page.getByTestId("shell-undo");
    const redoButton = page.getByTestId("shell-redo");

    // Fresh sample doc — nothing to undo / redo yet.
    await expect(undoButton).toBeDisabled();
    await expect(redoButton).toBeDisabled();

    // Type a unique marker we can search for. The funnel converts
    // the IME-style insert into a `docx:insert-text` mutation on
    // the bus, so the toolbar `canUndo` flips to true.
    const marker = "ZZUNDOTEST";
    await page.keyboard.type(marker);
    await expect(page.locator(".ProseMirror")).toContainText(marker);
    await expect(undoButton).toBeEnabled();

    // Keyboard undo — the toolbar enabled state should flip to
    // match (canRedo true, canUndo back to whatever the previous
    // state was for the sample doc).
    await page.keyboard.press("ControlOrMeta+z");
    await expect(page.locator(".ProseMirror")).not.toContainText(marker);
    await expect(redoButton).toBeEnabled();

    // Toolbar redo — should re-apply through the same bus stack
    // and bring the marker back. If the keyboard had wired into a
    // separate PM history stack (the pre-fix world), clicking the
    // toolbar Redo here would do nothing.
    await redoButton.click();
    await expect(page.locator(".ProseMirror")).toContainText(marker);

    // And keyboard redo chord (Cmd+Shift+Z) should now be a no-op
    // because there's nothing on the redo stack.
    await expect(redoButton).toBeDisabled();
  });
});
