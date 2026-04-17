import { expect, test } from "@playwright/test";
import { gotoEditor, selectAll } from "./_helpers";

/**
 * P1.2 / W5: toolbar parity coverage.
 *
 * The four cases below exercise the four toolbar groups added in this
 * batch — font size, color, alignment (graceful "not yet supported"),
 * and list toggle (graceful "not yet supported"). Cases that currently
 * resolve through `docx:format-range` actually mutate the document;
 * cases that have no backing command assert the toast instead.
 */
test.describe("editor: toolbar parity", () => {
  test("font size picker applies a font_size mark to the selection", async ({ page }) => {
    await gotoEditor(page);
    await selectAll(page);

    await page.getByLabel("Font size").selectOption({ value: "24" });

    // The format-range handler runs in-process; the bus re-projects PM
    // synchronously on subscribe so the mark is visible immediately.
    await expect(page.locator(".ProseMirror .pm-font-size").first()).toBeVisible({
      timeout: 5_000,
    });
  });

  test("color picker applies a color mark to the selection", async ({ page }) => {
    await gotoEditor(page);
    await selectAll(page);

    // The Font color button opens a popover; the swatches inside are
    // labelled `Font color: <name>` so we can pick one without coords.
    await page.getByRole("button", { name: "Font color" }).click();
    await page.getByRole("menuitem", { name: "Font color: Red" }).click();

    await expect(page.locator(".ProseMirror .pm-color").first()).toBeVisible({ timeout: 5_000 });
    // The mark renders as `style="color: #C00000"` (see schema.ts).
    const styleAttr = await page.locator(".ProseMirror .pm-color").first().getAttribute("style");
    expect(styleAttr ?? "").toMatch(/color:\s*#C00000/i);
  });

  test("alignment buttons surface a 'not yet supported' toast", async ({ page }) => {
    await gotoEditor(page);

    await page.getByLabel("Align center").click();

    // The toast wording comes from DocxEditor.surfaceUnsupported and is
    // intentionally distinct from the generic agent-funnel "deferred"
    // toast emitted by `onUnsupported(events)`.
    await expect(page.getByText(/alignment is not yet supported/i)).toBeVisible({
      timeout: 5_000,
    });
  });

  test("list toggles surface a 'not yet supported' toast", async ({ page }) => {
    await gotoEditor(page);

    await page.getByLabel("Bullet list").click();
    await expect(page.getByText(/bullet list is not yet supported/i)).toBeVisible({
      timeout: 5_000,
    });

    await page.getByLabel("Numbered list").click();
    await expect(page.getByText(/numbered list is not yet supported/i)).toBeVisible({
      timeout: 5_000,
    });
  });
});
