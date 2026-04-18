import { expect, test } from "@playwright/test";
import { gotoEditor, selectParagraphContaining } from "./_helpers";

test.describe("editor: comment composer (P2.5/W24)", () => {
  test("Add comment opens the composer and submits a user-written body", async ({ page }) => {
    await gotoEditor(page);

    const meta = page.getByText(/paragraphs · rev \d+ · (\d+) comments?/);
    await expect(meta).toBeVisible();
    const before = await meta.textContent();
    const commentsBefore = Number(before?.match(/(\d+) comments?/)?.[1] ?? "0");

    // `docx:add-comment` rejects multi-paragraph ranges, so anchor the
    // selection to the first paragraph only via a triple-click.
    await selectParagraphContaining(page, "Welcome");
    await page.getByTitle("Add comment").click();

    // The composer is the popover that replaced the hard-coded
    // `text: "Looks good?"` recipe (P2.5/W24). The selection snippet
    // must be quoted, the textarea must start empty, and submission
    // must be gated behind non-empty input.
    const composer = page.getByTestId("comment-composer");
    await expect(composer).toBeVisible();
    const textarea = page.getByTestId("comment-composer-text");
    await expect(textarea).toHaveValue("");
    const submit = page.getByTestId("comment-composer-submit");
    await expect(submit).toBeDisabled();

    await textarea.fill("Tighten this opening paragraph.");
    await expect(submit).toBeEnabled();
    await submit.click();

    await expect(composer).not.toBeVisible();
    await expect(page.getByText(/Comment added\./)).toBeVisible({ timeout: 7_500 });
    await expect(meta).not.toHaveText(before ?? "", { timeout: 7_500 });
    const after = await meta.textContent();
    const commentsAfter = Number(after?.match(/(\d+) comments?/)?.[1] ?? "0");
    expect(commentsAfter).toBeGreaterThan(commentsBefore);

    // The user's exact comment body must surface in the comments
    // sidebar, not the legacy "Looks good?" placeholder.
    await expect(page.getByText("Tighten this opening paragraph.")).toBeVisible();
  });

  test("Escape dismisses the composer without adding a comment", async ({ page }) => {
    await gotoEditor(page);
    const meta = page.getByText(/paragraphs · rev \d+ · (\d+) comments?/);
    const before = await meta.textContent();

    await selectParagraphContaining(page, "Welcome");
    await page.getByTitle("Add comment").click();
    const composer = page.getByTestId("comment-composer");
    await expect(composer).toBeVisible();

    await page.keyboard.press("Escape");
    await expect(composer).not.toBeVisible();
    // Counter unchanged.
    await expect(meta).toHaveText(before ?? "");
  });
});
