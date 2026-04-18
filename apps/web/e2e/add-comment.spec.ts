import { expect, test } from "@playwright/test";
import { gotoEditor, selectParagraphContaining } from "./_helpers";

test.describe("editor: comment lifecycle", () => {
  test("Add comment increments the comments counter", async ({ page }) => {
    await gotoEditor(page);

    const meta = page.getByText(/paragraphs · rev \d+ · (\d+) comments?/);
    await expect(meta).toBeVisible();
    const before = await meta.textContent();
    const commentsBefore = Number(before?.match(/(\d+) comments?/)?.[1] ?? "0");

    // `docx:add-comment` rejects multi-paragraph ranges, so anchor the
    // selection to the first paragraph only via a triple-click.
    await selectParagraphContaining(page, "Welcome");
    await page.getByTitle("Add comment").click();

    await expect(page.getByText(/Comment added\./)).toBeVisible({ timeout: 7_500 });
    await expect(meta).not.toHaveText(before ?? "", { timeout: 7_500 });
    const after = await meta.textContent();
    const commentsAfter = Number(after?.match(/(\d+) comments?/)?.[1] ?? "0");
    expect(commentsAfter).toBeGreaterThan(commentsBefore);
  });
});
