import { expect, test } from "@playwright/test";
import { focusEditor, gotoEditor } from "./_helpers";

test.describe("editor: pressing Enter splits paragraphs", () => {
  test("Enter at end of first paragraph adds a block", async ({ page }) => {
    await gotoEditor(page);
    await focusEditor(page);

    const meta = page.getByText(/(\d+) blocks · rev \d+ · \d+ comments/);
    await expect(meta).toBeVisible();
    const before = await meta.textContent();
    const blocksBefore = Number(before?.match(/(\d+) blocks/)?.[1] ?? "0");

    await page.keyboard.press("ControlOrMeta+End");
    await page.keyboard.press("Enter");
    await page.keyboard.type("New paragraph", { delay: 20 });

    await expect(meta).not.toHaveText(before ?? "", { timeout: 7_500 });
    const after = await meta.textContent();
    const blocksAfter = Number(after?.match(/(\d+) blocks/)?.[1] ?? "0");
    expect(blocksAfter).toBeGreaterThan(blocksBefore);
    await expect(page.locator(".ProseMirror").first()).toContainText("New paragraph");
  });
});
