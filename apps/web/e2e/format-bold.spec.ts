import { expect, test } from "@playwright/test";
import { gotoEditor, selectAll } from "./_helpers";

test.describe("editor: bold toolbar", () => {
  test("clicking Bold marks the selection as <strong>", async ({ page }) => {
    await gotoEditor(page);
    await selectAll(page);

    await page.getByTitle("Bold").click();

    // ProseMirror's default schema renders strong marks as <strong>;
    // because we selected everything, the first paragraph's "Welcome"
    // text is now wrapped.
    const strongs = page.locator(".ProseMirror strong");
    await expect(strongs.first()).toBeVisible();
    await expect(strongs.first()).toContainText("Welcome");
  });
});
