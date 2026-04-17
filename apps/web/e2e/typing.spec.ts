import { expect, test } from "@playwright/test";
import { focusEditor, gotoEditor } from "./_helpers";

test.describe("editor: typing flows through the command bus", () => {
  test("typed characters appear in the document", async ({ page }) => {
    await gotoEditor(page);
    await focusEditor(page);

    // Park the caret at the absolute start so we know exactly where the
    // text will land.
    await page.keyboard.press("ControlOrMeta+Home");
    await page.keyboard.type("HELLO ", { delay: 20 });

    await expect(page.locator(".ProseMirror").first()).toContainText("HELLO");
  });
});
