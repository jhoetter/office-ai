import { readFile } from "node:fs/promises";
import { expect, test } from "@playwright/test";
import { focusEditor, gotoEditor } from "./_helpers";

test.describe("editor: export → download", () => {
  test("Export downloads a valid OOXML zip after a trivial edit", async ({ page }) => {
    await gotoEditor(page);
    await focusEditor(page);
    await page.keyboard.press("ControlOrMeta+Home");
    await page.keyboard.type("E2E ", { delay: 20 });

    const [download] = await Promise.all([
      page.waitForEvent("download"),
      page.getByRole("button", { name: /Export/i }).click(),
    ]);

    const path = await download.path();
    expect(path).toBeTruthy();
    const bytes = await readFile(path!);
    // OOXML packages are ZIPs; the local file header magic is "PK\x03\x04".
    expect(bytes.subarray(0, 4).toString("hex")).toBe("504b0304");
    expect(bytes.length).toBeGreaterThan(500);
  });
});
