import { expect, test } from "@playwright/test";
import { gotoEditor } from "./_helpers";

test.describe("editor: open bundled sample", () => {
  test("renders the sample document and metadata", async ({ page }) => {
    await gotoEditor(page);

    await expect(page.getByText("DOCX editor")).toBeVisible();
    await expect(page.getByText("Welcome to officeAI")).toBeVisible();
    // Metadata strip is hidden below md breakpoint; the configured
    // viewport (1280×800) keeps it visible.
    await expect(page.getByText(/paragraphs · rev \d+ · \d+ comments?/)).toBeVisible();
  });
});
