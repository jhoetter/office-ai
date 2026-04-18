import { expect, test } from "@playwright/test";
import { gotoEditor, selectParagraphContaining } from "./_helpers";

/**
 * P1.2 / W5: comments sidebar.
 *
 * Each test seeds a comment via the toolbar's "Add comment" composer
 * (P2.5/W24 replaced the old hard-coded "Looks good?" recipe with a
 * real popover) and then exercises one of the sidebar affordances.
 */
async function addComment(page: import("@playwright/test").Page, body: string): Promise<void> {
  await selectParagraphContaining(page, "Welcome");
  await page.getByTitle("Add comment").click();
  const composer = page.getByTestId("comment-composer");
  await expect(composer).toBeVisible();
  await page.getByTestId("comment-composer-text").fill(body);
  await page.getByTestId("comment-composer-submit").click();
  await expect(page.getByText(/Comment added\./)).toBeVisible({ timeout: 7_500 });
}

test.describe("editor: comments sidebar", () => {
  test("adding a comment via the toolbar surfaces it as a thread in the sidebar", async ({ page }) => {
    await gotoEditor(page);
    await addComment(page, "Looks good?");

    const sidebar = page.getByTestId("comments-sidebar");
    await expect(sidebar).toBeVisible();
    const thread = sidebar.locator("[data-testid='comment-thread']").first();
    await expect(thread).toBeVisible({ timeout: 7_500 });
    await expect(thread).toContainText("Looks good?");
    await expect(thread).toHaveAttribute("data-resolved", "false");

    // Clicking the author scrolls the editor to and flashes the comment
    // anchor; the temporary class is added then removed after ~1.6s. We
    // assert it lands at all by waiting for the highlight class.
    await thread.getByRole("button", { name: /scroll to .* comment/i }).click();
    await expect(page.locator(".ProseMirror .pm-comment-mark").first()).toBeVisible();
  });

  test("replying to a comment dispatches docx:reply-comment and renders the reply", async ({ page }) => {
    await gotoEditor(page);
    await addComment(page, "Looks good?");

    const thread = page.getByTestId("comments-sidebar").locator("[data-testid='comment-thread']").first();
    await thread.getByLabel("Reply").fill("Sounds great");
    await thread.getByRole("button", { name: "Send reply" }).click();

    await expect(page.getByText(/Reply added\./)).toBeVisible({ timeout: 7_500 });
    await expect(thread).toContainText("Sounds great");
  });

  test("resolving a comment marks the thread resolved and shows the badge", async ({ page }) => {
    await gotoEditor(page);
    await addComment(page, "Looks good?");

    const thread = page.getByTestId("comments-sidebar").locator("[data-testid='comment-thread']").first();
    await thread.getByLabel("Resolve comment").click();

    await expect(page.getByText(/Comment resolved\./)).toBeVisible({ timeout: 7_500 });
    await expect(thread).toHaveAttribute("data-resolved", "true", { timeout: 7_500 });
    await expect(thread.getByText(/Resolved/i)).toBeVisible();
  });
});
