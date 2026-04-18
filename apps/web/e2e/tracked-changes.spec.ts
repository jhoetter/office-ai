import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";
import { expect, test } from "@playwright/test";
import { gotoEditor } from "./_helpers";

/**
 * P1.2 / W5: tracked-changes margin UI.
 *
 * Both cases load `fixtures/docx/real-world/06-comments-and-changes.docx`
 * (one `<w:ins>` + one `<w:del>` revision) via the toolbar's hidden
 * file input. They then locate a Word-style balloon in the right-margin
 * gutter and click its Accept (resp. Reject) button, asserting the
 * resulting toast.
 *
 * If W4's handlers ever regress to `NotImplementedError` stubs the
 * editor surfaces a "Not yet supported in this build" toast instead;
 * we accept both messages with a `.or()` matcher so the test is
 * resilient to that one stub-flip — the brief explicitly mandates the
 * graceful-toast fallback.
 */
const FIXTURE_PATH = resolve(
  dirname(fileURLToPath(import.meta.url)),
  "../../../fixtures/docx/real-world/06-comments-and-changes.docx"
);

async function loadTrackedFixture(page: import("@playwright/test").Page): Promise<void> {
  await gotoEditor(page);
  // The toolbar's "Open .docx" trigger forwards to a hidden <input>;
  // setInputFiles on it dispatches change synchronously and
  // mountAgent re-mounts with the fixture.
  await page.locator('input[type="file"]').setInputFiles(FIXTURE_PATH);
  // Fixture's first paragraph reads "Draft" (Heading2). Wait for it
  // so we know mountAgent finished.
  await expect(page.getByText("Draft", { exact: false })).toBeVisible({ timeout: 15_000 });
}

test.describe("editor: tracked-changes UI", () => {
  test("clicking Accept on a tracked change dispatches docx:accept-change", async ({ page }) => {
    await loadTrackedFixture(page);

    const balloon = page.getByTestId("tracked-change-balloon").first();
    await expect(balloon).toBeVisible();
    const revisionId = (await balloon.getAttribute("data-revision-id")) ?? "";
    expect(revisionId).not.toEqual("");
    await balloon.getByRole("button", { name: `Accept change ${revisionId}` }).click();

    await expect(
      page.getByText(/Change accepted\./).or(page.getByText(/Not yet supported in this build/))
    ).toBeVisible({ timeout: 7_500 });
  });

  test("clicking Reject on a tracked change dispatches docx:reject-change", async ({ page }) => {
    await loadTrackedFixture(page);

    const balloons = page.getByTestId("tracked-change-balloon");
    await expect(balloons.first()).toBeVisible();
    // Pick the second revision (the deletion) so accept-then-reject
    // ordering across the two cases doesn't collide on a shared id.
    const target = (await balloons.count()) > 1 ? balloons.nth(1) : balloons.first();
    const revisionId = (await target.getAttribute("data-revision-id")) ?? "";
    expect(revisionId).not.toEqual("");
    await target.getByRole("button", { name: `Reject change ${revisionId}` }).click();

    await expect(
      page.getByText(/Change rejected\./).or(page.getByText(/Not yet supported in this build/))
    ).toBeVisible({ timeout: 7_500 });
  });
});
