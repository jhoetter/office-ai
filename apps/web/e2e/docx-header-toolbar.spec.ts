import { resolve } from "node:path";
import { expect, test } from "@playwright/test";
import { gotoEditor } from "./_helpers";

/**
 * Word-style header/footer authoring (W4 plan: hf-toolbar + hf-tests).
 *
 * Loads `02-report-headers-footers.docx` (which actually carries
 * `word/header1.xml` + `word/footer1.xml` parts), focuses the header
 * zone, and asserts the contextual toolbar surfaces the "Insert page #"
 * + "Insert page count" + "Insert image" buttons. Then clicks
 * "Insert page #" and verifies a `.pm-hf-field` token lands inside the
 * focused header zone — proving the typed `PageNumberFieldLeaf` made
 * it from command bus → snapshot → re-rendered DOM.
 */
const FIXTURE_PATH = resolve(process.cwd(), "../../fixtures/docx/real-world/02-report-headers-footers.docx");

async function loadHeaderFixture(page: import("@playwright/test").Page): Promise<void> {
  await gotoEditor(page);
  await page.locator('input[type="file"][accept*=".docx"]').setInputFiles(FIXTURE_PATH);
  // Wait for the loaded fixture: the header band must be visible and
  // its zone must be authoring-ready (contenteditable=true). Without
  // a header part, the zone stays read-only — so this gate also
  // proves the H/F parts were parsed.
  const headerZoneContent = page
    .locator(".pm-page-header-band .pm-page-zone-header .pm-page-zone-content")
    .first();
  await expect(headerZoneContent).toBeVisible({ timeout: 15_000 });
  await expect(headerZoneContent).toHaveAttribute("contenteditable", "true", {
    timeout: 15_000,
  });
}

test.describe("docx header/footer in-place authoring toolbar", () => {
  test("focuses header → exposes Insert page # / count / image buttons", async ({ page }) => {
    await loadHeaderFixture(page);

    const headerZoneContent = page
      .locator(".pm-page-header-band .pm-page-zone-header .pm-page-zone-content")
      .first();
    await headerZoneContent.click();

    // The H/F controls now live inside the contextual "Kopf- und
    // Fußzeile" ribbon tab. The Ribbon auto-activates that tab on
    // header-zone focus, so the buttons should be present in the
    // active ribbon surface.
    const ribbon = page.getByTestId("docx-toolbar");
    await expect(ribbon).toHaveAttribute("data-active-tab", "hf-tools", { timeout: 5_000 });
    await expect(page.getByTestId("docx-hf-insert-page-number")).toBeVisible();
    await expect(page.getByTestId("docx-hf-insert-page-count")).toBeVisible();
    await expect(page.getByTestId("docx-hf-insert-image")).toBeVisible();
  });

  test("clicking Insert page # mints a PAGE field token in the header", async ({ page }) => {
    await loadHeaderFixture(page);

    const headerZoneContent = page
      .locator(".pm-page-header-band .pm-page-zone-header .pm-page-zone-content")
      .first();
    await headerZoneContent.click();

    // Move caret to the very end of the header so the field doesn't
    // split an existing run mid-word and break the readability of the
    // assertion below.
    await page.keyboard.press("ControlOrMeta+End");

    await page.getByTestId("docx-hf-insert-page-number").click();

    // The page-decorations rich-render path stamps PAGE fields as
    // `.pm-hf-field` tokens with `[PAGE]` placeholder text inside
    // the contenteditable host.
    const fieldToken = headerZoneContent.locator(".pm-hf-field").first();
    await expect(fieldToken).toBeVisible({ timeout: 5_000 });
    await expect(fieldToken).toHaveText(/PAGE/);
  });
});
