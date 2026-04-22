import { expect, test } from "@playwright/test";
import { gotoEditor } from "./_helpers";

/**
 * Smoke spec for the Word-style "Restrict editing" affordance landed
 * in this session. The Review tab now hosts a `Schützen` group whose
 * single button opens the {@link import("../app/editor/ProtectDocumentDialog").ProtectDocumentDialog}.
 * The dialog dispatches `docx:set-protection`, which lifts the
 * snapshot's `protectionState.enabled` flag — the toolbar button
 * mirrors this flag back through `aria-pressed`, so we can assert
 * the round-trip without poking into the document model.
 */
test.describe("docx restrict editing", () => {
  test("opens the dialog from the Review tab and persists the toggle", async ({ page }) => {
    await gotoEditor(page);

    // Switch to the Review tab where the new Schützen group lives.
    await page.getByTestId("ribbon-tab-review").click();
    const ribbon = page.getByTestId("docx-toolbar");
    await expect(ribbon).toHaveAttribute("data-active-tab", "review");

    const protectBtn = page.getByTestId("docx-protect-document");
    await expect(protectBtn).toBeVisible();

    // Open the dialog, pick "Read-only", enforce, and submit. The
    // dialog mounts native radios + checkboxes with stable testIds
    // so the spec doesn't depend on the surrounding label text.
    await protectBtn.click();
    await page.getByTestId("docx-protect-readonly").check();
    await page.getByTestId("docx-protect-apply").click();

    // The toolbar reflects the active protection by toggling
    // `aria-pressed=true` on the Protect button.
    await expect(protectBtn).toHaveAttribute("aria-pressed", "true", { timeout: 5_000 });
  });
});
