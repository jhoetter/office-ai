import { expect, test, type Page } from "@playwright/test";
import { gotoEditor, gotoXlsxEditor } from "./_helpers";

/**
 * Rich Export — dialog smoke per product.
 *
 * The dropdown routes formats with options (PDF, PNG zip, …) through
 * the rich Export dialog. The dialog also opens explicitly via the
 * trailing "More options…" item. These specs check the full path
 * including the network call to `/api/convert`, which we mock so the
 * test doesn't require LibreOffice on the runner.
 */

async function gotoPptxEditor(page: Page): Promise<void> {
  await page.goto("/pptx-editor", { timeout: 60_000 });
  await expect(page.getByTestId("pptx-sidebar")).toBeVisible({ timeout: 60_000 });
  await expect(page.getByTestId("pptx-sidebar").getByRole("button").first()).toBeVisible({
    timeout: 60_000,
  });
}

async function mockConvertApi(page: Page, body: Buffer = Buffer.from("%PDF-1.4 stub")): Promise<void> {
  await page.route("**/api/convert", async (route) => {
    await route.fulfill({
      status: 200,
      headers: {
        "Content-Type": "application/pdf",
        "Content-Disposition": 'attachment; filename="stub.pdf"',
      },
      body,
    });
  });
}

test.describe("export dialog — DOCX", () => {
  test.setTimeout(60_000);

  test("clicking PDF in the dropdown opens the dialog with PDF preselected", async ({ page }) => {
    await gotoEditor(page);
    await page.getByTestId("shell-export").click();
    await page.getByTestId("shell-export-pdf").click();

    const dialog = page.getByTestId("shell-export-dialog");
    await expect(dialog).toBeVisible();
    // Filename preview reflects the selected format's extension.
    await expect(page.getByTestId("shell-export-filename-preview")).toContainText(".pdf");
    // PDF-specific options are rendered.
    await expect(page.getByTestId("shell-export-option-pageSize")).toBeVisible();
    await expect(page.getByTestId("shell-export-option-embedFonts")).toBeVisible();
  });

  test("'More options…' opens the dialog without a forced selection", async ({ page }) => {
    await gotoEditor(page);
    await page.getByTestId("shell-export").click();
    await page.getByTestId("shell-export-open-dialog").click();

    await expect(page.getByTestId("shell-export-dialog")).toBeVisible();
    // Default selection is the first format (DOCX) — no options pane.
    await expect(page.getByTestId("shell-export-filename-preview")).toContainText(".docx");
  });

  test("Cancel closes the dialog without exporting", async ({ page }) => {
    await gotoEditor(page);
    await page.getByTestId("shell-export").click();
    await page.getByTestId("shell-export-pdf").click();

    await page.getByTestId("shell-export-cancel").click();
    await expect(page.getByTestId("shell-export-dialog")).toHaveCount(0);
  });

  test("PDF export through the dialog hits /api/convert and downloads", async ({ page }) => {
    await mockConvertApi(page, Buffer.from("%PDF-1.4 docx-mock"));
    await gotoEditor(page);

    await page.getByTestId("shell-export").click();
    await page.getByTestId("shell-export-pdf").click();

    const downloadPromise = page.waitForEvent("download");
    await page.getByTestId("shell-export-confirm").click();
    const download = await downloadPromise;
    expect(download.suggestedFilename()).toMatch(/\.pdf$/);
    await expect(page.getByTestId("shell-export-dialog")).toHaveCount(0);
  });
});

test.describe("export dialog — XLSX", () => {
  test.setTimeout(60_000);

  test("PDF dialog renders orientation + fit-to-width controls", async ({ page }) => {
    await gotoXlsxEditor(page);

    await page.getByTestId("shell-export").click();
    await page.getByTestId("shell-export-pdf").click();

    await expect(page.getByTestId("shell-export-dialog")).toBeVisible();
    await expect(page.getByTestId("shell-export-option-orientation")).toBeVisible();
    await expect(page.getByTestId("shell-export-option-fitToWidth")).toBeVisible();
  });
});

test.describe("export dialog — PPTX", () => {
  test.setTimeout(60_000);

  test("PNG zip dialog exposes resolution + slide range fields", async ({ page }) => {
    await gotoPptxEditor(page);

    await page.getByTestId("shell-export").click();
    await page.getByTestId("shell-export-png-zip").click();

    await expect(page.getByTestId("shell-export-dialog")).toBeVisible();
    await expect(page.getByTestId("shell-export-option-scale")).toBeVisible();
    await expect(page.getByTestId("shell-export-option-slideRange")).toBeVisible();
  });
});
