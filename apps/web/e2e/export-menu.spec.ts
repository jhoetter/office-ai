import { expect, test, type Page } from "@playwright/test";
import { gotoEditor, gotoXlsxEditor } from "./_helpers";

/**
 * Rich Export — dropdown smoke per product.
 *
 * Each editor exposes the same `shell-export` button. Clicking it
 * always opens a categorized dropdown populated from the product
 * adapter's `exportFormats` — DOCX/XLSX use the fine-grained
 * Native / PDF & web / Data / Images groups, PPTX collapses the
 * deck-level entries into a single "Whole deck" group beneath
 * "This slide".
 * "Instant" formats download immediately; formats with options route
 * through the rich Export dialog, and a trailing "More options…" item
 * always opens the dialog explicitly.
 */

async function gotoPptxEditor(page: Page): Promise<void> {
  await page.goto("/pptx-editor", { timeout: 60_000 });
  await expect(page.getByTestId("pptx-sidebar")).toBeVisible({ timeout: 60_000 });
  await expect(page.getByTestId("pptx-sidebar").getByRole("button").first()).toBeVisible({
    timeout: 60_000,
  });
}

async function openExportDropdown(page: Page): Promise<void> {
  await page.getByTestId("shell-export").click();
  await expect(page.getByTestId("shell-export-open-dialog")).toBeVisible();
}

test.describe("export dropdown — DOCX", () => {
  test.setTimeout(60_000);

  test("lists native + pdf-web + data formats", async ({ page }) => {
    await gotoEditor(page);
    await openExportDropdown(page);

    await expect(page.getByTestId("shell-export-docx")).toBeVisible();
    await expect(page.getByTestId("shell-export-pdf")).toBeVisible();
    await expect(page.getByTestId("shell-export-html")).toBeVisible();
    await expect(page.getByTestId("shell-export-txt")).toBeVisible();
    await expect(page.getByTestId("shell-export-md")).toBeVisible();
  });

  test("instant TXT export triggers a download", async ({ page }) => {
    await gotoEditor(page);
    await openExportDropdown(page);

    const downloadPromise = page.waitForEvent("download");
    await page.getByTestId("shell-export-txt").click();
    const download = await downloadPromise;
    expect(download.suggestedFilename()).toMatch(/\.txt$/);
  });

  test("instant Markdown export triggers a download", async ({ page }) => {
    await gotoEditor(page);
    await openExportDropdown(page);

    const downloadPromise = page.waitForEvent("download");
    await page.getByTestId("shell-export-md").click();
    const download = await downloadPromise;
    expect(download.suggestedFilename()).toMatch(/\.md$/);
  });
});

test.describe("export dropdown — XLSX", () => {
  test.setTimeout(60_000);

  test("lists data formats including all-sheet zip", async ({ page }) => {
    await gotoXlsxEditor(page);
    await openExportDropdown(page);

    await expect(page.getByTestId("shell-export-xlsx")).toBeVisible();
    await expect(page.getByTestId("shell-export-csv")).toBeVisible();
    await expect(page.getByTestId("shell-export-csv-all")).toBeVisible();
    await expect(page.getByTestId("shell-export-tsv")).toBeVisible();
    await expect(page.getByTestId("shell-export-json")).toBeVisible();
  });

  test("instant CSV export triggers a download", async ({ page }) => {
    await gotoXlsxEditor(page);
    await openExportDropdown(page);

    const downloadPromise = page.waitForEvent("download");
    await page.getByTestId("shell-export-csv").click();
    const download = await downloadPromise;
    expect(download.suggestedFilename()).toMatch(/\.csv$/);
  });
});

test.describe("export dropdown — PPTX", () => {
  test.setTimeout(60_000);

  test("lists current-slide + native + pdf-web + image bundles", async ({ page }) => {
    await gotoPptxEditor(page);
    await openExportDropdown(page);

    // The "This slide" group sits at the top with single-slide
    // exports for the slide that's currently in view.
    await expect(page.getByTestId("shell-export-slide-png")).toBeVisible();
    await expect(page.getByTestId("shell-export-slide-jpeg")).toBeVisible();
    await expect(page.getByTestId("shell-export-slide-svg")).toBeVisible();
    await expect(page.getByTestId("shell-export-slide-pdf")).toBeVisible();

    await expect(page.getByTestId("shell-export-pptx")).toBeVisible();
    await expect(page.getByTestId("shell-export-pdf")).toBeVisible();
    await expect(page.getByTestId("shell-export-svg-zip")).toBeVisible();
    await expect(page.getByTestId("shell-export-png-zip")).toBeVisible();
  });

  test("instant SVG zip export triggers a download", async ({ page }) => {
    await gotoPptxEditor(page);
    await openExportDropdown(page);

    const downloadPromise = page.waitForEvent("download");
    await page.getByTestId("shell-export-svg-zip").click();
    const download = await downloadPromise;
    expect(download.suggestedFilename()).toMatch(/\.zip$/);
  });

  test("instant 'current slide as SVG' triggers a single-file SVG download", async ({ page }) => {
    await gotoPptxEditor(page);
    await openExportDropdown(page);

    const downloadPromise = page.waitForEvent("download");
    await page.getByTestId("shell-export-slide-svg").click();
    const download = await downloadPromise;
    // Filename includes a `-slide{N}` suffix so single-slide exports
    // don't clobber a previously downloaded deck-level export.
    expect(download.suggestedFilename()).toMatch(/-slide\d+\.svg$/);
  });
});
