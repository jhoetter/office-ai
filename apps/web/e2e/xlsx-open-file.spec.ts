import path from "node:path";
import { fileURLToPath } from "node:url";
import { expect, test } from "@playwright/test";
import { gotoXlsxEditor } from "./_helpers";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const FIXTURE = path.resolve(__dirname, "../../../fixtures/xlsx/synthetic/01-single-sheet-numbers.xlsx");

test.describe("xlsx editor: open .xlsx from disk", () => {
  test("loading a fixture replaces the seeded sample workbook", async ({ page }) => {
    await gotoXlsxEditor(page);

    // Sanity: the bundled sample is the seeded workbook.
    await expect(page.getByTestId("filename")).toHaveText("sample.xlsx");
    await expect(page.getByTestId("cell-A1")).toContainText("Name");

    // Drive the hidden <input type="file"> directly — fastest path that
    // matches what the visible "Open" button triggers via .click().
    await page.getByTestId("open-xlsx-input").setInputFiles(FIXTURE);

    // The header label flips to the uploaded file's name.
    await expect(page.getByTestId("filename")).toHaveText("01-single-sheet-numbers.xlsx", {
      timeout: 10_000,
    });

    // The fixture's A1 / B1 are "Item" / "Qty"; the previous "Name" /
    // "Score" headers from the seeded sample must be gone.
    await expect(page.getByTestId("cell-A1")).toContainText("Item");
    await expect(page.getByTestId("cell-B1")).toContainText("Qty");
    await expect(page.getByTestId("cell-A2")).toContainText("Widget");

    // Opening a fresh file resets the agent → revision rolls back to 0.
    await expect(page.getByTestId("revision-badge")).toContainText("rev 0");
  });
});
