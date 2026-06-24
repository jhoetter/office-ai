import path from "node:path";
import { expect, test } from "@playwright/test";
import { gotoXlsxEditor } from "./_helpers";
import { fixturePath, requiredMatrixFixture } from "../../../tests/fixture-matrix.js";

// Use the shared matrix entry so this smoke follows the same canonical
// fixture selection as roundtrip and MCP tests.
const FIXTURE_ENTRY = requiredMatrixFixture("xlsx", { id: "xlsx.synthetic.single-sheet-numbers" });
const FIXTURE = fixturePath(FIXTURE_ENTRY);
const FIXTURE_NAME = path.basename(FIXTURE_ENTRY.path);

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
    await expect(page.getByTestId("filename")).toHaveText(FIXTURE_NAME, {
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
