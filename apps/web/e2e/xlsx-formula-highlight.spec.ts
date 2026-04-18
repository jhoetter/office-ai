import { expect, test } from "@playwright/test";
import { gotoXlsxEditor } from "./_helpers";

test.describe("xlsx editor: formula reference highlighting (P12a-c)", () => {
  test("colors distinct refs in the formula bar overlay", async ({ page }) => {
    await gotoXlsxEditor(page);

    // Type-to-edit on A1 → formula bar takes focus with `=`.
    await page.getByTestId("cell-A1").click();
    await page.getByTestId("formula-input").click();
    await page.getByTestId("formula-input").fill("=B1+C2+B1");

    // Two unique refs (B1 and C2) → two distinct token colours in
    // the overlay; B1 appears twice but shares a colour.
    const overlay = page.getByTestId("formula-highlight");
    const refSpans = overlay.locator('[data-token-kind="ref"]');
    await expect(refSpans).toHaveCount(3);

    const colors = await refSpans.evaluateAll((els) =>
      els.map((el) => (el as HTMLElement).style.color)
    );
    // B1 (index 0) and B1 (index 2) are the same; C2 (index 1) differs.
    expect(colors[0]).toBe(colors[2]);
    expect(colors[0]).not.toBe(colors[1]);

    // Each unique ref also paints a coloured border on the grid.
    await expect(page.getByTestId("ref-rect-0")).toBeVisible();
    await expect(page.getByTestId("ref-rect-1")).toBeVisible();
  });

  test("range refs render a single border that spans the whole rectangle", async ({ page }) => {
    await gotoXlsxEditor(page);

    await page.getByTestId("cell-A1").click();
    await page.getByTestId("formula-input").click();
    await page.getByTestId("formula-input").fill("=SUM(A2:B3)");

    const range = page.getByTestId("ref-rect-0");
    await expect(range).toBeVisible();

    const cellA2 = await page.getByTestId("cell-A2").boundingBox();
    const cellB3 = await page.getByTestId("cell-B3").boundingBox();
    const rect = await range.boundingBox();
    if (!cellA2 || !cellB3 || !rect) throw new Error("missing bounding boxes");

    // The rect should at minimum cover both A2 and B3 corners.
    expect(rect.x).toBeLessThanOrEqual(cellA2.x + 1);
    expect(rect.y).toBeLessThanOrEqual(cellA2.y + 1);
    expect(rect.x + rect.width).toBeGreaterThanOrEqual(cellB3.x + cellB3.width - 1);
    expect(rect.y + rect.height).toBeGreaterThanOrEqual(cellB3.y + cellB3.height - 1);
  });

  test("highlights disappear when the formula bar loses focus", async ({ page }) => {
    await gotoXlsxEditor(page);

    await page.getByTestId("cell-A1").click();
    await page.getByTestId("formula-input").fill("=B1+C2");
    await expect(page.getByTestId("ref-rect-0")).toBeVisible();
    await page.getByTestId("formula-input").press("Escape");

    // Escape clears the draft → no highlights for a non-formula cell.
    await expect(page.getByTestId("ref-rect-0")).toHaveCount(0);
  });
});
