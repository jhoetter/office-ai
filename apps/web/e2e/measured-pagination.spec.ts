import { expect, test } from "@playwright/test";
import { focusEditor, gotoEditor } from "./_helpers";

/**
 * P3.8 follow-up — measured pagination.
 *
 * The chunker has always supported `(blockIndex) => height-in-twips`
 * as a third break signal (after hard `<w:br w:type="page"/>` and the
 * `<w:lastRenderedPageBreak/>` hint). Wired up in this batch: the
 * page-decorations PM plugin reads `getBoundingClientRect()` per
 * top-level body block, converts CSS px → twips, and re-runs the
 * chunker so content overflowing the section's `pgSz - pgMar.top -
 * pgMar.bottom` content area flows onto a new page sheet.
 *
 * Without measured pagination the bundled welcome doc renders as a
 * single page sheet regardless of how many paragraphs the user adds.
 * With it, we expect at least one `pm-page-edge` widget (the visible
 * grey gap that separates two adjacent page sheets) once enough
 * paragraphs have been typed.
 */
test.describe("editor: measured pagination", () => {
  test("a long body produces multiple page sheets without any hard break", async ({ page }) => {
    await gotoEditor(page);

    const surface = page.locator(".ProseMirror").first();
    await expect(surface).toBeVisible();

    // Baseline: bundled welcome doc fits on one sheet.
    await expect(surface.locator(".pm-page-edge")).toHaveCount(0);

    // Type enough paragraphs to push past DIN A4 content height
    // (~933 CSS px = (16838 - 1417 - 1417) / 15). Each Enter creates a
    // new paragraph; each line is ~28 CSS px in the editor's default
    // 16px / 1.7 line-height. Sixty paragraphs ≈ 1700 px > 933 px so
    // we should overflow into at least one extra page.
    await focusEditor(page);
    await page.keyboard.press("ControlOrMeta+End");
    for (let i = 0; i < 60; i++) {
      await page.keyboard.press("Enter");
      await page.keyboard.type(`Paragraph ${i + 1} — pagination filler`);
    }

    // Allow the measurement RAF + force-recompute round trip to land.
    // The plugin runs one extra cycle after each measurement-driven
    // re-chunk; a short waitFor against the edge widget is enough.
    const edges = surface.locator(".pm-page-edge");
    await expect.poll(async () => edges.count(), { timeout: 5_000 }).toBeGreaterThan(0);

    // Sanity: every page-edge separator carries a `data-page-number`
    // attribute on its inner gap so the visual break is discoverable
    // by tests + assistive tech. The user-facing page number lives in
    // the status bar, not in a banner inside the gap (Word does not
    // paint a "Page N" pill between sheets either).
    const firstEdgeGap = edges.first().locator(".pm-page-gap");
    await expect(firstEdgeGap).toHaveAttribute("data-page-number", /\d+/);
  });

  test("editor card width matches US-Letter (12240 twips ≈ 816 CSS px)", async ({ page }) => {
    await gotoEditor(page);
    const surface = page.locator(".ProseMirror").first();
    await expect(surface).toBeVisible();

    // The bundled welcome doc declares `<w:pgSz w:w="12240"/>`. With
    // the `--pm-page-width` CSS variable plumbed through DocxEditor
    // we expect the white card to render at 12240/15 = 816 CSS px
    // (give or take sub-pixel rounding).
    const width = await surface.evaluate((el) => el.getBoundingClientRect().width);
    expect(width).toBeGreaterThan(810);
    expect(width).toBeLessThan(820);
  });

  test("page sheet always renders at full page height (Word-style fixed sheets)", async ({ page }) => {
    await gotoEditor(page);
    const surface = page.locator(".ProseMirror").first();
    await expect(surface).toBeVisible();

    // The bundled welcome doc has only ~5 short paragraphs that
    // collectively take far less vertical space than a US-Letter
    // content area (~864 CSS px). Without the per-chunk filler the
    // white card would hug the content and the sheet would end just
    // below "command bus, the same path…". With the filler, every
    // chunk pads its body region down to the full content area so
    // the sheet visually matches a Word page even when it's mostly
    // empty.
    //
    // Wait for the measurement RAF round-trip to settle before
    // sampling: the filler shrinks from "full content area" to its
    // true size on the second frame.
    await expect
      .poll(async () => surface.locator(".pm-page-filler").count(), {
        timeout: 5_000,
      })
      .toBeGreaterThan(0);

    const fillerHeight = await surface
      .locator(".pm-page-filler")
      .first()
      .evaluate((el) => el.getBoundingClientRect().height);
    // For US-Letter @ 1" margins the content area is (15840 - 1440 -
    // 1440) / 15 = 864 CSS px. The welcome body is ~150 CSS px so
    // the filler should take up the lion's share of the remaining
    // space — easily > 400 px.
    expect(fillerHeight).toBeGreaterThan(400);

    // And the white card itself must therefore be at least the
    // content-area height (864 px) — taller still once cap-top /
    // cap-bottom chrome is added on top.
    const cardHeight = await surface.evaluate((el) => el.getBoundingClientRect().height);
    expect(cardHeight).toBeGreaterThan(800);
  });
});
