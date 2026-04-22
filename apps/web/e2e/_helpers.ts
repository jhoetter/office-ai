import type { Page } from "@playwright/test";
import { expect } from "@playwright/test";

/**
 * Shared helpers for the smoke specs. Each spec navigates to /editor and
 * waits for the bundled sample document to mount before doing anything
 * else; the editor's "Loading…" overlay disappears once the agent has
 * mounted and the ProseMirror surface is hydrated.
 */
export async function gotoEditor(page: Page): Promise<void> {
  await page.goto("/editor");
  // Wait until the sample document text is visible — that's only true
  // once mountAgent() has finished and the PM doc has been projected.
  await expect(page.getByText("Welcome to office-ai", { exact: false })).toBeVisible({
    timeout: 15_000,
  });
}

/**
 * Navigate to the XLSX editor and wait until the synthetic workbook is
 * mounted. The seeded `Score` header cell only renders once
 * `XlsxAgent.fromBuffer` has finished and the grid has been projected.
 */
export async function gotoXlsxEditor(page: Page): Promise<void> {
  await page.goto("/xlsx-editor");
  await expect(page.getByTestId("cell-B1")).toContainText("Score", { timeout: 15_000 });
}

/** Place caret at the very start of the editable surface. */
export async function focusEditor(page: Page): Promise<void> {
  const surface = page.locator(".ProseMirror").first();
  await surface.click();
  await surface.focus();
}

/**
 * Select every character in the editor by issuing Cmd/Ctrl+A. Doing
 * select-all from the editor surface is more robust across browsers
 * than counting keystrokes from a `Cmd+Home` origin (PM does not always
 * honour the macOS document-start shortcut) or relying on a synthesized
 * dblclick to fire native word selection.
 */
export async function selectAll(page: Page): Promise<void> {
  await focusEditor(page);
  await page.keyboard.press("ControlOrMeta+a");
}

/**
 * Select the entire textblock that contains `word` via a triple-click.
 * ProseMirror's default `handleTripleClick` selects the parent textblock,
 * which gives us a single-paragraph selection — exactly what handlers
 * like `docx:add-comment` accept (they reject cross-paragraph ranges).
 */
export async function selectParagraphContaining(page: Page, word: string): Promise<void> {
  const target = page.locator(".ProseMirror").getByText(word, { exact: false }).first();
  await target.scrollIntoViewIfNeeded();
  await target.click({ clickCount: 3 });
}
