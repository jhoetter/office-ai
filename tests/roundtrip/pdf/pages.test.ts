import { describe, expect, it } from "vitest";
import { PdfAgent } from "@officeai/pdf";
import { isPdfBytes, loadFixture } from "./helpers.js";

/**
 * Page-level roundtrip checks: rotation, reorder, multi-size, and the
 * 50-page perf-budget fixture. Each test parses a fixture, asserts a
 * structural invariant, applies one mutation through the command bus,
 * exports, re-parses, and verifies the mutation stuck.
 *
 * `sourceIndex` monotonicity is asserted on every fixture as a guard
 * against the parser silently shuffling the source-page mapping.
 */

describe("PDF roundtrip — pages", () => {
  it("rotated-pages.pdf preserves per-page rotation across exportFile", async () => {
    const agent = await PdfAgent.fromBuffer(await loadFixture("rotated-pages.pdf"));
    const pages = agent.getSnapshot().root.pages;
    expect(pages).toHaveLength(4);
    expect(pages.map((p) => p.rotation)).toEqual([0, 90, 180, 270]);
    pages.forEach((p, i) => expect(p.sourceIndex).toBe(i));

    const exported = await agent.exportFile();
    expect(isPdfBytes(exported)).toBe(true);
    const reparsed = await PdfAgent.fromBuffer(exported);
    expect(reparsed.getSnapshot().root.pages.map((p) => p.rotation)).toEqual([0, 90, 180, 270]);
  });

  it("pdf:rotate-pages command bus persists rotation deltas", async () => {
    const agent = await PdfAgent.fromBuffer(await loadFixture("simple-text-3page.pdf"));
    await agent.applyCommand({
      type: "pdf:rotate-pages",
      payload: { pages: [2], delta: 90 },
    });
    const exported = await agent.exportFile();
    const reparsed = await PdfAgent.fromBuffer(exported);
    const rotations = reparsed.getSnapshot().root.pages.map((p) => p.rotation);
    expect(rotations).toEqual([0, 90, 0]);
  });

  it("pdf:reorder-pages command bus persists the new page order", async () => {
    const agent = await PdfAgent.fromBuffer(await loadFixture("simple-text-3page.pdf"));
    const before = agent.getSnapshot().root.pages.map((p) => p.text.replace(/\s+/g, " ").trim());
    expect(before[0]).toMatch(/PAGE_1/);
    expect(before[1]).toMatch(/PAGE_2/);
    expect(before[2]).toMatch(/PAGE_3/);

    await agent.applyCommand({
      type: "pdf:reorder-pages",
      payload: { order: [3, 1, 2] },
    });
    const exported = await agent.exportFile();
    const reparsed = await PdfAgent.fromBuffer(exported);
    const after = reparsed.getSnapshot().root.pages.map((p) => p.text.replace(/\s+/g, " ").trim());
    expect(after[0]).toMatch(/PAGE_3/);
    expect(after[1]).toMatch(/PAGE_1/);
    expect(after[2]).toMatch(/PAGE_2/);
  });

  it("multi-size-pages.pdf reports distinct widths/heights per page", async () => {
    const agent = await PdfAgent.fromBuffer(await loadFixture("multi-size-pages.pdf"));
    const pages = agent.getSnapshot().root.pages;
    expect(pages).toHaveLength(4);
    // Letter (612), A4 (~595), Legal (612 wide too — distinguishable by height), A3 (~842)
    const widths = pages.map((p) => Math.round(p.width));
    const heights = pages.map((p) => Math.round(p.height));
    expect(widths[0]).toBe(612); // Letter
    expect(heights[0]).toBe(792);
    expect(widths[1]).toBe(595); // A4
    expect(widths[2]).toBe(612); // Legal
    expect(heights[2]).toBe(1008);
    expect(widths[3]).toBe(842); // A3
    pages.forEach((p, i) => expect(p.sourceIndex).toBe(i));
  });

  it("large-50page.pdf parses fast and projects 50 pages with monotonic sourceIndex", async () => {
    const t0 = Date.now();
    const agent = await PdfAgent.fromBuffer(await loadFixture("large-50page.pdf"));
    const elapsed = Date.now() - t0;
    const pages = agent.getSnapshot().root.pages;
    expect(pages).toHaveLength(50);
    pages.forEach((p, i) => expect(p.sourceIndex).toBe(i));
    pages.forEach((p, i) => expect(p.pageNumber).toBe(i + 1));
    // Soft budget: tens of pages on cold pdfjs init should finish in
    // under 8s even on slow CI workers. Generous on purpose.
    expect(elapsed).toBeLessThan(8000);
  });

  it("simple-text-1page.pdf round-trips byte-validly through exportFile", async () => {
    const agent = await PdfAgent.fromBuffer(await loadFixture("simple-text-1page.pdf"));
    const exported = await agent.exportFile();
    expect(isPdfBytes(exported)).toBe(true);
    const reparsed = await PdfAgent.fromBuffer(exported);
    expect(reparsed.getSnapshot().root.pages).toHaveLength(1);
    expect(reparsed.getSnapshot().root.pages[0].sourceIndex).toBe(0);
  });
});
