import { describe, expect, it } from "vitest";
import { PdfAgent } from "@officeai/pdf";
import { addWatermark } from "@officeai/pdf-edit";
import { isPdfBytes, loadFixture } from "./helpers.js";

/**
 * Annotation roundtrip: ensure highlight, sticky note, and link
 * annotations on the input fixtures survive a parse + re-parse, and
 * that an `addWatermark` mutation produces a still-parseable PDF.
 *
 * The serializer doesn't actively touch annotations, so the goal here
 * is to guard against the parser silently dropping kinds.
 */

describe("PDF roundtrip — annotations", () => {
  it("with-link-annot.pdf surfaces a single link annotation with URL", async () => {
    const agent = await PdfAgent.fromBuffer(await loadFixture("with-link-annot.pdf"));
    const annots = agent.getSnapshot().root.annotations;
    expect(annots).toHaveLength(1);
    expect(annots[0].kind).toBe("link");
    expect(annots[0].pageNumber).toBe(1);
    // pdf.js normalizes URI annotations through the URL constructor,
    // so the trailing slash on the host-only form is expected.
    expect(annots[0].url).toMatch(/^https:\/\/cursor\.com\/?$/);
  });

  it("with-highlight-annot.pdf surfaces highlight + sticky note", async () => {
    const agent = await PdfAgent.fromBuffer(await loadFixture("with-highlight-annot.pdf"));
    const annots = agent.getSnapshot().root.annotations;
    expect(annots.length).toBeGreaterThanOrEqual(2);
    const kinds = annots.map((a) => a.kind).sort();
    expect(kinds).toContain("highlight");
    expect(kinds).toContain("note");
    for (const a of annots) {
      expect(a.pageNumber).toBe(1);
      expect(a.rect).toHaveLength(4);
    }
    expect(agent.getSnapshot().root.pages[0].hasAnnotations).toBe(true);
  });

  it("annotations survive a no-op exportFile + re-parse", async () => {
    const bytes = await loadFixture("with-highlight-annot.pdf");
    const before = (await PdfAgent.fromBuffer(bytes)).getSnapshot().root.annotations.length;
    const agent = await PdfAgent.fromBuffer(bytes);
    const exported = await agent.exportFile();
    const after = (await PdfAgent.fromBuffer(exported)).getSnapshot().root.annotations.length;
    expect(after).toBe(before);
  });

  it("addWatermark on simple-text-1page.pdf produces a still-parseable PDF", async () => {
    const bytes = await loadFixture("simple-text-1page.pdf");
    const stamped = await addWatermark(bytes, { text: "DRAFT", opacity: 0.2 });
    expect(isPdfBytes(stamped)).toBe(true);
    const agent = await PdfAgent.fromBuffer(stamped);
    expect(agent.getSnapshot().root.pages).toHaveLength(1);
    // Watermark text isn't an annotation; we just want to confirm the
    // mutated PDF still parses with the same page count.
  });
});
