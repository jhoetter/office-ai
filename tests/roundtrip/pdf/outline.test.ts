import { describe, expect, it } from "vitest";
import { PdfAgent } from "@officeai/pdf";
import type { PdfOutlineNode } from "@officeai/pdf";
import { loadFixture } from "./helpers.js";

function flatten(nodes: ReadonlyArray<PdfOutlineNode>): PdfOutlineNode[] {
  const out: PdfOutlineNode[] = [];
  const visit = (n: PdfOutlineNode): void => {
    out.push(n);
    for (const c of n.children) visit(c);
  };
  for (const n of nodes) visit(n);
  return out;
}

/**
 * `with-outline.pdf` has a hand-rolled 2-level outline:
 *   - Chapter 1 — Introduction
 *   - Chapter 2 — Body
 *       § 2.1
 *       § 2.2
 *   - Chapter 3 — Conclusion
 *
 * Tests assert depth, ordering, and page-destination resolution.
 */

describe("PDF roundtrip — outline", () => {
  it("with-outline.pdf surfaces a 2-level outline tree", async () => {
    const agent = await PdfAgent.fromBuffer(await loadFixture("with-outline.pdf"));
    const outline = agent.getSnapshot().root.outline;
    expect(outline).toHaveLength(3);
    expect(outline[0].title).toMatch(/Chapter 1/);
    expect(outline[1].title).toMatch(/Chapter 2/);
    expect(outline[2].title).toMatch(/Chapter 3/);
    expect(outline[0].children).toHaveLength(0);
    expect(outline[1].children).toHaveLength(2);
    expect(outline[2].children).toHaveLength(0);
    expect(outline[1].children[0].title).toMatch(/§2\.1/);
    expect(outline[1].children[1].title).toMatch(/§2\.2/);
  });

  it("outline destinations resolve to in-range page numbers", async () => {
    const agent = await PdfAgent.fromBuffer(await loadFixture("with-outline.pdf"));
    const snap = agent.getSnapshot();
    const total = snap.root.pages.length;
    for (const node of flatten(snap.root.outline)) {
      if (node.pageNumber !== undefined) {
        expect(node.pageNumber).toBeGreaterThanOrEqual(1);
        expect(node.pageNumber).toBeLessThanOrEqual(total);
      }
    }
  });

  it("toMarkdown emits one heading per outline entry plus per-page headings", async () => {
    const agent = await PdfAgent.fromBuffer(await loadFixture("with-outline.pdf"));
    const md = agent.toMarkdown();
    expect(md).toMatch(/Chapter 1/);
    expect(md).toMatch(/Chapter 2/);
    expect(md).toMatch(/Chapter 3/);
    // PdfAgent's markdown projection emits per-page section headers.
    expect(md).toMatch(/Page 1/);
    expect(md).toMatch(/Page 3/);
  });

  it("simple-text-1page.pdf has an empty outline", async () => {
    const agent = await PdfAgent.fromBuffer(await loadFixture("simple-text-1page.pdf"));
    expect(agent.getSnapshot().root.outline).toEqual([]);
  });
});
