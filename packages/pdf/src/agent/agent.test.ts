import { afterEach, describe, expect, it } from "vitest";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";
import { PdfAgent } from "./agent.js";

const buildSamplePdf = async (): Promise<Uint8Array> => {
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const page1 = pdf.addPage([612, 792]);
  page1.drawText("Hello world. The quick brown fox jumps over the lazy dog.", {
    x: 50,
    y: 720,
    size: 14,
    font,
    color: rgb(0, 0, 0),
  });
  const page2 = pdf.addPage([612, 792]);
  page2.drawText("Second page text. Lorem ipsum dolor sit amet.", {
    x: 50,
    y: 720,
    size: 14,
    font,
  });
  const page3 = pdf.addPage([612, 792]);
  page3.drawText("Third page. The fox returns again.", {
    x: 50,
    y: 720,
    size: 14,
    font,
  });
  pdf.setTitle("Test Doc");
  pdf.setAuthor("Office AI");
  return pdf.save();
};

describe("PdfAgent", () => {
  let agent: PdfAgent;

  afterEach(() => {
    agent = undefined as unknown as PdfAgent;
  });

  it("parses a PDF and exposes its pages + metadata", async () => {
    const buf = await buildSamplePdf();
    agent = await PdfAgent.fromBuffer(buf);
    const snap = agent.getSnapshot();
    expect(snap.root.pages.length).toBe(3);
    expect(snap.root.metadata.title).toBe("Test Doc");
    expect(snap.root.metadata.author).toBe("Office AI");
    expect(snap.root.pages[0].text.length).toBeGreaterThan(0);
    expect(snap.root.pages[0].hasTextLayer).toBe(true);
  });

  it("renders a markdown projection that includes per-page text", async () => {
    const buf = await buildSamplePdf();
    agent = await PdfAgent.fromBuffer(buf);
    const md = agent.toMarkdown();
    expect(md).toContain("# Test Doc");
    expect(md).toContain("### Page 1");
    expect(md).toContain("### Page 2");
    expect(md).toContain("### Page 3");
  });

  it("searches across pages and reports per-page hits", async () => {
    const buf = await buildSamplePdf();
    agent = await PdfAgent.fromBuffer(buf);
    const hits = agent.search({ query: "fox" });
    expect(hits.length).toBeGreaterThanOrEqual(2);
    expect(hits.every((h) => h.match.toLowerCase() === "fox")).toBe(true);
  });

  it("supports range queries", async () => {
    const buf = await buildSamplePdf();
    agent = await PdfAgent.fromBuffer(buf);
    const range = agent.getRange({ kind: "pdf-pages", start: 2, end: 3 });
    expect(range.pages.map((p) => p.pageNumber)).toEqual([2, 3]);
  });

  it("applies a rotate-pages command and persists it through export", async () => {
    const buf = await buildSamplePdf();
    agent = await PdfAgent.fromBuffer(buf);
    const mutation = await agent.applyCommand({
      type: "pdf:rotate-pages",
      payload: { pages: [1], delta: 90 },
    });
    expect(mutation.diff.changes.length).toBe(1);
    expect(agent.getSnapshot().root.pages[0].rotation).toBe(90);
    const out = await agent.exportFile();
    expect(out.byteLength).toBeGreaterThan(0);

    const reopened = await PdfAgent.fromBuffer(out);
    expect(reopened.getSnapshot().root.pages[0].rotation).toBe(90);
  });

  it("applies reorder-pages and persists the new order through export", async () => {
    const buf = await buildSamplePdf();
    agent = await PdfAgent.fromBuffer(buf);
    await agent.applyCommand({
      type: "pdf:reorder-pages",
      payload: { order: [3, 1, 2] },
    });
    const out = await agent.exportFile();
    const reopened = await PdfAgent.fromBuffer(out);
    expect(reopened.getSnapshot().root.pages[0].text).toContain("Third page");
    expect(reopened.getSnapshot().root.pages[1].text).toContain("Hello world");
  });
});
