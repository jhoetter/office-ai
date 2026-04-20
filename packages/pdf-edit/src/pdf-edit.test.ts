import { describe, expect, it } from "vitest";
import { PDFDocument, StandardFonts } from "pdf-lib";
import { rotatePages } from "./rotate-pages.js";
import { reorderPages } from "./reorder-pages.js";
import { deletePages } from "./delete-pages.js";
import { extractPages } from "./extract-pages.js";
import { mergePdfs } from "./merge.js";
import { splitPdf } from "./split.js";
import { setMetadata } from "./set-metadata.js";
import { addPageNumbers } from "./add-page-numbers.js";
import { addWatermark } from "./watermark.js";
import { cropPages } from "./crop.js";
import { insertPages } from "./insert-pages.js";

const buildPdf = async (n: number, label = "P"): Promise<Uint8Array> => {
  const pdf = await PDFDocument.create();
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  for (let i = 1; i <= n; i++) {
    const page = pdf.addPage([612, 792]);
    page.drawText(`${label}-${i}`, { x: 50, y: 720, size: 24, font });
  }
  return pdf.save();
};

const pageCount = async (buf: Uint8Array): Promise<number> => (await PDFDocument.load(buf)).getPageCount();

describe("pdf-edit", () => {
  it("rotatePages applies a delta", async () => {
    const out = await rotatePages(await buildPdf(2), { pages: [1], delta: 90 });
    const reopened = await PDFDocument.load(out);
    expect(reopened.getPage(0).getRotation().angle).toBe(90);
    expect(reopened.getPage(1).getRotation().angle).toBe(0);
  });

  it("reorderPages permutes pages", async () => {
    const out = await reorderPages(await buildPdf(3), { order: [3, 1, 2] });
    expect(await pageCount(out)).toBe(3);
  });

  it("deletePages removes pages", async () => {
    const out = await deletePages(await buildPdf(4), { pages: [2, 4] });
    expect(await pageCount(out)).toBe(2);
  });

  it("extractPages keeps only the requested pages", async () => {
    const out = await extractPages(await buildPdf(5), { pages: [2, 4] });
    expect(await pageCount(out)).toBe(2);
  });

  it("mergePdfs concatenates inputs", async () => {
    const a = await buildPdf(2, "A");
    const b = await buildPdf(3, "B");
    const out = await mergePdfs({ inputs: [a, b] });
    expect(await pageCount(out)).toBe(5);
  });

  it("splitPdf single yields one PDF per page", async () => {
    const res = await splitPdf(await buildPdf(3), { strategy: { kind: "single" } });
    expect(res.parts.length).toBe(3);
    expect(await pageCount(res.parts[0].bytes)).toBe(1);
  });

  it("splitPdf every chunks pages", async () => {
    const res = await splitPdf(await buildPdf(5), { strategy: { kind: "every", n: 2 } });
    expect(res.parts.length).toBe(3);
    expect(await pageCount(res.parts[0].bytes)).toBe(2);
    expect(await pageCount(res.parts[2].bytes)).toBe(1);
  });

  it("setMetadata writes metadata fields", async () => {
    const out = await setMetadata(await buildPdf(1), { title: "X", author: "Y" });
    const reopened = await PDFDocument.load(out);
    expect(reopened.getTitle()).toBe("X");
    expect(reopened.getAuthor()).toBe("Y");
  });

  it("addPageNumbers does not break the document", async () => {
    const out = await addPageNumbers(await buildPdf(3), { position: "bottom-center" });
    expect(await pageCount(out)).toBe(3);
  });

  it("addWatermark does not break the document", async () => {
    const out = await addWatermark(await buildPdf(2), { text: "DRAFT" });
    expect(await pageCount(out)).toBe(2);
  });

  it("cropPages shrinks the crop box", async () => {
    const out = await cropPages(await buildPdf(1), { margins: [10, 10, 10, 10] });
    const reopened = await PDFDocument.load(out);
    const box = reopened.getPage(0).getCropBox();
    expect(box.width).toBe(612 - 20);
    expect(box.height).toBe(792 - 20);
  });

  it("insertPages inserts source pages at index", async () => {
    const host = await buildPdf(2, "H");
    const src = await buildPdf(2, "S");
    const out = await insertPages(host, { source: src, at: 1 });
    expect(await pageCount(out)).toBe(4);
  });
});
