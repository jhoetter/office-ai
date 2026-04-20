import { describe, expect, it } from "vitest";
import { PDFDocument, StandardFonts } from "pdf-lib";
import { addTextLayer } from "./text-layer.js";

const buildScanLikePdf = async (): Promise<Uint8Array> => {
  const pdf = await PDFDocument.create();
  const page = pdf.addPage([612, 792]);
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  page.drawText("(scan placeholder)", { x: 50, y: 720, size: 14, font });
  return pdf.save();
};

describe("addTextLayer (with injected recognise)", () => {
  it("appends invisible text spans driven by the recognise function", async () => {
    const buf = await buildScanLikePdf();
    const out = await addTextLayer(buf, {
      recognise: async () => [
        { text: "Hello", x: 50, y: 50, width: 60, height: 14, confidence: 0.95 },
        { text: "World", x: 120, y: 50, width: 60, height: 14, confidence: 0.94 },
      ],
    });
    const reopened = await PDFDocument.load(out);
    expect(reopened.getPageCount()).toBe(1);
    expect(out.byteLength).toBeGreaterThan(buf.byteLength);
  });

  it("rejects out-of-range pages", async () => {
    const buf = await buildScanLikePdf();
    await expect(addTextLayer(buf, { pages: [99], recognise: async () => [] })).rejects.toThrow(
      /out of range/
    );
  });
});
