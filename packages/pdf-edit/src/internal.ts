import { PDFDocument } from "pdf-lib";

export const loadPdf = async (buffer: Uint8Array): Promise<PDFDocument> =>
  PDFDocument.load(buffer, { updateMetadata: false });

export const validatePages = (pdf: PDFDocument, pages: ReadonlyArray<number>, label: string): void => {
  const total = pdf.getPageCount();
  for (const p of pages) {
    if (!Number.isInteger(p) || p < 1 || p > total) {
      throw new Error(`${label}: page ${p} out of range (1..${total})`);
    }
  }
};
