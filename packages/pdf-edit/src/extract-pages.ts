import { PDFDocument } from "pdf-lib";
import { loadPdf, validatePages } from "./internal.js";

export interface ExtractPagesOptions {
  readonly pages: ReadonlyArray<number>;
}

export const extractPages = async (
  buffer: Uint8Array,
  opts: ExtractPagesOptions,
): Promise<Uint8Array> => {
  if (opts.pages.length === 0) {
    throw new Error("pdf-edit/extract-pages: requires at least one page");
  }
  const src = await loadPdf(buffer);
  validatePages(src, opts.pages, "pdf-edit/extract-pages");
  const out = await PDFDocument.create();
  const copied = await out.copyPages(src, opts.pages.map((n) => n - 1));
  for (const page of copied) out.addPage(page);
  return out.save();
};
