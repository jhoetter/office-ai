import { PDFDocument } from "pdf-lib";
import { loadPdf, validatePages } from "./internal.js";

export interface DeletePagesOptions {
  readonly pages: ReadonlyArray<number>;
}

export const deletePages = async (buffer: Uint8Array, opts: DeletePagesOptions): Promise<Uint8Array> => {
  if (opts.pages.length === 0) {
    throw new Error("pdf-edit/delete-pages: requires at least one page");
  }
  const src = await loadPdf(buffer);
  const total = src.getPageCount();
  if (opts.pages.length >= total) {
    throw new Error("pdf-edit/delete-pages: cannot delete every page");
  }
  validatePages(src, opts.pages, "pdf-edit/delete-pages");
  const drop = new Set(opts.pages);
  const keep = Array.from({ length: total }, (_, i) => i + 1).filter((n) => !drop.has(n));
  const out = await PDFDocument.create();
  const copied = await out.copyPages(
    src,
    keep.map((n) => n - 1)
  );
  for (const page of copied) out.addPage(page);
  return out.save();
};
