import { PDFDocument } from "pdf-lib";
import { loadPdf, validatePages } from "./internal.js";

export interface ReorderPagesOptions {
  /** Permutation of 1..N. Length must match the source page count. */
  readonly order: ReadonlyArray<number>;
}

export const reorderPages = async (
  buffer: Uint8Array,
  opts: ReorderPagesOptions,
): Promise<Uint8Array> => {
  const src = await loadPdf(buffer);
  const total = src.getPageCount();
  if (opts.order.length !== total) {
    throw new Error(
      `pdf-edit/reorder-pages: order length ${opts.order.length} != page count ${total}`,
    );
  }
  const seen = new Set<number>();
  for (const n of opts.order) {
    if (seen.has(n)) throw new Error(`pdf-edit/reorder-pages: duplicate page ${n}`);
    seen.add(n);
  }
  validatePages(src, opts.order, "pdf-edit/reorder-pages");

  const out = await PDFDocument.create();
  const indices = opts.order.map((n) => n - 1);
  const copied = await out.copyPages(src, indices);
  for (const page of copied) out.addPage(page);
  return out.save();
};
