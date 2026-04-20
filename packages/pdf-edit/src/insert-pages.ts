import { loadPdf, validatePages } from "./internal.js";

export interface InsertPagesOptions {
  readonly source: Uint8Array;
  /** 1-indexed page numbers in `source` to insert. Default: all pages. */
  readonly sourcePages?: ReadonlyArray<number>;
  /** Insertion point in the host (0..N). 0 = start, N = end. */
  readonly at: number;
}

export const insertPages = async (
  hostBuffer: Uint8Array,
  opts: InsertPagesOptions,
): Promise<Uint8Array> => {
  const host = await loadPdf(hostBuffer);
  const src = await loadPdf(opts.source);
  const srcCount = src.getPageCount();
  const hostCount = host.getPageCount();
  if (opts.at < 0 || opts.at > hostCount) {
    throw new Error(`pdf-edit/insert-pages: at ${opts.at} out of range (0..${hostCount})`);
  }
  const srcPages = opts.sourcePages ?? Array.from({ length: srcCount }, (_, i) => i + 1);
  validatePages(src, srcPages, "pdf-edit/insert-pages");

  const copied = await host.copyPages(src, srcPages.map((n) => n - 1));
  for (let i = 0; i < copied.length; i++) host.insertPage(opts.at + i, copied[i]);
  return host.save();
};
