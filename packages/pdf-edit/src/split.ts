import { PDFDocument } from "pdf-lib";
import { loadPdf } from "./internal.js";

export type SplitPdfStrategy =
  | { readonly kind: "ranges"; readonly ranges: ReadonlyArray<readonly [number, number]> }
  | { readonly kind: "every"; readonly n: number }
  | { readonly kind: "single" };

export interface SplitPdfOptions {
  readonly strategy: SplitPdfStrategy;
}

export interface SplitPdfPart {
  readonly label: string;
  readonly pages: ReadonlyArray<number>;
  readonly bytes: Uint8Array;
}

export interface SplitPdfResult {
  readonly parts: ReadonlyArray<SplitPdfPart>;
}

const buildPart = async (src: PDFDocument, pages: ReadonlyArray<number>): Promise<Uint8Array> => {
  const out = await PDFDocument.create();
  const copied = await out.copyPages(
    src,
    pages.map((n) => n - 1)
  );
  for (const page of copied) out.addPage(page);
  return out.save();
};

export const splitPdf = async (buffer: Uint8Array, opts: SplitPdfOptions): Promise<SplitPdfResult> => {
  const src = await loadPdf(buffer);
  const total = src.getPageCount();
  const parts: SplitPdfPart[] = [];

  switch (opts.strategy.kind) {
    case "single": {
      for (let i = 1; i <= total; i++) {
        parts.push({
          label: `page-${i}`,
          pages: [i],
          bytes: await buildPart(src, [i]),
        });
      }
      break;
    }
    case "every": {
      const n = opts.strategy.n;
      if (!Number.isInteger(n) || n < 1) throw new Error(`pdf-edit/split: every.n must be ≥ 1`);
      for (let start = 1; start <= total; start += n) {
        const end = Math.min(total, start + n - 1);
        const pages = Array.from({ length: end - start + 1 }, (_, i) => start + i);
        parts.push({
          label: `pages-${start}-${end}`,
          pages,
          bytes: await buildPart(src, pages),
        });
      }
      break;
    }
    case "ranges": {
      for (const [start, end] of opts.strategy.ranges) {
        if (start < 1 || end > total || end < start) {
          throw new Error(`pdf-edit/split: invalid range ${start}-${end} (1..${total})`);
        }
        const pages = Array.from({ length: end - start + 1 }, (_, i) => start + i);
        parts.push({
          label: `pages-${start}-${end}`,
          pages,
          bytes: await buildPart(src, pages),
        });
      }
      break;
    }
  }

  return { parts };
};
