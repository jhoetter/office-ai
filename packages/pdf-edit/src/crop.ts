import { loadPdf, validatePages } from "./internal.js";

export interface CropPagesOptions {
  /** Margins to crop, in PDF user-units. Order: [left, top, right, bottom]. */
  readonly margins: readonly [number, number, number, number];
  /** 1-indexed pages to crop. Default: all. */
  readonly pages?: ReadonlyArray<number>;
}

export const cropPages = async (
  buffer: Uint8Array,
  opts: CropPagesOptions,
): Promise<Uint8Array> => {
  const pdf = await loadPdf(buffer);
  const total = pdf.getPageCount();
  const target = opts.pages ?? Array.from({ length: total }, (_, i) => i + 1);
  validatePages(pdf, target, "pdf-edit/crop");
  const set = new Set(target);
  const [left, top, right, bottom] = opts.margins;
  if (left < 0 || top < 0 || right < 0 || bottom < 0) {
    throw new Error("pdf-edit/crop: margins must be ≥ 0");
  }
  const pages = pdf.getPages();
  pages.forEach((page, i) => {
    if (!set.has(i + 1)) return;
    const box = page.getCropBox();
    const newWidth = box.width - left - right;
    const newHeight = box.height - top - bottom;
    if (newWidth <= 0 || newHeight <= 0) {
      throw new Error(`pdf-edit/crop: margins exceed page ${i + 1} size`);
    }
    page.setCropBox(box.x + left, box.y + bottom, newWidth, newHeight);
  });
  return pdf.save();
};
