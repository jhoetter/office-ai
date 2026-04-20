import { degrees } from "pdf-lib";
import { loadPdf, validatePages } from "./internal.js";

export interface RotatePagesOptions {
  /** 1-indexed page numbers. Use ["all"] to apply to every page. */
  readonly pages: ReadonlyArray<number> | "all";
  /** Delta in degrees, must be a multiple of 90. */
  readonly delta: 90 | 180 | 270 | -90 | -180 | -270;
}

const isValidDelta = (n: number): n is 90 | 180 | 270 | -90 | -180 | -270 =>
  [90, 180, 270, -90, -180, -270].includes(n);

const normalize = (raw: number): 0 | 90 | 180 | 270 => {
  const r = ((raw % 360) + 360) % 360;
  if (r === 0 || r === 90 || r === 180 || r === 270) return r as 0 | 90 | 180 | 270;
  return 0;
};

export const rotatePages = async (buffer: Uint8Array, opts: RotatePagesOptions): Promise<Uint8Array> => {
  if (!isValidDelta(opts.delta)) {
    throw new Error(`pdf-edit/rotate-pages: delta ${opts.delta} is not a multiple of 90`);
  }
  const pdf = await loadPdf(buffer);
  const total = pdf.getPageCount();
  const target = opts.pages === "all" ? Array.from({ length: total }, (_, i) => i + 1) : opts.pages;
  validatePages(pdf, target, "pdf-edit/rotate-pages");

  const set = new Set(target);
  const pages = pdf.getPages();
  pages.forEach((page, i) => {
    if (!set.has(i + 1)) return;
    const next = normalize(page.getRotation().angle + opts.delta);
    page.setRotation(degrees(next));
  });
  return pdf.save();
};
