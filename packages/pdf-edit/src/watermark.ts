import { StandardFonts, degrees, rgb } from "pdf-lib";
import { loadPdf } from "./internal.js";

export interface AddWatermarkOptions {
  readonly text: string;
  readonly opacity?: number;
  readonly fontSize?: number;
  /** Rotation in degrees, default 30. */
  readonly rotate?: number;
  /** RGB triple in 0..1, default #888. */
  readonly color?: { r: number; g: number; b: number };
  /** 1-indexed pages to watermark. Default: all. */
  readonly pages?: ReadonlyArray<number>;
}

export const addWatermark = async (buffer: Uint8Array, opts: AddWatermarkOptions): Promise<Uint8Array> => {
  const pdf = await loadPdf(buffer);
  const font = await pdf.embedFont(StandardFonts.HelveticaBold);
  const fontSize = opts.fontSize ?? 60;
  const opacity = opts.opacity ?? 0.18;
  const rotate = opts.rotate ?? 30;
  const color = opts.color ?? { r: 0.4, g: 0.4, b: 0.4 };
  const total = pdf.getPageCount();
  const set = new Set(opts.pages ?? Array.from({ length: total }, (_, i) => i + 1));

  const pages = pdf.getPages();
  pages.forEach((page, i) => {
    if (!set.has(i + 1)) return;
    const { width, height } = page.getSize();
    const textWidth = font.widthOfTextAtSize(opts.text, fontSize);
    page.drawText(opts.text, {
      x: (width - textWidth) / 2,
      y: height / 2,
      size: fontSize,
      font,
      color: rgb(color.r, color.g, color.b),
      opacity,
      rotate: degrees(rotate),
    });
  });

  return pdf.save();
};
