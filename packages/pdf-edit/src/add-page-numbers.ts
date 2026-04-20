import { StandardFonts, rgb } from "pdf-lib";
import { loadPdf } from "./internal.js";

export type PageNumberPosition =
  | "top-left"
  | "top-center"
  | "top-right"
  | "bottom-left"
  | "bottom-center"
  | "bottom-right";

export interface AddPageNumbersOptions {
  readonly position?: PageNumberPosition;
  readonly fontSize?: number;
  readonly margin?: number;
  /** First page that should display a number (1-indexed). Default: 1. */
  readonly startAt?: number;
  /** Format string with `{page}` and `{total}` placeholders. */
  readonly format?: string;
}

const DEFAULTS = {
  position: "bottom-center" as const,
  fontSize: 10,
  margin: 24,
  startAt: 1,
  format: "{page} / {total}",
};

export const addPageNumbers = async (
  buffer: Uint8Array,
  opts: AddPageNumbersOptions = {},
): Promise<Uint8Array> => {
  const pdf = await loadPdf(buffer);
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const merged = { ...DEFAULTS, ...opts };
  const total = pdf.getPageCount();

  const pages = pdf.getPages();
  pages.forEach((page, i) => {
    const pageNumber = i + 1;
    if (pageNumber < merged.startAt) return;
    const text = merged.format
      .replace("{page}", String(pageNumber))
      .replace("{total}", String(total));
    const { width, height } = page.getSize();
    const textWidth = font.widthOfTextAtSize(text, merged.fontSize);

    let x = merged.margin;
    let y = merged.margin;
    if (merged.position === "top-left" || merged.position === "top-center" || merged.position === "top-right") {
      y = height - merged.margin - merged.fontSize;
    }
    if (merged.position === "top-center" || merged.position === "bottom-center") {
      x = (width - textWidth) / 2;
    } else if (merged.position === "top-right" || merged.position === "bottom-right") {
      x = width - merged.margin - textWidth;
    }
    page.drawText(text, {
      x,
      y,
      size: merged.fontSize,
      font,
      color: rgb(0, 0, 0),
    });
  });

  return pdf.save();
};
