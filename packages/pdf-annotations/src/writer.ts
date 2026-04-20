import {
  PDFArray,
  PDFDocument,
  PDFHexString,
  PDFName,
  PDFNumber,
  PDFRef,
  PDFString,
} from "pdf-lib";
import type { AnnotationInput, ColorInput, RectInput } from "./types.js";

export interface AddAnnotationsOptions {
  readonly annotations: ReadonlyArray<AnnotationInput>;
}

const isoNow = (): string => new Date().toISOString();

const pdfDate = (iso: string): string => {
  const d = new Date(iso);
  const pad = (n: number, w = 2): string => String(n).padStart(w, "0");
  return (
    "D:" +
    d.getUTCFullYear() +
    pad(d.getUTCMonth() + 1) +
    pad(d.getUTCDate()) +
    pad(d.getUTCHours()) +
    pad(d.getUTCMinutes()) +
    pad(d.getUTCSeconds()) +
    "Z"
  );
};

const colorTuple = (c: ColorInput | undefined, fallback: ColorInput): [number, number, number] => {
  const v = c ?? fallback;
  return [v.r, v.g, v.b];
};

const rectArray = (pdf: PDFDocument, rect: RectInput): PDFArray => {
  const arr = pdf.context.obj([
    PDFNumber.of(rect[0]),
    PDFNumber.of(rect[1]),
    PDFNumber.of(rect[2]),
    PDFNumber.of(rect[3]),
  ]) as PDFArray;
  return arr;
};

const colorArray = (pdf: PDFDocument, c: [number, number, number]): PDFArray =>
  pdf.context.obj([PDFNumber.of(c[0]), PDFNumber.of(c[1]), PDFNumber.of(c[2])]) as PDFArray;

const ensureAnnots = (pdf: PDFDocument, pageIndex: number): PDFArray => {
  const page = pdf.getPage(pageIndex);
  const node = page.node;
  const annots = node.get(PDFName.of("Annots"));
  if (annots === undefined) {
    const arr = pdf.context.obj([]) as PDFArray;
    node.set(PDFName.of("Annots"), arr);
    return arr;
  }
  if (annots instanceof PDFRef) {
    const looked = pdf.context.lookup(annots);
    if (looked instanceof PDFArray) return looked;
  }
  if (annots instanceof PDFArray) return annots;
  const arr = pdf.context.obj([]) as PDFArray;
  node.set(PDFName.of("Annots"), arr);
  return arr;
};

const appendAnnotation = (
  pdf: PDFDocument,
  pageIndex: number,
  dict: Record<string, unknown>,
): void => {
  const ref = pdf.context.register(pdf.context.obj(dict as never));
  const annots = ensureAnnots(pdf, pageIndex);
  annots.push(ref);
};

/**
 * Append PDF annotations onto an existing document. Implemented for the
 * highest-value subset (highlight, sticky-note, free-text, link). The
 * remaining subtypes documented in /spec/pdf/annotation-model.md are
 * structural extensions of this writer.
 */
export const addAnnotations = async (
  buffer: Uint8Array,
  opts: AddAnnotationsOptions,
): Promise<Uint8Array> => {
  const pdf = await PDFDocument.load(buffer, { updateMetadata: false });
  const total = pdf.getPageCount();

  for (const ann of opts.annotations) {
    if (ann.pageNumber < 1 || ann.pageNumber > total) {
      throw new Error(`pdf-annotations: page ${ann.pageNumber} out of range (1..${total})`);
    }
    const pageIndex = ann.pageNumber - 1;
    const created = pdfDate(isoNow());
    const author = ann.author ?? "Office AI";

    switch (ann.kind) {
      case "highlight": {
        const color = colorTuple(ann.color, { r: 1, g: 0.95, b: 0 });
        const [x1, y1, x2, y2] = ann.rect;
        const quad = pdf.context.obj([
          PDFNumber.of(x1),
          PDFNumber.of(y2),
          PDFNumber.of(x2),
          PDFNumber.of(y2),
          PDFNumber.of(x1),
          PDFNumber.of(y1),
          PDFNumber.of(x2),
          PDFNumber.of(y1),
        ]);
        appendAnnotation(pdf, pageIndex, {
          Type: PDFName.of("Annot"),
          Subtype: PDFName.of("Highlight"),
          Rect: rectArray(pdf, ann.rect),
          QuadPoints: quad,
          C: colorArray(pdf, color),
          T: PDFString.of(author),
          M: PDFString.of(created),
          CreationDate: PDFString.of(created),
          F: PDFNumber.of(4),
          Contents: PDFHexString.fromText(ann.contents ?? ann.text ?? ""),
        });
        break;
      }
      case "sticky-note": {
        const color = colorTuple(ann.color, { r: 1, g: 0.85, b: 0.2 });
        appendAnnotation(pdf, pageIndex, {
          Type: PDFName.of("Annot"),
          Subtype: PDFName.of("Text"),
          Rect: rectArray(pdf, ann.rect),
          C: colorArray(pdf, color),
          T: PDFString.of(author),
          M: PDFString.of(created),
          CreationDate: PDFString.of(created),
          F: PDFNumber.of(4),
          Open: false,
          Name: PDFName.of("Comment"),
          Contents: PDFHexString.fromText(ann.contents),
        });
        break;
      }
      case "free-text": {
        const color = colorTuple(ann.color, { r: 0, g: 0, b: 0 });
        const fontSize = ann.fontSize ?? 12;
        const da = `0 0 0 rg /Helv ${fontSize} Tf`;
        appendAnnotation(pdf, pageIndex, {
          Type: PDFName.of("Annot"),
          Subtype: PDFName.of("FreeText"),
          Rect: rectArray(pdf, ann.rect),
          C: colorArray(pdf, color),
          T: PDFString.of(author),
          M: PDFString.of(created),
          CreationDate: PDFString.of(created),
          F: PDFNumber.of(4),
          DA: PDFString.of(da),
          Contents: PDFHexString.fromText(ann.contents),
        });
        break;
      }
      case "link": {
        const dict: Record<string, unknown> = {
          Type: PDFName.of("Annot"),
          Subtype: PDFName.of("Link"),
          Rect: rectArray(pdf, ann.rect),
          F: PDFNumber.of(4),
          Border: pdf.context.obj([PDFNumber.of(0), PDFNumber.of(0), PDFNumber.of(0)]),
        };
        if (ann.url !== undefined) {
          dict.A = pdf.context.obj({
            Type: PDFName.of("Action"),
            S: PDFName.of("URI"),
            URI: PDFString.of(ann.url),
          });
        } else if (ann.destPage !== undefined) {
          if (ann.destPage < 1 || ann.destPage > total) {
            throw new Error(`pdf-annotations/link: destPage ${ann.destPage} out of range`);
          }
          const targetRef = pdf.getPage(ann.destPage - 1).ref;
          dict.Dest = pdf.context.obj([targetRef, PDFName.of("Fit")]);
        }
        appendAnnotation(pdf, pageIndex, dict);
        break;
      }
    }
  }

  return pdf.save();
};
