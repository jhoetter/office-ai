import { PDFDocument, degrees } from "pdf-lib";
import type { PdfSnapshot } from "../model/types.js";
import { PdfSerializeError } from "./errors.js";

export interface PdfSerializeOptions {
  /**
   * If true (default), preserves the original PDF byte structure as
   * much as possible by loading via pdf-lib (which uses incremental
   * update by default). Set to false to force a full rewrite.
   */
  readonly incremental?: boolean;
}

/**
 * Serialize a PdfSnapshot back into PDF bytes by replaying mutations
 * (page order, page rotation, metadata) onto the original PDF buffer.
 *
 * NOTE: Annotations, form-field values, and bookmark/outline edits are
 * out of scope for this initial serializer — those round-trip via the
 * dedicated capability packages (@officeai/pdf-annotations,
 * @officeai/pdf-forms) per /spec/pdf/editing-pipeline.md.
 */
export const serializePdf = async (
  snapshot: PdfSnapshot,
  originalBuffer: Uint8Array,
  opts: PdfSerializeOptions = {},
): Promise<Uint8Array> => {
  try {
    const pdf = await PDFDocument.load(originalBuffer, {
      updateMetadata: false,
    });

    const originalPageCount = pdf.getPageCount();
    if (originalPageCount === 0) {
      throw new PdfSerializeError("source PDF has zero pages");
    }

    const md = snapshot.root.metadata;
    if (md.title !== undefined) pdf.setTitle(md.title);
    if (md.author !== undefined) pdf.setAuthor(md.author);
    if (md.subject !== undefined) pdf.setSubject(md.subject);
    if (md.keywords !== undefined) pdf.setKeywords([md.keywords]);
    if (md.creator !== undefined) pdf.setCreator(md.creator);
    if (md.producer !== undefined) pdf.setProducer(md.producer);

    const desiredOrder = snapshot.root.pages.map((p) => p.pageNumber);
    const desiredRotations = snapshot.root.pages.map((p) => p.rotation);

    if (desiredOrder.length === originalPageCount) {
      const isIdentity = desiredOrder.every((n, i) => n === i + 1);
      if (!isIdentity) {
        const indices = desiredOrder.map((n) => n - 1);
        const reordered = await PDFDocument.create();
        const copied = await reordered.copyPages(pdf, indices);
        for (const page of copied) reordered.addPage(page);
        if (md.title !== undefined) reordered.setTitle(md.title);
        if (md.author !== undefined) reordered.setAuthor(md.author);
        if (md.subject !== undefined) reordered.setSubject(md.subject);
        if (md.keywords !== undefined) reordered.setKeywords([md.keywords]);
        if (md.creator !== undefined) reordered.setCreator(md.creator);
        if (md.producer !== undefined) reordered.setProducer(md.producer);
        const pages = reordered.getPages();
        pages.forEach((page, i) => page.setRotation(degrees(desiredRotations[i])));
        return reordered.save({ useObjectStreams: opts.incremental !== false });
      }
    } else {
      throw new PdfSerializeError(
        `cannot add or remove pages without re-serialization (got ${desiredOrder.length}, original has ${originalPageCount}). ` +
          `Use @officeai/pdf-edit for page-level structural changes and pass the resulting buffer back through parsePdf.`,
      );
    }

    const pages = pdf.getPages();
    pages.forEach((page, i) => {
      const desired = desiredRotations[i];
      const current = page.getRotation().angle;
      if (desired !== current) page.setRotation(degrees(desired));
    });

    return pdf.save({ useObjectStreams: opts.incremental !== false });
  } catch (err) {
    if (err instanceof PdfSerializeError) throw err;
    throw new PdfSerializeError(
      err instanceof Error ? `Failed to serialize PDF: ${err.message}` : "Failed to serialize PDF",
      err,
    );
  }
};
