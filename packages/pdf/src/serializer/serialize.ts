import { addAnnotations, type AnnotationInput } from "@officeai/pdf-annotations";
import { PDFDocument, degrees } from "pdf-lib";
import type { PdfAnnotation, PdfSnapshot } from "../model/types.js";
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
 * (page order, page rotation, metadata, session annotations) onto the
 * original PDF buffer.
 *
 * Session-added highlight + sticky annotations route through
 * `@officeai/pdf-annotations` after the page-identity copy pass.
 * Form-field values and bookmark/outline edits remain out of scope
 * for this serializer per /spec/pdf/editing-pipeline.md.
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

    const desiredSourceIndices = snapshot.root.pages.map((p) => p.sourceIndex);
    const desiredRotations = snapshot.root.pages.map((p) => p.rotation);

    if (desiredSourceIndices.some((i) => i < 0 || i >= originalPageCount)) {
      throw new PdfSerializeError(
        `serializer cannot resolve source pages: this snapshot contains pages with no source mapping (sourceIndex out of [0..${originalPageCount - 1}]). ` +
          `Use @officeai/pdf-edit for structural inserts/extracts/merges and feed the resulting buffer back through parsePdf.`,
      );
    }

    const isIdentity =
      desiredSourceIndices.length === originalPageCount &&
      desiredSourceIndices.every((srcIdx, i) => srcIdx === i);

    let intermediate: Uint8Array;
    if (!isIdentity) {
      const reordered = await PDFDocument.create();
      const copied = await reordered.copyPages(pdf, desiredSourceIndices);
      for (const page of copied) reordered.addPage(page);
      if (md.title !== undefined) reordered.setTitle(md.title);
      if (md.author !== undefined) reordered.setAuthor(md.author);
      if (md.subject !== undefined) reordered.setSubject(md.subject);
      if (md.keywords !== undefined) reordered.setKeywords([md.keywords]);
      if (md.creator !== undefined) reordered.setCreator(md.creator);
      if (md.producer !== undefined) reordered.setProducer(md.producer);
      const pages = reordered.getPages();
      pages.forEach((page, i) => page.setRotation(degrees(desiredRotations[i])));
      intermediate = await reordered.save({ useObjectStreams: opts.incremental !== false });
    } else {
      const pages = pdf.getPages();
      pages.forEach((page, i) => {
        const desired = desiredRotations[i];
        const current = page.getRotation().angle;
        if (desired !== current) page.setRotation(degrees(desired));
      });
      intermediate = await pdf.save({ useObjectStreams: opts.incremental !== false });
    }

    const sessionAnnotations = snapshot.root.annotations.filter((a) => a.source === "session");
    if (sessionAnnotations.length === 0) return intermediate;

    const writable = sessionAnnotations
      .map(toAnnotationInput)
      .filter((input): input is AnnotationInput => input !== null);
    if (writable.length === 0) return intermediate;

    return addAnnotations(intermediate, { annotations: writable });
  } catch (err) {
    if (err instanceof PdfSerializeError) throw err;
    throw new PdfSerializeError(
      err instanceof Error ? `Failed to serialize PDF: ${err.message}` : "Failed to serialize PDF",
      err,
    );
  }
};

/**
 * Map our document-model annotation projection onto the writer's
 * input shape. Returns `null` for kinds we don't (yet) round-trip —
 * the editor surfaces only Highlight + Sticky for now, and free-text
 * lands in a follow-up.
 */
const toAnnotationInput = (a: PdfAnnotation): AnnotationInput | null => {
  const base = {
    pageNumber: a.pageNumber,
    rect: a.rect,
    ...(a.author !== undefined ? { author: a.author } : {}),
    ...(a.contents !== undefined ? { contents: a.contents } : {}),
    ...(a.color !== undefined ? { color: a.color } : {}),
  } as const;
  switch (a.kind) {
    case "highlight":
      return { ...base, kind: "highlight" };
    case "note":
      return { ...base, kind: "sticky-note", contents: a.contents ?? "" };
    case "free-text":
      return { ...base, kind: "free-text", contents: a.contents ?? "" };
    case "link":
      return {
        ...base,
        kind: "link",
        ...(a.url !== undefined ? { url: a.url } : {}),
        ...(a.destPage !== undefined ? { destPage: a.destPage } : {}),
      };
    default:
      return null;
  }
};
