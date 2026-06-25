import { addAnnotations, type AnnotationInput } from "@officeai/pdf-annotations";
import { PDFDocument, degrees } from "pdf-lib";
import type { PdfAnnotation, PdfSnapshot } from "../model/types.js";
import { PdfSerializeError } from "./errors.js";

export interface PdfSerializeOptions {
  /**
   * If true (default), preserves the original PDF byte structure as
   * much as possible by loading via pdf-lib (which uses incremental
   * update by default). Set to false to force a full rewrite.
   *
   * Kept for backwards compatibility. New callers should prefer
   * `mode`, which records the requested export policy explicitly.
   */
  readonly incremental?: boolean;
  readonly mode?: PdfExportMode;
}

export type PdfExportMode = "auto" | "incremental" | "rewrite" | "diagnosticOnly";
export type PdfEffectiveExportMode = "incremental" | "rewrite" | "diagnosticOnly";

export interface PdfExportDiagnostic {
  readonly level: "info" | "warning" | "error";
  readonly code: string;
  readonly message: string;
}

export interface PdfExportPlan {
  readonly requestedMode: PdfExportMode;
  readonly effectiveMode: PdfEffectiveExportMode;
  readonly incremental: boolean;
  readonly structuralRewrite: boolean;
  readonly reason: string;
  readonly originalPageCount: number;
  readonly outputPageCount: number;
  readonly sessionAnnotationCount: number;
  readonly writableAnnotationCount: number;
  readonly skippedAnnotationCount: number;
  readonly hasSignatures: boolean;
  readonly encrypted: boolean;
  readonly missingTextLayerPages: ReadonlyArray<number>;
  readonly diagnostics: ReadonlyArray<PdfExportDiagnostic>;
}

export interface PdfSerializeResult {
  readonly plan: PdfExportPlan;
  readonly bytes?: Uint8Array;
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
  opts: PdfSerializeOptions = {}
): Promise<Uint8Array> => {
  const result = await serializePdfWithPlan(snapshot, originalBuffer, opts);
  if (!result.bytes) {
    throw new PdfSerializeError("diagnosticOnly PDF export does not produce bytes");
  }
  return result.bytes;
};

export const serializePdfWithPlan = async (
  snapshot: PdfSnapshot,
  originalBuffer: Uint8Array,
  opts: PdfSerializeOptions = {}
): Promise<PdfSerializeResult> => {
  const plan = await planPdfExport(snapshot, originalBuffer, opts);
  if (plan.effectiveMode === "diagnosticOnly") return { plan };
  const blocking = plan.diagnostics.find((diagnostic) => diagnostic.level === "error");
  if (blocking) {
    throw new PdfSerializeError(blocking.message);
  }
  const bytes = await serializePdfBytes(snapshot, originalBuffer, {
    ...opts,
    incremental: plan.effectiveMode === "incremental",
  });
  return { bytes, plan };
};

export const planPdfExport = async (
  snapshot: PdfSnapshot,
  originalBuffer: Uint8Array,
  opts: PdfSerializeOptions = {}
): Promise<PdfExportPlan> => {
  try {
    const pdf = await PDFDocument.load(originalBuffer, {
      updateMetadata: false,
    });
    const originalPageCount = pdf.getPageCount();
    if (originalPageCount === 0) {
      throw new PdfSerializeError("source PDF has zero pages");
    }
    return buildExportPlan(snapshot, originalPageCount, opts);
  } catch (err) {
    if (err instanceof PdfSerializeError) throw err;
    throw new PdfSerializeError(
      err instanceof Error ? `Failed to plan PDF export: ${err.message}` : "Failed to plan PDF export",
      err
    );
  }
};

const serializePdfBytes = async (
  snapshot: PdfSnapshot,
  originalBuffer: Uint8Array,
  opts: PdfSerializeOptions = {}
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
          `Use @officeai/pdf-edit for structural inserts/extracts/merges and feed the resulting buffer back through parsePdf.`
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
      err
    );
  }
};

function requestedModeFrom(opts: PdfSerializeOptions): PdfExportMode {
  if (opts.mode) return opts.mode;
  return opts.incremental === false ? "rewrite" : "auto";
}

function buildExportPlan(
  snapshot: PdfSnapshot,
  originalPageCount: number,
  opts: PdfSerializeOptions
): PdfExportPlan {
  const requestedMode = requestedModeFrom(opts);
  const desiredSourceIndices = snapshot.root.pages.map((p) => p.sourceIndex);
  const invalidSourceMapping = desiredSourceIndices.some((i) => i < 0 || i >= originalPageCount);
  const structuralRewrite =
    invalidSourceMapping ||
    desiredSourceIndices.length !== originalPageCount ||
    !desiredSourceIndices.every((srcIdx, i) => srcIdx === i);
  const canUseIncremental = !structuralRewrite;
  const sessionAnnotations = snapshot.root.annotations.filter((a) => a.source === "session");
  const writableAnnotations = sessionAnnotations
    .map(toAnnotationInput)
    .filter((input): input is AnnotationInput => input !== null);
  const skippedAnnotationCount = sessionAnnotations.length - writableAnnotations.length;
  const encrypted =
    snapshot.root.metadata.encryption?.hasUserPassword === true ||
    snapshot.root.metadata.encryption?.hasOwnerPassword === true;
  const missingTextLayerPages = snapshot.root.pages
    .filter((page) => !page.hasTextLayer)
    .map((page) => page.pageNumber);
  const effectiveMode: PdfEffectiveExportMode =
    requestedMode === "diagnosticOnly"
      ? "diagnosticOnly"
      : requestedMode === "rewrite"
        ? "rewrite"
        : canUseIncremental
          ? "incremental"
          : "rewrite";
  const diagnostics: PdfExportDiagnostic[] = [
    {
      level: "info",
      code: "pdf-export-policy",
      message: `PDF export policy requested=${requestedMode}, effective=${effectiveMode}.`,
    },
    {
      level: "info",
      code: "pdf-export-mode",
      message:
        effectiveMode === "incremental"
          ? "PDF export will use the incremental-compatible annotation/update path."
          : effectiveMode === "rewrite"
            ? "PDF export will use the full rewrite path."
            : "PDF export is diagnostic-only; no bytes will be written.",
    },
  ];

  if (requestedMode === "incremental" && !canUseIncremental) {
    diagnostics.push({
      level: "error",
      code: "pdf-export-incremental-unavailable",
      message:
        "Requested incremental PDF export, but the current snapshot requires a full rewrite because page source mappings changed.",
    });
  }
  if (structuralRewrite) {
    diagnostics.push({
      level: "warning",
      code: "pdf-export-rewrite-required",
      message: "Full PDF rewrite required because page order/count/source mappings changed.",
    });
  }
  if (snapshot.root.signatureCount > 0) {
    diagnostics.push({
      level: "warning",
      code: "pdf-export-signature-risk",
      message:
        "Signature fields are present; OfficeAI does not validate signatures and any export may invalidate them.",
    });
  }
  if (encrypted) {
    diagnostics.push({
      level: "warning",
      code: "pdf-export-encryption-risk",
      message:
        "PDF encryption flags are present; export may be limited by owner/user permissions even if parsing succeeded.",
    });
  }
  if (missingTextLayerPages.length > 0) {
    diagnostics.push({
      level: "warning",
      code: "pdf-export-text-layer-missing",
      message: `Page(s) ${missingTextLayerPages.join(", ")} lack a text layer; text-based annotation placement may need OCR.`,
    });
  }
  if (skippedAnnotationCount > 0) {
    diagnostics.push({
      level: "warning",
      code: "pdf-export-annotation-skipped",
      message: `${skippedAnnotationCount} session annotation(s) are not writable by the current PDF annotation writer.`,
    });
  }

  return {
    requestedMode,
    effectiveMode,
    incremental: effectiveMode === "incremental",
    structuralRewrite,
    reason: structuralRewrite
      ? "page source mappings changed"
      : effectiveMode === "diagnosticOnly"
        ? "diagnostic-only requested"
        : "page source mappings are unchanged",
    originalPageCount,
    outputPageCount: snapshot.root.pages.length,
    sessionAnnotationCount: sessionAnnotations.length,
    writableAnnotationCount: writableAnnotations.length,
    skippedAnnotationCount,
    hasSignatures: snapshot.root.signatureCount > 0,
    encrypted,
    missingTextLayerPages,
    diagnostics,
  };
}

/**
 * Map our document-model annotation projection onto the writer's
 * input shape. Returns `null` for kinds we don't (yet) round-trip —
 * the editor surfaces only Highlight + Sticky for now, and free-text
 * lands in a follow-up.
 */
function toAnnotationInput(a: PdfAnnotation): AnnotationInput | null {
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
}
