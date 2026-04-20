/**
 * In-memory PDF document model. Spec: /spec/pdf/document-model.md.
 *
 * The model is intentionally a *projection* of the underlying PDF —
 * not a faithful reconstruction. The original byte buffer is the source
 * of truth for round-trip integrity (incremental update on save).
 *
 * Coordinates use PDF user-space convention: origin at lower-left,
 * units are 1/72 inch by default. Page-level rect: [x1, y1, x2, y2].
 */
import type { DocumentSnapshot, NodeId } from "@officeai/core";

export type PdfRotation = 0 | 90 | 180 | 270;

export type PdfRect = readonly [number, number, number, number];

export interface PdfMetadata {
  readonly title?: string;
  readonly author?: string;
  readonly subject?: string;
  readonly keywords?: string;
  readonly creator?: string;
  readonly producer?: string;
  readonly creationDate?: string;
  readonly modificationDate?: string;
  readonly pdfVersion?: string;
  readonly linearized?: boolean;
  readonly encryption?: { readonly hasUserPassword: boolean; readonly hasOwnerPassword: boolean };
}

export interface PdfPage {
  readonly id: NodeId;
  /** 1-indexed page number in the current document order. */
  readonly pageNumber: number;
  /** PDF user-units. */
  readonly width: number;
  readonly height: number;
  readonly rotation: PdfRotation;
  /** Optional page label, e.g. "iv", "A-12". */
  readonly label?: string;
  /** Plain-text reading-order projection (best-effort, may be empty for scans). */
  readonly text: string;
  /** True if the page has a text layer (selectable text). */
  readonly hasTextLayer: boolean;
  /** True if any annotations live on the page. */
  readonly hasAnnotations: boolean;
  /** True if any form-field widgets live on the page. */
  readonly hasFormFields: boolean;
}

export interface PdfOutlineNode {
  readonly id: NodeId;
  readonly title: string;
  /** 1-indexed page destination, if resolvable. */
  readonly pageNumber?: number;
  readonly uri?: string;
  readonly children: ReadonlyArray<PdfOutlineNode>;
}

export type PdfAnnotationKind =
  | "highlight"
  | "underline"
  | "strikethrough"
  | "squiggly"
  | "note"
  | "free-text"
  | "ink"
  | "line"
  | "rectangle"
  | "ellipse"
  | "polygon"
  | "polyline"
  | "stamp"
  | "link"
  | "redaction"
  | "unknown";

export interface PdfAnnotation {
  readonly id: NodeId;
  readonly kind: PdfAnnotationKind;
  /** Native PDF subtype (e.g. "Highlight", "Square"). */
  readonly subtype: string;
  readonly pageNumber: number;
  readonly rect: PdfRect;
  readonly contents?: string;
  readonly author?: string;
  readonly color?: { r: number; g: number; b: number; a?: number };
  /** External URI for link annotations. */
  readonly url?: string;
  /** Internal goto-page destination for link annotations. */
  readonly destPage?: number;
  readonly createdAt?: string;
  /** Native PDF object number, if known. Used for incremental save. */
  readonly nativeObjectNumber?: number;
}

export type PdfFormFieldType =
  | "text"
  | "checkbox"
  | "radio"
  | "choice"
  | "button"
  | "signature"
  | "unknown";

export interface PdfFormField {
  readonly id: NodeId;
  readonly name: string;
  readonly type: PdfFormFieldType;
  readonly value?: string | boolean;
  readonly options?: ReadonlyArray<string>;
  readonly readOnly: boolean;
  readonly required: boolean;
  readonly maxLength?: number;
  readonly multiline?: boolean;
  readonly password?: boolean;
  readonly pageNumber: number;
  readonly rect: PdfRect;
}

export interface PdfAttachment {
  readonly id: NodeId;
  readonly name: string;
  readonly bytes: number;
}

export interface PdfComment {
  readonly id: NodeId;
  readonly author: string;
  readonly text: string;
  readonly resolved?: boolean;
  readonly parentId?: NodeId;
  readonly createdAt?: string;
  /** Anchor: page + normalized rect (0..1) for stable position across rotation/zoom. */
  readonly pageNumber: number;
  readonly normalizedRect: PdfRect;
}

export interface PdfDocument {
  readonly metadata: PdfMetadata;
  readonly pages: ReadonlyArray<PdfPage>;
  readonly outline: ReadonlyArray<PdfOutlineNode>;
  readonly annotations: ReadonlyArray<PdfAnnotation>;
  readonly formFields: ReadonlyArray<PdfFormField>;
  readonly attachments: ReadonlyArray<PdfAttachment>;
  readonly comments: ReadonlyArray<PdfComment>;
  /** Number of digital signatures detected (read-only at this time). */
  readonly signatureCount: number;
  /** Engine kind that produced this snapshot's read paths. */
  readonly engineKind: "pdfjs" | "pdfium";
}

/**
 * Headless command-bus snapshot for PDF, mirroring DocxSnapshot /
 * XlsxSnapshot / PptxSnapshot. The original PDF buffer is preserved
 * out-of-band on the agent so incremental save can produce minimal
 * deltas.
 */
export type PdfSnapshot = DocumentSnapshot<PdfDocument> & { readonly format: "pdf" };
