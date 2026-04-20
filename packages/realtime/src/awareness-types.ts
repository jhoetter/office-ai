import type { AnonymousIdentity } from "./identity";

/**
 * Per-product cursor / selection envelopes. Kept as a union so the
 * presence layer can render with full type-safety on the consumer
 * side via an exhaustive switch.
 */

/** DOCX caret position, expressed in ProseMirror absolute positions. */
export interface DocxCursor {
  readonly product: "docx";
  /** PM `head` (caret) position. */
  readonly head: number;
  /** PM `anchor` position. Equal to `head` when no range is selected. */
  readonly anchor: number;
}

/** XLSX active selection, mirroring the editor's `Selection` shape. */
export interface XlsxSelection {
  readonly product: "xlsx";
  readonly sheetName: string;
  /** A1 anchor cell, e.g. `"B4"`. */
  readonly anchor: string;
  /** A1 range, e.g. `"B4:D7"`. Equal to `anchor` for a single-cell pick. */
  readonly range: string;
}

/** PPTX active slide + selected shapes. */
export interface PptxSelection {
  readonly product: "pptx";
  readonly slideId: string;
  readonly shapeIds: ReadonlyArray<string>;
}

/** PDF active page + optional pixel-rect selection (normalized 0..1). */
export interface PdfSelection {
  readonly product: "pdf";
  /** 1-indexed page number of the active viewport. */
  readonly pageNumber: number;
  /** Optional normalized rect (0..1) describing what the user has selected on the page. */
  readonly normalizedRect?: readonly [number, number, number, number];
}

/**
 * The full awareness state every peer publishes. The `cursor` field
 * is product-tagged so receivers can ignore peers in a different
 * product (e.g. a DOCX tab and an XLSX tab both joined the same
 * `local-…` room while playing with the demo).
 */
export interface AwarenessState {
  readonly user: AnonymousIdentity;
  readonly product: "docx" | "xlsx" | "pptx" | "pdf";
  readonly cursor?: DocxCursor | XlsxSelection | PptxSelection | PdfSelection;
  /** Wall-clock when this state was last published — tooltip "active 12s ago". */
  readonly lastSeen: number;
}

/**
 * What the UI consumes per remote peer (after filtering out our own
 * client id). Surfaced from the room client's `useAwareness()` hook.
 */
export interface RemotePresence {
  readonly clientId: number;
  readonly state: AwarenessState;
}
