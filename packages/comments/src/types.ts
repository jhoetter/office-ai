/**
 * Format-agnostic comment contract. Each format-specific package
 * (`@officeai/docx`, `@officeai/xlsx`, `@officeai/pptx`) keeps its own
 * native comment model — `DocxComment`, `XlsxThreadedComment`,
 * `PptxComment` — but exposes a `CommentsProvider` adapter that
 * normalises into these canonical shapes so a single React UI can drive
 * all three editors.
 *
 * The canonical comment is intentionally lossy: it carries plain text
 * (no run-level formatting), a single author display name, and a
 * format-agnostic `CommentAnchor`. Round-tripping through the provider
 * preserves the format-specific blob via the `nativeRef` opaque field.
 */

/** A simple plain-text comment body. UI never edits formatted runs. */
export type CommentText = string;

/**
 * Where in the document the comment is anchored. Different formats
 * support different anchor flavours — DOCX has paragraph-level ranges,
 * XLSX has cell coordinates, PPTX has free-floating pins on slides.
 * The provider picks the variant that matches its native model.
 */
export type CommentAnchor =
  | {
      readonly kind: "docx-range";
      /** First paragraph the comment range covers; used for sidebar grouping. */
      readonly paragraphIndex: number;
      /**
       * Opaque format-specific range descriptor. The DOCX adapter sets
       * this to a `DocxSelection` so the `add` command has the full
       * `(start, end)` it needs; the shared UI never inspects it.
       */
      readonly range?: unknown;
    }
  | {
      readonly kind: "xlsx-cell";
      /** Sheet display name (matches `Sheet.name`). */
      readonly sheet: string;
      /** A1 single-cell ref, e.g. `"B7"`. */
      readonly ref: string;
    }
  | {
      readonly kind: "pptx-pin";
      readonly slideIndex: number;
      /** EMU coordinates for the pin. */
      readonly xEmu: number;
      readonly yEmu: number;
      /** Optional anchored shape id; if set, the pin moves with the shape. */
      readonly shapeId?: string;
    }
  | {
      /**
       * PDF region anchor. Coordinates are NORMALIZED to the PDF page's
       * MediaBox (0..1) so the comment stays correctly placed under
       * page rotation, zoom, and reflow. The optional `nativeAnnotId`
       * is the PDF object number when the comment is also written into
       * the document as a native /Text annotation; absent for adapter-
       * only comments that live in our snapshot only.
       */
      readonly kind: "pdf-region";
      /** 1-indexed page number, matching the PDF spec convention. */
      readonly pageNumber: number;
      /** Normalised rect [x1, y1, x2, y2] in 0..1 of the page MediaBox. */
      readonly normalizedRect: readonly [number, number, number, number];
      readonly nativeAnnotId?: string;
    }
  | { readonly kind: "none" };

/** Normalised comment as the shared UI sees it. */
export interface CommentBody {
  readonly id: string;
  readonly author: string;
  readonly text: CommentText;
  readonly createdAt?: string;
  readonly resolved?: boolean;
  readonly parentId?: string;
  readonly anchor: CommentAnchor;
  /**
   * Opaque reference back to the format-specific comment object. The
   * shared UI doesn't read this — it's a debugging aid for adapters and
   * for tests that want to assert round-trip identity.
   */
  readonly nativeRef?: unknown;
}

/** A top-level comment plus its chronologically-ordered replies. */
export interface CommentThread {
  readonly parent: CommentBody;
  readonly replies: ReadonlyArray<CommentBody>;
}

/**
 * Provider-driven contract: every editor gives the shared UI a
 * `CommentsProvider` it can call `add/reply/resolve/delete` against.
 * Each method returns a Promise so the adapter can debounce, batch, or
 * just dispatch a CommandBus command.
 */
export interface CommentsProvider {
  /** All comments grouped into threads, in stable display order. */
  threads(): ReadonlyArray<CommentThread>;
  /** Add a new top-level comment at `anchor`. */
  add(input: { author: string; text: string; anchor: CommentAnchor }): Promise<string>;
  /** Reply to an existing thread. Returns the new comment's id. */
  reply(input: { parentId: string; author: string; text: string }): Promise<string>;
  /** Toggle the resolved state of a thread parent. */
  resolve(commentId: string, resolved: boolean): Promise<void>;
  /** Delete a comment (or a reply). Adapters decide the cascade. */
  delete(commentId: string): Promise<void>;
  /** Optional edit hook — adapters that don't support edits can omit. */
  edit?(commentId: string, text: string): Promise<void>;
  /**
   * Optional callback fired when a thread is selected in the UI; the
   * adapter typically uses this to scroll its native canvas to the
   * anchor and flash a highlight.
   */
  onScrollTo?(commentId: string): void;
}
