import type { NodeId } from "@officeai/core";

// ─── Common ───────────────────────────────────────────────────────────────

export interface PptxTextRange {
  readonly paragraph: number;
  readonly start: number;
  readonly end: number;
}

export interface TextFormatPayload {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean | string;
  readonly strike?: boolean;
  readonly fontFamily?: string;
  readonly fontSizeHundredths?: number;
  readonly color?: string;
  /**
   * Character highlight (background colour behind glyphs). 6-character
   * RRGGBB hex (no `#`). Pass the empty string to clear the highlight
   * on the matched range; `undefined` leaves the existing highlight
   * untouched (patch semantics).
   */
  readonly highlight?: string;
}

// ─── P0 payloads ──────────────────────────────────────────────────────────

/**
 * The 11 PowerPoint-standard layout kinds the agent / picker can request.
 * `unknown` is intentionally excluded — agents shouldn't manufacture
 * unclassified layouts.
 */
export type LayoutKindPayload =
  | "title"
  | "titleAndContent"
  | "sectionHeader"
  | "twoContent"
  | "comparison"
  | "titleOnly"
  | "blank"
  | "contentWithCaption"
  | "pictureWithCaption"
  | "titleSlide"
  | "bigNumber";

export interface AddSlidePayload {
  readonly at?: number;
  readonly layoutPartPath?: string;
  /**
   * Pick a built-in layout kind by classification rather than partPath.
   * Resolved against the deck's existing layouts; if the deck doesn't
   * have one of that kind we synthesise a minimal layout part on demand.
   */
  readonly layoutKind?: LayoutKindPayload;
  /** Stamp the layout's placeholders as concrete shapes on the new slide. */
  readonly clonePlaceholders?: boolean;
}

// ─── Comments ─────────────────────────────────────────────────────────────

export interface AddCommentPayload {
  readonly slideIndex: number;
  readonly author: string;
  readonly text: string;
  /** Pin position on the slide, in EMU. Defaults to centre when omitted. */
  readonly xEmu?: number;
  readonly yEmu?: number;
  /**
   * Optional anchored shape id. When set, the comment is pinned at
   * `(xEmu, yEmu)` *and* tagged with the shape id; the editor uses
   * this to paint a yellow indicator over the anchored shape until
   * the thread is resolved. Persisted via the
   * `officeai:shapeAnchor` extension on `<p:cm>` so it round-trips.
   */
  readonly shapeId?: string;
}

export interface ReplyCommentPayload {
  readonly slideIndex: number;
  /** Top-level comment id this reply belongs to (`${authorId}:${idx}`). */
  readonly parentId: string;
  readonly author: string;
  readonly text: string;
}

export interface ResolveCommentPayload {
  readonly slideIndex: number;
  readonly commentId: string;
  readonly resolved: boolean;
}

export interface DeleteCommentPayload {
  readonly slideIndex: number;
  readonly commentId: string;
}

export interface EditCommentPayload {
  readonly slideIndex: number;
  readonly commentId: string;
  readonly text: string;
}

export interface SetSlideNotesPayload {
  readonly slideIndex: number;
  /**
   * New plain-text body for the notes part. Each line becomes one
   * paragraph (`<a:p>`); existing run-level formatting is dropped — the
   * notes editor is intentionally plain-text for now. Pass an empty
   * string to clear the notes (the part stays around so we don't churn
   * relationships on every keystroke).
   */
  readonly text: string;
}

export interface SetSlideLayoutPayload {
  readonly slideIndex: number;
  readonly layoutPartPath?: string;
  readonly layoutKind?: LayoutKindPayload;
  /** Re-stamp placeholders, preserving content where idx matches. Defaults to true. */
  readonly clonePlaceholders?: boolean;
}

export interface DeleteSlidePayload {
  readonly slideIndex: number;
}

export interface DuplicateSlidePayload {
  readonly slideIndex: number;
}

export interface MoveSlidePayload {
  readonly from: number;
  readonly to: number;
}

export interface SetTextPayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly text: string;
}

export interface SetPositionPayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly x: number;
  readonly y: number;
}

export interface SetSizePayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly width: number;
  readonly height: number;
}

// ─── P1 payloads ──────────────────────────────────────────────────────────

export interface FormatTextPayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly range: PptxTextRange;
  readonly format: TextFormatPayload;
}

export interface InsertImagePayload {
  readonly slideIndex: number;
  readonly data: Uint8Array | ArrayBuffer;
  readonly mimeType: string;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly altText?: string;
  readonly name?: string;
}

export interface AddTextBoxPayload {
  readonly slideIndex: number;
  readonly text: string;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  readonly name?: string;
}

// ─── Shape primitives (P2-lite) ───────────────────────────────────────────

/**
 * Subset of `<a:prstGeom prst="…">` values that the editor surfaces in its
 * "Insert shape" menu. Anything else can still be parsed and round-tripped
 * verbatim — the renderer just won't draw a typed glyph for it. Sticking
 * to this short list keeps the shape picker UI manageable.
 */
export type ShapePreset =
  | "rect"
  | "roundRect"
  | "ellipse"
  | "triangle"
  | "rtTriangle"
  | "diamond"
  | "line"
  | "rightArrow";

export interface AddShapePayload {
  readonly slideIndex: number;
  readonly preset: ShapePreset;
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  /** 6-char hex (with or without leading `#`). Defaults to a friendly accent. */
  readonly fill?: string;
  /** 6-char hex; if omitted the shape draws without an explicit outline. */
  readonly stroke?: string;
  readonly strokeWidthEmu?: number;
  /** Optional initial label rendered inside the shape. */
  readonly text?: string;
  readonly name?: string;
}

export interface DeleteShapePayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
}

export interface SetShapeFillPayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  /** 6-char hex; pass `null` to clear the fill (renders as transparent). */
  readonly fill: string | null;
}

// ─── Alignment / distribution (multi-select operations) ───────────────────

/**
 * Mirrors PowerPoint's "Align" submenu. `left/right/center-h` operate on
 * x-axis; `top/bottom/middle-v` operate on y-axis. Each shape is moved
 * so its corresponding edge or center matches the union bounding box of
 * the selection.
 */
export type AlignMode = "left" | "center-h" | "right" | "top" | "middle-v" | "bottom";

export interface AlignShapesPayload {
  readonly slideIndex: number;
  readonly shapeIds: ReadonlyArray<NodeId>;
  readonly mode: AlignMode;
}

/**
 * Mirrors PowerPoint's "Distribute Horizontally / Vertically". The two
 * extreme shapes stay anchored; intermediate shapes get re-positioned so
 * their *centres* are equidistant along the chosen axis.
 */
export interface DistributeShapesPayload {
  readonly slideIndex: number;
  readonly shapeIds: ReadonlyArray<NodeId>;
  readonly axis: "horizontal" | "vertical";
}

// ─── F2 (Tables) payloads ─────────────────────────────────────────────────

export interface TableSetCellTextPayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly row: number;
  readonly column: number;
  readonly text: string;
}

export interface TableAddRowPayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly at?: number;
  readonly height?: number;
}

export interface TableDeleteRowPayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly row: number;
}

export interface TableAddColumnPayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly at?: number;
  readonly width?: number;
}

export interface TableDeleteColumnPayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly column: number;
}

// ─── F3 (Charts) payloads ─────────────────────────────────────────────────

export interface SetChartTitlePayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  /** New title; pass `null` to remove the title. */
  readonly title: string | null;
}

export interface SetChartDataPayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly categories: ReadonlyArray<string>;
  readonly series: ReadonlyArray<{
    readonly name?: string;
    readonly values: ReadonlyArray<number>;
  }>;
}

export interface SetChartTypePayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly chartType: "bar" | "line" | "pie" | "area";
}

// ─── Connectors (lines) payloads ──────────────────────────────────────────

/**
 * One end of a connector. Mirrors the model's `ConnectorEndpoint` but
 * uses `cNvPrId` only — the agent never deals in `NodeId`s for the
 * target shape because connectors must survive shape re-mounting (which
 * mints fresh NodeIds). `cNvPrId` is stable across the document
 * lifetime.
 *
 * `side` is one of the five named anchor sides exposed by the renderer
 * (`n`/`s`/`e`/`w`/`center`); the canvas's snap-to-anchor result can be
 * passed straight through.
 */
export type ConnectorEndpointPayload =
  | { readonly kind: "free"; readonly xEmu: number; readonly yEmu: number }
  | {
      readonly kind: "anchored";
      readonly targetCNvPrId: number;
      readonly side: "n" | "s" | "e" | "w" | "center";
    };

export type ConnectorTypePayload = "straight" | "elbow" | "curved";

export type ConnectorEndShapePayload = "none" | "arrow" | "triangle" | "oval";

export interface AddConnectorPayload {
  readonly slideIndex: number;
  readonly connectorType: ConnectorTypePayload;
  readonly start: ConnectorEndpointPayload;
  readonly end: ConnectorEndpointPayload;
  /** 6-char hex (with or without leading `#`). Defaults to `374151` (slate-700). */
  readonly strokeColor?: string;
  /** EMU width. Defaults to ~0.75pt (`9525`). */
  readonly strokeWidthEmu?: number;
  readonly headEnd?: ConnectorEndShapePayload;
  readonly tailEnd?: ConnectorEndShapePayload;
  readonly name?: string;
}

export interface SetConnectorEndpointPayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly which: "start" | "end";
  readonly endpoint: ConnectorEndpointPayload;
}

export interface SetConnectorStylePayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly connectorType?: ConnectorTypePayload;
  readonly strokeColor?: string;
  readonly strokeWidthEmu?: number;
  readonly headEnd?: ConnectorEndShapePayload;
  readonly tailEnd?: ConnectorEndShapePayload;
}

// ─── F4 (Animations) payloads ─────────────────────────────────────────────

export interface SetSlideTransitionPayload {
  readonly slideIndex: number;
  /** Pass `"none"` to remove an existing transition. */
  readonly kind: "none" | "fade" | "push" | "wipe" | "split" | "cut";
  readonly speed?: "slow" | "med" | "fast";
}

export interface AddShapeAnimationPayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly effect: "appear" | "fade" | "fly-in" | "wipe";
  /** Insert position in the main entrance sequence. Defaults to append. */
  readonly at?: number;
  readonly durationMs?: number;
}

export interface RemoveShapeAnimationPayload {
  readonly slideIndex: number;
  readonly animationId: NodeId;
}

export interface ReorderShapeAnimationsPayload {
  readonly slideIndex: number;
  /** New order, must be a permutation of the slide's current animations[].id. */
  readonly order: ReadonlyArray<NodeId>;
}

// ─── Type tags ────────────────────────────────────────────────────────────

export const PPTX_COMMAND_TYPES = {
  addSlide: "pptx:add-slide",
  deleteSlide: "pptx:delete-slide",
  duplicateSlide: "pptx:duplicate-slide",
  moveSlide: "pptx:move-slide",
  setText: "pptx:set-text",
  setPosition: "pptx:set-position",
  setSize: "pptx:set-size",
  formatText: "pptx:format-text",
  insertImage: "pptx:insert-image",
  addTextBox: "pptx:add-text-box",
  addShape: "pptx:add-shape",
  deleteShape: "pptx:delete-shape",
  setShapeFill: "pptx:set-shape-fill",
  alignShapes: "pptx:align-shapes",
  distributeShapes: "pptx:distribute-shapes",
  tableSetCellText: "pptx:table-set-cell-text",
  tableAddRow: "pptx:table-add-row",
  tableDeleteRow: "pptx:table-delete-row",
  tableAddColumn: "pptx:table-add-column",
  tableDeleteColumn: "pptx:table-delete-column",
  setChartTitle: "pptx:set-chart-title",
  setChartData: "pptx:set-chart-data",
  setChartType: "pptx:set-chart-type",
  setSlideTransition: "pptx:set-slide-transition",
  addShapeAnimation: "pptx:add-shape-animation",
  removeShapeAnimation: "pptx:remove-shape-animation",
  reorderShapeAnimations: "pptx:reorder-shape-animations",
  addConnector: "pptx:add-connector",
  setConnectorEndpoint: "pptx:set-connector-endpoint",
  setConnectorStyle: "pptx:set-connector-style",
  setSlideLayout: "pptx:set-slide-layout",
  setSlideNotes: "pptx:set-slide-notes",
  addComment: "pptx:add-comment",
  replyComment: "pptx:reply-comment",
  resolveComment: "pptx:resolve-comment",
  deleteComment: "pptx:delete-comment",
  editComment: "pptx:edit-comment",
} as const;

export type PptxCommandType = (typeof PPTX_COMMAND_TYPES)[keyof typeof PPTX_COMMAND_TYPES];
