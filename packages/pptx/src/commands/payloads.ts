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
  /**
   * Plain-text replacement. Splits on `\n` into paragraphs, each
   * collapsed to a single run that inherits the first existing
   * paragraph/run's properties. This was the original v0 contract;
   * prefer `paragraphs` for non-trivial commits so multi-run
   * formatting (bold/italic/colour spans within a paragraph) isn't
   * silently flattened.
   */
  readonly text?: string;
  /**
   * Structured replacement. When present, takes precedence over
   * `text` and lets callers preserve the existing per-run formatting
   * structure (bold spans, mixed colours, line breaks). Each paragraph
   * is replaced wholesale; runs inherit per-run properties when
   * provided, otherwise inherit from the corresponding existing
   * shape paragraph/run when available.
   */
  readonly paragraphs?: ReadonlyArray<SetTextParagraphPatch>;
}

export interface SetTextParagraphPatch {
  /** Paragraph properties; if omitted, inherits from the same-index existing paragraph. */
  readonly properties?: import("../model/types.js").TextParagraphProperties;
  readonly runs: ReadonlyArray<SetTextRunPatch>;
}

export interface SetTextRunPatch {
  readonly text: string;
  /**
   * Run properties; if omitted, inherits from the run referenced by
   * `inheritFromRun` (when present and valid), otherwise from the
   * first run of the same-index existing paragraph, otherwise from
   * the shape's first run.
   */
  readonly properties?: import("../model/types.js").TextRunProperties;
  readonly isLineBreak?: boolean;
  /**
   * Index hint pointing at the original run this slice originated
   * from, so the command handler can copy its full property bag
   * (including opaque XML) without losing fidelity.
   */
  readonly inheritFromRun?: number;
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

/**
 * Replace the bitmap behind an existing `Picture` with a new file
 * while preserving its position, size, alt-text and `spPrTail` (so any
 * border / shadow / corner-radius styling stays intact). Mints a new
 * media part if the bytes don't match an existing one (sha-256 dedup),
 * adds a slide-rels entry if needed, and updates the picture's
 * `mediaRelId`/`mediaPartPath`.
 */
export interface ReplacePictureMediaPayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly data: Uint8Array | ArrayBuffer;
  readonly mimeType: string;
  /** Optional new alt text — leave undefined to preserve the existing descr. */
  readonly altText?: string;
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

// ─── Text alignment / anchor ──────────────────────────────────────────────

/**
 * Mirrors PowerPoint's "Align Text" Left/Center/Right/Justify in the
 * Home ribbon. Per-paragraph; when `paragraphs` is omitted the change
 * is applied to every paragraph in the shape (matching the behaviour
 * the user gets by clicking "Align" with the shape selected but no
 * text-edit caret open). Pass `alignment: null` to clear an existing
 * `<a:pPr algn>` so the paragraph re-inherits from its style chain.
 */
export interface SetParagraphAlignmentPayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly alignment: "left" | "center" | "right" | "justify" | null;
  /**
   * Optional 0-based paragraph indices to target. Omit to apply the
   * change to every paragraph in the shape.
   */
  readonly paragraphs?: ReadonlyArray<number>;
}

export type TextAnchor = "top" | "middle" | "bottom";

/**
 * Mirrors PowerPoint's "Align Text" Top/Middle/Bottom in the Home
 * ribbon — the shape-wide vertical anchor that lives on
 * `<a:bodyPr anchor="t|ctr|b">`. Pass `anchor: null` to clear an
 * existing override so the shape re-inherits the layout/master
 * default (PowerPoint treats no attribute as "top").
 */
export interface SetTextAnchorPayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly anchor: TextAnchor | null;
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

export type ReorderShapeMode = "to-front" | "to-back" | "forward" | "backward";

export interface ReorderShapePayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly mode: ReorderShapeMode;
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
  /**
   * Reference frame for the alignment math.
   *   "selection" (default) — every shape snaps to the union box of
   *     the selection; needs ≥ 2 alignable shapes. Mirrors
   *     PowerPoint's "Align Selected Objects".
   *   "slide"               — every shape snaps to the slide bounds
   *     (`<p:sldSz>`); works with a single selected shape too. Mirrors
   *     PowerPoint's "Align to Slide".
   */
  readonly relativeTo?: "selection" | "slide";
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

/**
 * Group two-or-more top-level shapes into a `GroupShape`. The group's
 * bounding box is the union of the children's positions/sizes; the
 * children are removed from the slide and re-inserted as the group's
 * `children`. Connectors anchored to the regrouped shapes keep working
 * because they reference shapes by `cNvPrId`, not by tree path.
 *
 * Restrictions:
 *   • All shape ids must resolve to top-level shapes (no nested-group
 *     children — PowerPoint allows this but we keep it simple for now).
 *   • Every shape must carry an explicit `position` and `size`. Implicit
 *     placeholders are refused so the synthesised group has a meaningful
 *     bounding box.
 *   • At least 2 shape ids are required.
 */
export interface GroupShapesPayload {
  readonly slideIndex: number;
  readonly shapeIds: ReadonlyArray<NodeId>;
  readonly name?: string;
}

/**
 * Dissolve a `GroupShape`, re-inserting its children at the group's
 * position in the slide's shape array. Children keep their original
 * absolute positions (the group's `chOff` was synthesised on `group` to
 * equal its `position`, so this is a no-op coordinate-wise).
 */
export interface UngroupShapePayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
}

/**
 * Duplicate a single shape on a slide. The clone is appended to the
 * top-level shape array (it ends up in front of every other shape) and
 * gets a unique `cNvPrId` so connectors and animations don't get
 * confused. Optional `dxEmu`/`dyEmu` nudge the clone's position so it
 * doesn't perfectly overlap the source. Connector targets are NOT
 * rewritten — duplicated connectors keep referencing the original
 * source/end shapes, matching PowerPoint's `Cmd+D` behaviour.
 */
export interface DuplicateShapePayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  /** Default 228_600 (¼ inch) — same nudge as PowerPoint's `Cmd+D`. */
  readonly dxEmu?: number;
  /** Default 228_600 (¼ inch). */
  readonly dyEmu?: number;
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
      /**
       * Optional position along the side, in [0, 1]. Defaults to 0.5
       * (the cardinal midpoint) when omitted. Ignored for "center".
       */
      readonly t?: number;
    };

export type ConnectorTypePayload = "straight" | "elbow" | "curved";

export type ConnectorEndShapePayload = "none" | "arrow" | "triangle" | "oval";

export type ConnectorDashStylePayload = "solid" | "dashed" | "dotted" | "longDash" | "dashDot";

export interface AddConnectorPayload {
  readonly slideIndex: number;
  readonly connectorType: ConnectorTypePayload;
  readonly start: ConnectorEndpointPayload;
  readonly end: ConnectorEndpointPayload;
  /** 6-char hex (with or without leading `#`). Defaults to `374151` (slate-700). */
  readonly strokeColor?: string;
  /** EMU width. Defaults to ~0.75pt (`9525`). */
  readonly strokeWidthEmu?: number;
  readonly strokeDash?: ConnectorDashStylePayload;
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
  readonly strokeDash?: ConnectorDashStylePayload;
  readonly headEnd?: ConnectorEndShapePayload;
  readonly tailEnd?: ConnectorEndShapePayload;
}

/**
 * Adjust the perpendicular offset of one of an elbow connector's
 * routed segments. Pass `valueEmu = null` to clear an existing
 * waypoint (the segment will revert to the auto-routed midpoint).
 */
export interface SetConnectorWaypointPayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
  readonly segmentIndex: number;
  readonly valueEmu: number | null;
}

/**
 * Drop all user-supplied waypoints and let the auto-router pick the
 * polyline from scratch. Useful after a layout change makes a
 * previously-tweaked elbow look strange. No-op for non-elbow types.
 */
export interface RerouteConnectorPayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
}

/**
 * Reverse a connector's direction: start ↔ end and head ↔ tail. The
 * route reverses with the endpoints, so a connector that pointed
 * left-to-right now points right-to-left without any other change to
 * the model.
 */
export interface SwapConnectorDirectionPayload {
  readonly slideIndex: number;
  readonly shapeId: NodeId;
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
  setParagraphAlignment: "pptx:set-paragraph-alignment",
  setTextAnchor: "pptx:set-text-anchor",
  insertImage: "pptx:insert-image",
  replacePictureMedia: "pptx:replace-picture-media",
  addTextBox: "pptx:add-text-box",
  addShape: "pptx:add-shape",
  deleteShape: "pptx:delete-shape",
  setShapeFill: "pptx:set-shape-fill",
  reorderShape: "pptx:reorder-shape",
  duplicateShape: "pptx:duplicate-shape",
  groupShapes: "pptx:group-shapes",
  ungroupShape: "pptx:ungroup-shape",
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
  setConnectorWaypoint: "pptx:set-connector-waypoint",
  rerouteConnector: "pptx:reroute-connector",
  swapConnectorDirection: "pptx:swap-connector-direction",
  setSlideLayout: "pptx:set-slide-layout",
  setSlideNotes: "pptx:set-slide-notes",
  addComment: "pptx:add-comment",
  replyComment: "pptx:reply-comment",
  resolveComment: "pptx:resolve-comment",
  deleteComment: "pptx:delete-comment",
  editComment: "pptx:edit-comment",
} as const;

export type PptxCommandType = (typeof PPTX_COMMAND_TYPES)[keyof typeof PPTX_COMMAND_TYPES];
