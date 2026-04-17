import type { DocumentSnapshot, NodeId } from "@officeai/core";

/**
 * DOCX in-memory model. See spec/docx/document-model.md.
 *
 * Treated as immutable from the outside; mutations produce new snapshots.
 */

export interface DocxSnapshot extends DocumentSnapshot<DocxDocument> {
  readonly format: "docx";
  /** Internal: the OOXML container. Not serialized; not part of equality. */
  readonly container: import("@officeai/core").ooxml.OoxmlContainer;
  /** Per-part dirty flags consulted by the serializer. */
  readonly dirty: DocxDirtyFlags;
}

export interface DocxDirtyFlags {
  body: boolean;
  comments: boolean;
  rels: boolean;
  contentTypes: boolean;
  /**
   * Whether `word/commentsExtended.xml` (W15 thread + resolved metadata)
   * has been touched and must be re-emitted on save. Independent of
   * `comments`: resolving a comment changes only the extended part, while
   * adding a new comment dirties both `comments` and `commentsExtended`.
   */
  commentsExtended: boolean;
}

export interface DocxDocument {
  readonly id: NodeId;
  readonly body: ReadonlyArray<BlockNode>;
  readonly comments: ReadonlyArray<DocxComment>;
  readonly documentRootAttrs: Readonly<Record<string, string>>;
}

/* ── Block-level ─────────────────────────────────────────────────────────── */

export type BlockNode = Paragraph | Table | SectionBreak | OpaqueBlock;

export interface Paragraph {
  readonly kind: "paragraph";
  readonly id: NodeId;
  readonly properties: ParagraphProperties;
  readonly children: ReadonlyArray<InlineNode>;
}

export interface ParagraphProperties {
  readonly styleId?: string;
  readonly alignment?: "left" | "center" | "right" | "justify";
  readonly indentation?: {
    readonly left?: number;
    readonly right?: number;
    readonly firstLine?: number;
    readonly hanging?: number;
  };
  readonly spacing?: {
    readonly before?: number;
    readonly after?: number;
    readonly line?: number;
    readonly lineRule?: "auto" | "exact" | "atLeast";
  };
  readonly numbering?: { readonly numId: number; readonly ilvl: number };
  /** XML children of <w:pPr> we don't model explicitly. */
  readonly opaqueProps?: ReadonlyArray<OpaqueXml>;
}

export interface Table {
  readonly kind: "table";
  readonly id: NodeId;
  readonly raw: OpaqueXml;
}

export interface SectionBreak {
  readonly kind: "section-break";
  readonly id: NodeId;
  readonly raw: OpaqueXml;
}

export interface OpaqueBlock {
  readonly kind: "opaque-block";
  readonly id: NodeId;
  readonly raw: OpaqueXml;
}

/* ── Inline ──────────────────────────────────────────────────────────────── */

export type InlineNode =
  | Run
  | Hyperlink
  | CommentRangeStart
  | CommentRangeEnd
  | CommentReference
  | RevisionWrapper
  | OpaqueInline;

export interface Run {
  readonly kind: "run";
  readonly id: NodeId;
  readonly properties: RunProperties;
  readonly children: ReadonlyArray<RunChild>;
}

export interface RunProperties {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean | string;
  readonly strike?: boolean;
  readonly fontFamily?: string;
  readonly fontSize?: number;
  readonly color?: string;
  readonly highlight?: string;
  readonly opaqueProps?: ReadonlyArray<OpaqueXml>;
}

export type RunChild = TextLeaf | BreakLeaf | TabLeaf | DrawingLeaf | OpaqueRunChild;

export interface TextLeaf {
  readonly kind: "text";
  readonly id: NodeId;
  readonly text: string;
  readonly xmlSpacePreserve: boolean;
  /** When this leaf came from <w:delText> (inside w:del). */
  readonly isDelText?: boolean;
}

export interface BreakLeaf {
  readonly kind: "break";
  readonly id: NodeId;
  readonly breakType?: "page" | "column" | "textWrapping";
}

export interface TabLeaf {
  readonly kind: "tab";
  readonly id: NodeId;
}

export interface DrawingLeaf {
  readonly kind: "drawing";
  readonly id: NodeId;
  readonly raw: OpaqueXml;
}

export interface OpaqueRunChild {
  readonly kind: "opaque";
  readonly id: NodeId;
  readonly raw: OpaqueXml;
}

export interface Hyperlink {
  readonly kind: "hyperlink";
  readonly id: NodeId;
  readonly relationshipId?: string;
  readonly anchor?: string;
  readonly children: ReadonlyArray<Run>;
}

export interface CommentRangeStart {
  readonly kind: "comment-range-start";
  readonly id: NodeId;
  readonly commentId: string;
}
export interface CommentRangeEnd {
  readonly kind: "comment-range-end";
  readonly id: NodeId;
  readonly commentId: string;
}
export interface CommentReference {
  readonly kind: "comment-reference";
  readonly id: NodeId;
  readonly commentId: string;
}

export interface RevisionWrapper {
  readonly kind: "revision";
  readonly id: NodeId;
  readonly revisionType: "ins" | "del";
  readonly author: string;
  readonly date: string;
  readonly revisionId: string;
  readonly children: ReadonlyArray<InlineNode>;
}

export interface OpaqueInline {
  readonly kind: "opaque-inline";
  readonly id: NodeId;
  readonly raw: OpaqueXml;
}

/* ── Comments part ───────────────────────────────────────────────────────── */

export interface DocxComment {
  readonly id: string;
  readonly author: string;
  readonly initials?: string;
  readonly date: string;
  readonly body: ReadonlyArray<BlockNode>;
  /**
   * Whether the comment thread has been resolved. Driven by
   * `word/commentsExtended.xml` (`w15:commentEx[@w15:done='1']`). When the
   * field is absent in OOXML it is treated as `false` (open).
   */
  readonly resolved?: boolean;
  /**
   * Parent comment id when this is a reply. Drives the `w15:parentPaIdRef`
   * cross-reference in `word/commentsExtended.xml`. Top-level comments leave
   * this undefined.
   */
  readonly parentId?: string;
  /**
   * Stable W14 paragraph id of the comment's first body paragraph. OOXML's
   * `commentsExtended.xml` keys threading and resolved-state by this paraId,
   * not by the comment id. Captured on parse if present; minted on demand
   * by the serializer when a comment needs an extended-metadata entry.
   */
  readonly paraId?: string;
}

/* ── Opaque carrier ──────────────────────────────────────────────────────── */

export interface OpaqueXml {
  readonly tag: string;
  readonly attrs: Readonly<Record<string, string>>;
  /** preserveOrder subtree (the value of entry[tag]). */
  readonly subtree: ReadonlyArray<unknown>;
  /** Original entry attrs map keyed with the @_ prefix. Used by the serializer. */
  readonly rawAttrs: Readonly<Record<string, string>>;
}

/* ── Position / Selection ────────────────────────────────────────────────── */

export interface DocxPosition {
  readonly paragraph: number;
  readonly run?: number;
  readonly offset?: number;
}

export interface DocxSelection {
  readonly start: DocxPosition;
  readonly end: DocxPosition;
}
