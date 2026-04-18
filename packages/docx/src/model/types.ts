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
  /**
   * Set of header/footer part paths (e.g. `"word/header1.xml"`) that have
   * been mutated since load. Untouched parts are re-emitted from cached
   * bytes; touched parts are re-serialized from `headersAndFooters`. The
   * set is the empty `Set` for a freshly-parsed snapshot, which preserves
   * the byte-equality invariant on round-trip.
   */
  headersAndFooters: ReadonlySet<string>;
  /**
   * Set of media part paths (e.g. `"word/media/image1.png"`) that have
   * been added or changed since load. Untouched media parts are
   * preserved verbatim by the cloned container — the serializer only
   * writes back parts in this set. Added in P1.3 / W8 (image insertion).
   */
  media: ReadonlySet<string>;
  /**
   * Set of relationships-part paths (e.g.
   * `"word/_rels/document.xml.rels"`) that have been modified since load.
   * Untouched rels parts are left alone in the cloned container so they
   * round-trip byte-identical. Added in P1.3 / W8.
   */
  relationships: ReadonlySet<string>;
  /**
   * Whether `word/numbering.xml` itself has been mutated and must be
   * re-emitted on save. Added in P1.4 / W10. The flag is intentionally
   * scoped to changes in the typed `NumberingDefinitions` carrier
   * (`DocxDocument.numbering`); a paragraph swapping its `numId` /
   * `ilvl` lives in `word/document.xml` and only dirties `body`.
   */
  numbering: boolean;
}

export interface DocxDocument {
  readonly id: NodeId;
  readonly body: ReadonlyArray<BlockNode>;
  readonly comments: ReadonlyArray<DocxComment>;
  /**
   * Header / footer parts discovered via `word/_rels/document.xml.rels`. The
   * order matches load order (relationship order). Untouched parts may be
   * re-emitted from the original `Uint8Array` cache; touched parts are
   * re-serialized from this typed model. See `parser/headers-footers.ts`
   * and `serializer/headers-footers.ts`.
   */
  readonly headersAndFooters: ReadonlyArray<HeaderFooterPart>;
  /**
   * Binary media parts (`word/media/*.{png,jpg,...}`) keyed by part path.
   * Populated on load by `parseMediaParts`. New entries added by
   * `docx:insert-image`. Untouched parts round-trip byte-identical via
   * the container's part cache; only entries in `dirty.media` are
   * re-written by the serializer. Added in P1.3 / W8.
   */
  readonly media: ReadonlyMap<string, MediaPart>;
  /**
   * Relationships parts (`*.rels`) parsed into a typed map keyed by the
   * **owning part path** (NOT the rels part path itself). For example,
   * the rels for `word/document.xml` live at
   * `word/_rels/document.xml.rels` but are keyed here by
   * `"word/document.xml"`. The package-level rels (`_rels/.rels`) are
   * keyed by the empty string `""`. Added in P1.3 / W8.
   *
   * Mutating commands that touch rels (currently only
   * `docx:insert-image`) replace the relevant array and add the owning
   * part path to `dirty.relationships`. Untouched parts are not
   * re-emitted, so their bytes (and attribute order) survive the
   * round-trip exactly.
   */
  readonly relationships: ReadonlyMap<string, ReadonlyArray<Relationship>>;
  /**
   * Typed numbering definitions parsed from `word/numbering.xml`.
   * `undefined` when the part is absent (which is the common case — a
   * doc that contains no list paragraphs has no `numbering.xml` at
   * all). Added in P1.4 / W10.
   */
  readonly numbering?: NumberingDefinitions;
  readonly documentRootAttrs: Readonly<Record<string, string>>;
}

/**
 * Typed projection of `word/numbering.xml`. Mirrors the OOXML
 * `<w:numbering>` shape: a flat collection of abstract numbering
 * definitions (the templates, keyed by string id) plus a flat
 * collection of concrete `<w:num>` instances (keyed by integer numId)
 * pointing at one of those abstracts.
 *
 * Untouched documents ride the container's part cache; we only emit
 * `numbering.xml` from this typed carrier when `dirty.numbering` is
 * set. Unknown children of `<w:abstractNum>` and `<w:num>` are
 * captured as `OpaqueXml` so the round-trip stays lossless when (a
 * future workstream) does mutate the part. Added in P1.4 / W10.
 */
export interface NumberingDefinitions {
  readonly abstractNums: ReadonlyMap<string, AbstractNum>;
  readonly nums: ReadonlyMap<number, NumInstance>;
}

export interface AbstractNum {
  readonly id: string;
  readonly multiLevelType?: "singleLevel" | "multilevel" | "hybridMultilevel";
  readonly levels: ReadonlyArray<NumberingLevel>;
  readonly raw?: OpaqueXml;
}

export interface NumberingLevel {
  readonly ilvl: number;
  readonly numFmt?: string;
  readonly lvlText?: string;
  readonly start?: number;
  readonly pPr?: ParagraphProperties;
  readonly rPr?: RunProperties;
  readonly opaqueProps?: ReadonlyArray<OpaqueXml>;
}

export interface NumInstance {
  readonly numId: number;
  readonly abstractNumId: string;
  readonly lvlOverrides?: ReadonlyArray<{ readonly ilvl: number; readonly startOverride?: number }>;
  readonly raw?: OpaqueXml;
}

/**
 * A binary media part stored under `word/media/`. The bytes are kept as a
 * `Uint8Array` (no `Buffer` — packages/docx must work in browser bundles
 * too). `digest` is the hex-encoded SHA-256 of `bytes`, computed once on
 * load, and used for de-duplication: two `docx:insert-image` calls with
 * identical bytes share the same media part.
 */
export interface MediaPart {
  readonly partPath: string;
  readonly mimeType: string;
  readonly bytes: Uint8Array;
  readonly digest: string;
}

/**
 * Typed `*.rels` entry. Mirrors `@officeai/core` `Relationship`. We
 * re-export the shape here so consumers of the docx model don't have to
 * reach across to the OPC package layer for a simple type. Added in
 * P1.3 / W8.
 */
export interface Relationship {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly targetMode?: "Internal" | "External";
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

/**
 * Typed `<w:tbl>` carrier (P1.3 / W7).
 *
 * Byte-preservation contract: when a table has not been mutated since load it
 * carries the original `<w:tbl>` subtree in `raw`, and the serializer emits
 * those bytes verbatim. Any mutation that produces a new `Table` MUST drop
 * `raw` (set it to `undefined`); the serializer treats the absence of `raw`
 * as "regenerate from typed fields". This per-table marker is what lets a
 * `set-cell-content` against table A leave table B byte-identical even
 * though `dirty.body` is set on the snapshot.
 */
export interface Table {
  readonly kind: "table";
  readonly id: NodeId;
  readonly properties: TableProperties;
  readonly grid: ReadonlyArray<TableGridCol>;
  readonly rows: ReadonlyArray<TableRow>;
  /**
   * Original `<w:tbl>` subtree captured at parse time. Present on freshly
   * parsed tables and on tables that have not been touched by a mutating
   * command. MUST be omitted (or `undefined`) on any `Table` produced by a
   * mutating command — the serializer uses its presence as the per-table
   * "clean / re-emit cached bytes" signal.
   */
  readonly raw?: OpaqueXml;
}

export interface TableRow {
  readonly kind: "table-row";
  readonly id: NodeId;
  readonly properties: TableRowProperties;
  readonly cells: ReadonlyArray<TableCell>;
}

export interface TableCell {
  readonly kind: "table-cell";
  readonly id: NodeId;
  readonly properties: TableCellProperties;
  /** Block-level cell body: paragraphs and (recursively) nested tables. */
  readonly body: ReadonlyArray<BlockNode>;
}

/** A `<w:gridCol>` entry in `<w:tblGrid>`. */
export interface TableGridCol {
  /** Column width in twips (`w:w`). Optional — some tables omit it. */
  readonly w?: number;
}

export interface TableWidth {
  readonly value: number;
  readonly type: "auto" | "dxa" | "pct" | "nil";
}

/**
 * Modeled subset of `<w:tblPr>`. The fields we don't model (borders, shading,
 * styles, look flags, …) are preserved verbatim in `opaqueProps[]` so the
 * round-trip stays lossless. Order is the OOXML canonical order on emit
 * (`tblPr` is order-sensitive in the schema; we keep typed fields first then
 * append opaque children in their original document order).
 */
export interface TableProperties {
  readonly width?: TableWidth;
  readonly jc?: "left" | "center" | "right" | "start" | "end";
  readonly opaqueProps?: ReadonlyArray<OpaqueXml>;
}

/** Modeled subset of `<w:trPr>`. */
export interface TableRowProperties {
  /** `<w:trHeight w:val="…" w:hRule="…"/>`. */
  readonly trHeight?: { readonly value: number; readonly rule?: "auto" | "exact" | "atLeast" };
  /** `<w:tblHeader/>` — when true, the row repeats as a header on each page. */
  readonly header?: boolean;
  readonly opaqueProps?: ReadonlyArray<OpaqueXml>;
}

/** Modeled subset of `<w:tcPr>`. */
export interface TableCellProperties {
  /** `<w:gridSpan w:val="N"/>` — horizontal merge. Defaults to 1. */
  readonly gridSpan?: number;
  /** `<w:vMerge>` — vertical merge marker. `"restart"` opens a region; `"continue"` continues it. */
  readonly vMerge?: "restart" | "continue";
  /** `<w:tcW>` cell width. */
  readonly tcW?: TableWidth;
  readonly opaqueProps?: ReadonlyArray<OpaqueXml>;
}

export interface SectionBreak {
  readonly kind: "section-break";
  readonly id: NodeId;
  readonly raw: OpaqueXml;
}

/**
 * A `<w:sdt>` / `<w:fldSimple>` / `<mc:AlternateContent>` / `<w:smartTag>` /
 * `<w:customXml>` wrapper at the body level. The wrapper itself is preserved
 * verbatim through `raw` (so byte-identical round-trip is unaffected), but
 * its inner content is **also** parsed into typed `children` so the
 * renderer can surface the underlying paragraphs as real
 * `<h1>`/`<p>` nodes instead of a single italic preview chip.
 *
 * Dirty-tracking contract (P2.3 / W15):
 *
 *   - `subtreeDirty === false` (default) → serializer re-emits `raw`
 *     verbatim. `children` is purely a render-side projection.
 *   - `subtreeDirty === true` → serializer reconstructs the wrapper by
 *     splicing serialized `children` into the wrapper's content slot
 *     (e.g. `<w:sdtContent>`). Mutations that touch a child of an
 *     opaque carrier MUST flip this flag and clear `raw` derivatives.
 *
 * No mutating command currently writes through an opaque carrier, so
 * `subtreeDirty` is always `false` in this iteration. The flag and the
 * dirty serializer path are introduced now so that a future "edit
 * inside an SDT" mutation can flip them without changing the carrier
 * shape again.
 */
export interface OpaqueBlock {
  readonly kind: "opaque-block";
  readonly id: NodeId;
  readonly raw: OpaqueXml;
  readonly children?: ReadonlyArray<BlockNode>;
  readonly subtreeDirty?: boolean;
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

/**
 * A `<w:drawing>` leaf inside a run. Discriminated union over `subkind`:
 *
 * - `"inline-image"` — a fully typed `<wp:inline>` containing
 *   `<a:graphic><a:graphicData uri=".../picture"><pic:pic>`. Captures the
 *   relationship id, EMU dimensions, and `<wp:docPr>` metadata so
 *   mutation commands (and the AI's projection layer) can reason about
 *   them without parsing the subtree. The original subtree is cached in
 *   `raw` so the serializer can re-emit byte-identical bytes when no
 *   field has changed.
 * - `"opaque"` — every other drawing (charts, shapes, SmartArt, anchored
 *   floats with text-wrap, embedded objects, …). Stored as opaque XML
 *   and round-tripped verbatim. The typed promotion lands one drawing
 *   class at a time; this is the catch-all bucket.
 */
export type DrawingLeaf = InlineImageDrawing | OpaqueDrawing;

export interface InlineImageDrawing {
  readonly kind: "drawing";
  readonly subkind: "inline-image";
  readonly id: NodeId;
  /** `r:embed` relationship id resolving to a `word/media/*` part. */
  readonly relId: string;
  /** Display width in OOXML EMUs (`<wp:extent cx>`). 9525 EMU = 1px @ 96 DPI. */
  readonly cx: number;
  /** Display height in OOXML EMUs (`<wp:extent cy>`). */
  readonly cy: number;
  /** `<wp:docPr id>` — must be unique across the document. */
  readonly docPrId: number;
  /** `<wp:docPr name>` — Word displays this in the alt-text dialog. */
  readonly name: string;
  /** Optional `<wp:docPr descr>` ("alt text"). */
  readonly descr?: string;
  /** Extra `<wp:inline>` / `<a:blip>` / `<pic:nvPicPr>` bits we model. */
  readonly properties?: InlineImageProperties;
  /**
   * Original `<w:drawing>` subtree captured at parse time. When present
   * AND no typed field has changed, the serializer re-emits these bytes
   * verbatim (byte-preservation fast path, mirrors `Table.raw`). Mutating
   * commands MUST drop `raw` on the new leaf they produce.
   */
  readonly raw?: OpaqueXml;
}

export interface InlineImageProperties {
  /** Optional `<wp:effectExtent>` carried verbatim. */
  readonly effectExtent?: { readonly t: number; readonly r: number; readonly b: number; readonly l: number };
  /** `<wp:inline distT/distB/distL/distR>` margins around the image. */
  readonly distT?: number;
  readonly distB?: number;
  readonly distL?: number;
  readonly distR?: number;
  /** `<wp:docPr title>` (some Word builds carry both `name` and `title`). */
  readonly title?: string;
  /** Catch-all for unmodelled wp:inline / pic:pic subtree fragments. */
  readonly opaqueProps?: ReadonlyArray<OpaqueXml>;
}

export interface OpaqueDrawing {
  readonly kind: "drawing";
  readonly subkind: "opaque";
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

/**
 * Inline analogue of `OpaqueBlock`. See that type for the dirty-tracking
 * contract; the same rules apply for inline carriers (mostly `<w:sdt>`,
 * `<w:fldSimple>`, `<w:smartTag>` appearing inside a paragraph).
 */
export interface OpaqueInline {
  readonly kind: "opaque-inline";
  readonly id: NodeId;
  readonly raw: OpaqueXml;
  readonly children?: ReadonlyArray<InlineNode>;
  readonly subtreeDirty?: boolean;
}

/* ── Header / footer parts ───────────────────────────────────────────────── */

/**
 * A typed `word/header*.xml` or `word/footer*.xml` part. Both kinds have the
 * same body shape (paragraph-level OOXML), so they share a single carrier
 * differentiated by `kind`. Tables inside header / footer parts are kept as
 * `OpaqueBlock` this round (typed table mutation lands in P1.3 / W7).
 */
export interface HeaderFooterPart {
  readonly kind: "header" | "footer";
  /** Stable id used by handlers to address this part (the part path itself). */
  readonly id: string;
  /** OOXML part path, e.g. `"word/header1.xml"`. Doubles as the handler id. */
  readonly partPath: string;
  /**
   * Which section variant this part services. Drawn from the
   * `<w:headerReference w:type>` / `<w:footerReference w:type>` element in
   * `<w:sectPr>`. When the same part is referenced from multiple sections
   * with conflicting types, we record the first one we see; this is purely
   * informational metadata and does not change which bytes Word reads.
   */
  readonly target: "default" | "first" | "even";
  /** Original namespace declarations from the part root (`w:hdr` / `w:ftr`). */
  readonly rootAttrs: Readonly<Record<string, string>>;
  /** Block-level body of the part. Tables stay opaque this round. */
  readonly body: ReadonlyArray<BlockNode>;
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
