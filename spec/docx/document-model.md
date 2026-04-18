# DOCX — In-Memory Model

## Roots

```typescript
import type { NodeId, DocumentSnapshot } from "@officeai/core";

export interface DocxSnapshot extends DocumentSnapshot<DocxDocument> {
  readonly format: "docx";
}

export interface DocxDocument {
  readonly id: NodeId;
  /**
   * Body block-level children: paragraphs, tables, opaque blocks.
   * Order matters; index is the natural address of a block.
   */
  readonly body: ReadonlyArray<BlockNode>;
  /**
   * Comments part. Empty array if word/comments.xml is absent or empty.
   */
  readonly comments: ReadonlyArray<DocxComment>;
  /**
   * The original namespace declarations + non-body root attributes from
   * word/document.xml. Re-emitted verbatim.
   */
  readonly documentRootAttrs: Readonly<Record<string, string>>;
  /**
   * The full container, used by the serializer to access untouched parts.
   */
  readonly containerSerial: number; // bumped when container changes
}
```

## Block nodes

```typescript
export type BlockNode = Paragraph | Table | OpaqueBlock | SectionBreak;

export interface Paragraph {
  readonly kind: "paragraph";
  readonly id: NodeId;
  readonly properties: ParagraphProperties;
  /**
   * Inline children in document order: runs, hyperlinks, comment markers,
   * revision wrappers, opaque inline blocks.
   */
  readonly children: ReadonlyArray<InlineNode>;
}

export interface ParagraphProperties {
  /** Style id, e.g. "Heading1", "Title", "ListParagraph". */
  readonly styleId?: string;
  readonly alignment?: "left" | "center" | "right" | "justify";
  readonly indentation?: { left?: number; right?: number; firstLine?: number; hanging?: number };
  readonly spacing?: {
    before?: number;
    after?: number;
    line?: number;
    lineRule?: "auto" | "exact" | "atLeast";
  };
  /** Numbering reference: numId + ilvl. Preserved verbatim; not introspected. */
  readonly numbering?: { numId: number; ilvl: number };
  /**
   * Anything in <w:pPr> we don't model explicitly. Re-emitted verbatim.
   */
  readonly opaqueProps?: ReadonlyArray<OpaqueXml>;
}

export interface Table {
  readonly kind: "table";
  readonly id: NodeId;
  /** The whole <w:tbl> as opaque XML this session. P1 will model rows/cells. */
  readonly raw: OpaqueXml;
}

export interface SectionBreak {
  readonly kind: "section-break";
  readonly id: NodeId;
  /** <w:sectPr> raw. Preserved verbatim. */
  readonly raw: OpaqueXml;
}

/**
 * A "content-wrapper" carrier (SDT / simple field / MC alternate
 * content / smart tag / custom XML) at the body level. The wrapper
 * itself is preserved verbatim through `raw`, and its inner content
 * is also parsed into typed `children` so the renderer can surface
 * the wrapped paragraphs as real headings/paragraphs instead of
 * collapsing the whole subtree into an opaque preview chip.
 *
 * Dirty-tracking contract:
 *
 *   - `subtreeDirty === false` (default) → serializer re-emits `raw`
 *     verbatim. `children` is purely a render-side projection.
 *   - `subtreeDirty === true` → serializer reconstructs the wrapper
 *     by splicing serialized `children` into the wrapper's content
 *     slot (e.g. `<w:sdtContent>`). Mutations that touch a child of
 *     an opaque carrier MUST flip this flag.
 */
export interface OpaqueBlock {
  readonly kind: "opaque-block";
  readonly id: NodeId;
  readonly raw: OpaqueXml;
  readonly children?: ReadonlyArray<BlockNode>;
  readonly subtreeDirty?: boolean;
}
```

## Inline nodes

```typescript
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
  /** Leaf children inside <w:r>: text, breaks, tabs, inline images (opaque). */
  readonly children: ReadonlyArray<RunChild>;
}

export interface RunProperties {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean | string; // string = underline pattern
  readonly strike?: boolean;
  readonly fontFamily?: string;
  readonly fontSize?: number; // half-points; OOXML stores them this way
  readonly color?: string; // RRGGBB hex (no '#')
  readonly highlight?: string; // named color per OOXML
  readonly opaqueProps?: ReadonlyArray<OpaqueXml>;
}

export type RunChild =
  | { kind: "text"; id: NodeId; text: string; xmlSpacePreserve: boolean }
  | { kind: "break"; id: NodeId; breakType?: "page" | "column" | "textWrapping" }
  | { kind: "tab"; id: NodeId }
  | { kind: "drawing"; id: NodeId; raw: OpaqueXml } // images live here
  | { kind: "opaque"; id: NodeId; raw: OpaqueXml };

export interface Hyperlink {
  readonly kind: "hyperlink";
  readonly id: NodeId;
  /** Either rId (external; resolved via _rels) or anchor (internal bookmark). */
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
 * Inline analogue of `OpaqueBlock`. Same dirty-tracking contract.
 */
export interface OpaqueInline {
  readonly kind: "opaque-inline";
  readonly id: NodeId;
  readonly raw: OpaqueXml;
  readonly children?: ReadonlyArray<InlineNode>;
  readonly subtreeDirty?: boolean;
}
```

## Comments part

```typescript
export interface DocxComment {
  readonly id: string; // OOXML comment id (string)
  readonly author: string;
  readonly initials?: string;
  readonly date: string; // ISO 8601
  /**
   * Comment body as block nodes (paragraphs of runs).
   */
  readonly body: ReadonlyArray<BlockNode>;
}
```

## Opaque XML carrier

```typescript
/**
 * A node we parsed but do not model in detail. Carries the original
 * fast-xml-parser subtree verbatim (with namespaces) so the serializer can
 * re-emit it without information loss.
 */
export interface OpaqueXml {
  /** Tag name including namespace prefix, e.g. "w:fldSimple". */
  readonly tag: string;
  /** Original attributes as parsed (with @_ prefix stripped). */
  readonly attrs: Readonly<Record<string, string>>;
  /** Subtree as fast-xml-parser preserveOrder array; serializer feeds it back in. */
  readonly subtree: unknown;
}
```

## Position and Selection

```typescript
/** Address of a single point inside the body. */
export interface DocxPosition {
  /** Index in DocxDocument.body. */
  readonly paragraph: number;
  /**
   * Optional finer addressing within the paragraph. If `run` is omitted,
   * the position is "at the start of the paragraph".
   */
  readonly run?: number;
  /** Character offset within the resolved text node, or 0 for boundary. */
  readonly offset?: number;
}

export interface DocxSelection {
  readonly start: DocxPosition;
  readonly end: DocxPosition;
}
```

## Invariants

1. Every node has a `NodeId` minted by the parser.
2. `Paragraph.children` is non-empty after parsing (a paragraph with no
   runs gets a single empty run for editing convenience).
3. `Run.children` may be empty (an empty run is valid).
4. Hyperlinks may not nest hyperlinks.
5. `comment-range-start` always pairs with a `comment-range-end` (in the
   same body, possibly different paragraph).
6. `RevisionWrapper.revisionType === "ins"` runs may be removed wholesale by
   reject; `"del"` runs may be unwrapped wholesale by reject.

The parser enforces these invariants; the command bus preserves them.
