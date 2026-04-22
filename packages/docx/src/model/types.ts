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
  /**
   * Whether `word/styles.xml` itself has been mutated and must be
   * re-emitted on save. Added in P3.1 / W1 alongside the typed
   * `StylesPart`. P3 ships only a read-only cascade resolver, so the
   * flag is always `false` in this milestone; mutating commands that
   * write to `StylesPart` (R10 / P4) would flip it.
   */
  styles: boolean;
  /**
   * Set of chart part paths (e.g. `"word/charts/chart1.xml"`) that have
   * been added or mutated since load. Untouched chart parts ride the
   * container's part cache; only entries in this set are re-emitted by
   * the serializer (which also (re)builds the embedded
   * `word/embeddings/Microsoft_Excel_WorksheetN.xlsx` so Office's
   * "Edit Data" UI keeps working). Added alongside the typed chart
   * model.
   */
  charts: ReadonlySet<string>;
  /**
   * Whether `word/footnotes.xml` itself has been mutated and must be
   * re-emitted on save. Added in F1 alongside the typed
   * {@link FootnotesPart}. Untouched documents leave the flag at
   * `false` so the part round-trips byte-identical via the container's
   * cache. Footnote-authoring commands (`docx:insert-footnote`,
   * `docx:set-footnote-body`, `docx:delete-footnote`) flip it on; the
   * serializer regenerates `word/footnotes.xml` from the typed model
   * (re-using each footnote's `raw` envelope when the footnote itself
   * is untouched, so a single-footnote edit only re-emits that
   * footnote's bytes).
   */
  footnotes: boolean;
  /**
   * Set of embedded-binary part paths (e.g.
   * `"word/embeddings/oleObject1.xlsx"`) that have been added or
   * mutated since load. Untouched embedded parts ride the container's
   * part cache; only entries in this set are re-emitted by the
   * serializer (which also registers the matching content-type
   * override). Used by `docx:insert-spreadsheet` and any future
   * OLE-package authoring path.
   */
  embeddings: ReadonlySet<string>;
  /**
   * Whether `word/settings.xml` has been mutated and must be re-emitted
   * on save. Currently only flipped by `docx:set-protection`, which
   * patches / removes the `<w:documentProtection>` element on the
   * settings root.
   */
  settings: boolean;
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
  /**
   * Typed `word/styles.xml` projection (P3.1 / W1). `undefined` when the
   * part is absent (synthetic test fixtures usually omit it).
   *
   * The cascade resolver in `agent/style-resolver.ts` walks
   * `docDefaults.rPrDefault` → the paragraph's `styleId` chain
   * (basedOn) → the paragraph's own `pPr.rPr` → the run's `rPr` to
   * compute the effective formatting at any selection. The toolbar
   * dropdowns read this so that "Heading 1" text shows `16` / `Calibri`
   * even when the run carries no direct `<w:rPr>` of its own.
   */
  readonly styles?: StylesPart;
  /**
   * Typed `word/theme/theme1.xml` projection. `undefined` when the
   * package has no theme part (synthetic test fixtures and the older
   * welcome doc both omit it). When present, the style cascade
   * resolver consults it to translate `<w:rFonts w:asciiTheme="…"/>`
   * (and friends) into the literal typeface Word would render.
   *
   * Round-trip contract: the part survives byte-identical via the
   * container cache. We never re-emit it from this typed model — it
   * is parsed read-only and `dirty.theme` does not exist in
   * {@link DocxDirtyFlags}. Adding theme mutation is a future
   * workstream when font-scheme authoring lands.
   */
  readonly theme?: ThemePart;
  /**
   * Typed projection of every `word/charts/chart*.xml` part referenced
   * from `word/document.xml.rels`. Keyed by the chart part path. The
   * Word DrawingML chart schema is identical to PowerPoint's, so the
   * shape mirrors PPTX's `ChartPart` — categories, series, type, title,
   * plus optional pointers to the embedded xlsx workbook that powers
   * Office's "Edit Data" round-trip.
   *
   * Round-trip contract: untouched chart parts re-emit from cached
   * bytes via the container; mutating commands (insert/edit chart) add
   * the part path to {@link DocxDirtyFlags.charts} and the serializer
   * regenerates both the chart XML and the embedded workbook from this
   * typed model.
   */
  readonly charts: ReadonlyMap<string, ChartPart>;
  /**
   * Embedded binary parts (xlsx packages, OLE blobs, …) keyed by part
   * path. Used by OLE-Excel-spreadsheet authoring (`docx:insert-spreadsheet`)
   * to ship the live `.xlsx` workbook bytes that Office activates on
   * double-click. Mirrors {@link DocxDocument.media} but for the
   * `word/embeddings/` directory, which uses the
   * `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`
   * content type rather than image MIME types.
   *
   * Round-trip contract: untouched embedded parts re-emit byte-identical
   * via the container's part cache; entries in `dirty.embeddings` are
   * (re-)written by the serializer.
   */
  readonly embeddings: ReadonlyMap<string, EmbeddedBinaryPart>;
  /**
   * Typed projection of `word/footnotes.xml` (F1). `undefined` when the
   * package has no footnotes part — the common case for short notes /
   * letters.
   *
   * Round-trip contract: untouched parts ride the container's byte
   * cache; only when `dirty.footnotes` is set does the serializer
   * regenerate the part from this typed model. Per-footnote `raw`
   * envelopes preserve every footnote (including the standard
   * `separator` / `continuationSeparator` notes Word inserts) when
   * the document is saved without any footnote-level mutation.
   *
   * `endnotesPart` is a future workstream and intentionally not
   * surfaced here; existing `word/endnotes.xml` parts continue to
   * round-trip byte-identical via the container cache.
   */
  readonly footnotesPart?: FootnotesPart;
  readonly documentRootAttrs: Readonly<Record<string, string>>;
  /**
   * Verbatim contents of `word/settings.xml`, when the part is present.
   * Mutated by `docx:set-protection` (which surgically patches the
   * `<w:documentProtection>` child of `<w:settings>` and dirties
   * `dirty.settings`). Untouched documents leave this string as the
   * load-time bytes so the part round-trips byte-identical via the
   * container's part cache.
   */
  readonly settingsXml?: string;
}

/* ── Footnotes part (F1) ─────────────────────────────────────────────────── */

/**
 * `<w:footnote w:type>` discriminator. Word writes the four standard
 * values; an absent attribute is treated as `"normal"` (a regular
 * authored footnote).
 */
export type FootnoteType = "normal" | "separator" | "continuationSeparator" | "continuationNotice";

/**
 * Typed `<w:footnote>` projection. Reuses the body block schema for
 * `body`, so the same parser, serializer, and ProseMirror schema work
 * unchanged inside footnote text.
 *
 * Round-trip contract: when `raw` is present AND no field of the
 * footnote has been mutated since parse, the serializer re-emits the
 * cached subtree byte-for-byte (mirrors `Table.raw`). Mutating commands
 * (e.g. `docx:set-footnote-body`) MUST drop `raw` on the new
 * `Footnote` they produce.
 */
export interface Footnote {
  /** OOXML id; -1 and 0 are conventionally separator/continuation. */
  readonly id: number;
  readonly type: FootnoteType;
  /** Typed body blocks; identical schema to {@link DocxDocument.body}. */
  readonly body: ReadonlyArray<BlockNode>;
  /**
   * Original `<w:footnote>` subtree captured at parse time. Present on
   * freshly parsed footnotes and on footnotes that have not been
   * touched by a mutating command. Absent (`undefined`) on footnotes
   * produced by mutation; the serializer uses its presence as the
   * per-footnote "clean / re-emit cached bytes" signal.
   */
  readonly raw?: OpaqueXml;
}

/**
 * Typed `word/footnotes.xml` carrier. Holds every `<w:footnote>` —
 * including the standard `separator` (`w:id="-1"`) and
 * `continuationSeparator` (`w:id="0"`) Word inserts on every
 * footnote-bearing document — in load order, so byte-identical
 * round-trip is preserved on no-edit save.
 *
 * `tail` captures any unmodelled top-level child of `<w:footnotes>`
 * (extension elements, comments, etc.). Today the parser puts
 * everything into typed `footnotes`; the field is reserved for future
 * extensions that might find non-`w:footnote` children in the wild.
 */
export interface FootnotesPart {
  readonly footnotes: ReadonlyArray<Footnote>;
  /** Original `<w:footnotes>` root attributes (namespace decls, etc.). */
  readonly rootAttrs: Readonly<Record<string, string>>;
  /** Other top-level children of `<w:footnotes>` we don't model. */
  readonly tail?: ReadonlyArray<OpaqueXml>;
}

/**
 * `<w:footnoteReference w:id="N"/>` promoted from {@link OpaqueRunChild}
 * to a typed run-child leaf in F1. The renderer uses this to draw the
 * superscript reference inline, and command handlers
 * (`docx:delete-footnote`) walk every run looking for matching
 * `footnoteId`s.
 */
export interface FootnoteReferenceLeaf {
  readonly kind: "footnote-ref";
  readonly id: NodeId;
  readonly footnoteId: number;
  /**
   * `<w:footnoteReference w:customMarkFollows="1"/>` — when set, Word
   * emits the immediately-following `<w:t>` as the user-visible mark
   * instead of the auto-numbered glyph. Captured for byte-identical
   * round-trip; the renderer ignores it for now.
   */
  readonly customMarkFollows?: boolean;
}

/* ── Chart parts (DOCX, mirrors PPTX shape) ─────────────────────────────── */

/** One series in a `<c:chart>`'s active plot type. */
export interface ChartSeries {
  readonly id: NodeId;
  /** Series ordering, emitted as both `c:idx` and `c:order`. */
  readonly idx: number;
  /** Optional series legend label. */
  readonly name?: string;
  readonly values: ReadonlyArray<number>;
}

export type ChartType = "bar" | "line" | "pie" | "area" | "unsupported";

/**
 * Typed `word/charts/chart*.xml` projection. `ChartPart` carries the
 * minimum needed to re-emit a self-consistent chart XML document plus
 * pointers back to the embedded xlsx workbook (when present). Existing
 * Office-authored charts that we don't yet model fully end up with
 * `chartType: "unsupported"` and ride the container's byte cache.
 */
export interface ChartPart {
  readonly partPath: string;
  readonly contentType: string;
  readonly chartType: ChartType;
  readonly title?: string;
  readonly categories: ReadonlyArray<string>;
  readonly series: ReadonlyArray<ChartSeries>;
  /** Path to the embedded xlsx workbook backing this chart, if any. */
  readonly embeddingPartPath?: string;
  /** Relationship id from this chart to its embedded workbook. */
  readonly embeddingRelId?: string;
  /** Sheet name inside the embedded workbook (defaults to `Sheet1`). */
  readonly embeddingSheetName?: string;
}

/**
 * A `<w:drawing>` leaf containing a `<c:chart>` reference. Captures the
 * chart's relationship id (resolved through `word/document.xml.rels`),
 * EMU display dimensions, and `<wp:docPr>` metadata so the serializer
 * can rebuild the wrapper from the typed model. The actual chart
 * payload (categories, series, type, title) lives in
 * {@link DocxDocument.charts} keyed by part path.
 */
export interface ChartDrawing {
  readonly kind: "drawing";
  readonly subkind: "chart";
  readonly id: NodeId;
  /** `r:id` of the chart relationship in `word/document.xml.rels`. */
  readonly relId: string;
  /** Resolved part path (e.g. `"word/charts/chart1.xml"`). */
  readonly chartPartPath: string;
  /** Display width in OOXML EMUs (`<wp:extent cx>`). */
  readonly cx: number;
  /** Display height in OOXML EMUs. */
  readonly cy: number;
  readonly docPrId: number;
  readonly name: string;
  readonly descr?: string;
  readonly raw?: OpaqueXml;
}

/* ── Styles part (P3.1) ──────────────────────────────────────────────────── */

/**
 * Typed projection of `word/styles.xml`. Mirrors the OOXML
 * `<w:styles>` shape: a flat collection of style definitions plus the
 * `<w:docDefaults>` root that supplies the bottom of the cascade.
 *
 * Round-trip contract: `raw` is the original `<w:styles>` subtree at
 * load time. The serializer emits `raw` verbatim while `dirty.styles
 * === false`. P3 ships read-only style cascade — no mutation commands
 * for styles — so `dirty.styles` is always `false` and the part
 * round-trips byte-identical.
 */
export interface StylesPart {
  readonly docDefaults: {
    readonly rPrDefault?: RunProperties;
    readonly pPrDefault?: ParagraphProperties;
  };
  /** Keyed by styleId. Iteration order matches load order. */
  readonly styles: ReadonlyMap<string, StyleDefinition>;
  /** Captured but not modeled bits (latentStyles, doc parts, etc.). */
  readonly raw?: OpaqueXml;
}

/* ── Theme part (P3.9 — font-scheme awareness) ───────────────────────────── */

/**
 * Typed projection of `word/theme/theme1.xml`. Only the font scheme is
 * modeled today — that is what the style cascade resolver needs to
 * translate `<w:rFonts w:asciiTheme="majorHAnsi"/>` to a literal
 * typeface ("Aptos Display" in Word 2024+, "Cambria" in Office 2007,
 * whatever a custom theme defines).
 *
 * The full DrawingML theme (color scheme, format scheme, custom theme
 * elements, …) round-trips byte-identical via the container cache;
 * mutating it is out of scope here. Added in P3.9.
 */
export interface ThemePart {
  readonly partPath: string;
  readonly majorFont: ThemeFontEntry;
  readonly minorFont: ThemeFontEntry;
}

/**
 * One side (`majorFont` or `minorFont`) of `<a:fontScheme>`. Only the
 * Latin typeface is consumed by the resolver today; East-Asian (`ea`)
 * and complex-script (`cs`) typefaces are captured for completeness so
 * a future workstream can resolve runs that target those scripts.
 */
export interface ThemeFontEntry {
  readonly latin: string;
  readonly ea?: string;
  readonly cs?: string;
}

export interface StyleDefinition {
  readonly id: string;
  readonly type: "paragraph" | "character" | "table" | "numbering";
  readonly name?: string;
  readonly basedOn?: string;
  readonly next?: string;
  readonly link?: string;
  readonly hidden?: boolean;
  readonly default?: boolean;
  readonly rPr?: RunProperties;
  readonly pPr?: ParagraphProperties;
  /** Anything we don't model on this `<w:style>` (uiPriority, qFormat, …). */
  readonly opaqueProps?: ReadonlyArray<OpaqueXml>;
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
 * Embedded binary part stored under `word/embeddings/`. Currently only
 * used for OLE-Excel `.xlsx` packages but kept generic so future OLE
 * binaries (`.bin`, `.docx`, `.pptx` embeds) can ride the same map.
 *
 * `contentType` is the OOXML override registered in
 * `[Content_Types].xml` for the part — usually
 * `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`.
 */
export interface EmbeddedBinaryPart {
  readonly partPath: string;
  readonly contentType: string;
  /**
   * Materialised bytes. Always present for parts loaded from an
   * existing package; absent for fresh parts authored by
   * `docx:insert-spreadsheet`, where the serializer builds the bytes
   * lazily from `pendingGrid` (so the command-handler stays sync and
   * doesn't have to await `JSZip.generateAsync`).
   */
  readonly bytes?: Uint8Array;
  /**
   * Source 2D grid for a freshly-authored OLE-Excel embed. The
   * serializer reads this via `buildEmbeddedXlsx` to produce the real
   * `.xlsx` bytes the first time the part is flushed. Mutually
   * exclusive with `bytes` for fresh inserts; once the embed has been
   * round-tripped through `serializeDocx` the bytes field is filled.
   */
  readonly pendingGrid?: ReadonlyArray<ReadonlyArray<string | number | null | undefined>>;
  /** Worksheet name used when materialising `pendingGrid`. */
  readonly pendingSheetName?: string;
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

export type BlockNode = Paragraph | Table | SectionBreak | OpaqueBlock | WrapperMarker;

/**
 * Body-level marker that brackets a *lifted* content-wrapper carrier
 * (`<w:sdt>`, `mc:AlternateContent`, `<w:fldSimple>`, `<w:smartTag>`,
 * `<w:customXml>`) whose `<w:sdtContent>` (etc.) used to nest body
 * blocks. The parser splits such carriers into
 *   `[ wrapper-begin, ...inner blocks, wrapper-end ]`
 * so the renderer + page chunker can flow the inner paragraphs as
 * regular body content (they paginate, they pick up per-section
 * geometry, they no longer get orphaned on a page of their own
 * because the wrapper used to be one giant atom).
 *
 * Round-trip safety: when the body is *not* dirty, the serializer
 * still emits the original `word/document.xml` bytes verbatim, so the
 * SDT envelope is preserved byte-for-byte. When the body is dirty,
 * the serializer walks `body`, re-emits the wrapper envelope from
 * `wrapperRaw` at every `wrapper-begin`, serializes the bracketed
 * inner blocks back into the wrapper's content slot, and closes the
 * envelope at the matching `wrapper-end`.
 */
export interface WrapperMarker {
  readonly kind: "wrapper-marker";
  readonly id: NodeId;
  /** "begin" opens the envelope; "end" closes it. */
  readonly side: "begin" | "end";
  /**
   * Stable id shared by the begin / end pair. Same `wrapperId` for
   * both ends so the serializer can pair them up even when other
   * markers nest in between (rare, but possible — e.g. an SDT inside
   * an SDT).
   */
  readonly wrapperId: string;
  /**
   * The carrier's full XML subtree as captured by the parser. The
   * serializer uses this to rebuild the envelope (carrier tag +
   * properties + content shell) verbatim, replacing only the
   * inner-blocks slot with the freshly serialized content.
   *
   * Both ends carry the SAME `wrapperRaw` (it describes the whole
   * carrier) so a partial body slice that contains the begin marker
   * but not the end one is still self-describing.
   */
  readonly wrapperRaw: OpaqueXml;
}

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
  /**
   * Pagination control flags (Phase 1 of docx-fidelity-overhaul).
   *
   * These are parsed into typed booleans so the page chunker can split
   * the body into Word-flavoured pages without re-walking opaqueProps.
   * The serializer continues to emit the original `<w:keepNext/>` /
   * `<w:keepLines/>` / `<w:pageBreakBefore/>` / `<w:widowControl/>`
   * elements through `opaqueProps` (each typed value is **also** kept
   * in `opaqueProps` so byte-for-byte round-trip is unaffected).
   *
   * `keepNext` — paragraph stays on the same page as the next block.
   * `keepLines` — all lines of this paragraph stay on the same page.
   * `pageBreakBefore` — Word forces a page break before this paragraph.
   * `widowControl` — when `false`, Word may leave a single-line widow
   *   at the page top/bottom; when `true` (or absent), Word avoids it.
   *   The XML uses `<w:widowControl/>` (true) or `<w:widowControl w:val="0"/>` (false).
   */
  readonly keepNext?: boolean;
  readonly keepLines?: boolean;
  readonly pageBreakBefore?: boolean;
  readonly widowControl?: boolean;
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
 *
 * Phase 2 of docx-fidelity-overhaul adds typed projections of
 * `<w:tblBorders>`, `<w:tblCellMar>`, `<w:tblLayout>`, `<w:tblInd>` so
 * the renderer can paint Word-flavoured tables (border colors, cell
 * padding, fixed/auto layout). Parser keeps each new element in
 * `opaqueProps` as well so the serializer round-trip is unaffected.
 */
export interface TableProperties {
  readonly width?: TableWidth;
  readonly jc?: "left" | "center" | "right" | "start" | "end";
  readonly tblBorders?: TableBorders;
  /** `<w:tblCellMar>` — default cell padding (twips). */
  readonly tblCellMar?: BoxSides;
  /** `<w:tblLayout w:type="…"/>` — `auto` (default) or `fixed`. */
  readonly tblLayout?: "auto" | "fixed";
  /** `<w:tblInd w:w="…"/>` — table indent in twips (positive = right of left margin). */
  readonly tblInd?: TableWidth;
  readonly opaqueProps?: ReadonlyArray<OpaqueXml>;
}

/**
 * One border side. `<w:top val="single" sz="4" color="000000"/>` becomes
 * `{ style: "single", size: 4, color: "000000" }`. `style` may be the
 * OOXML "nil" or "none" markers, in which case the renderer omits the
 * border entirely.
 */
export interface BorderSide {
  readonly style: string;
  /** Size in eighths of a point (`w:sz` units). */
  readonly size?: number;
  /** Hex RGB color (no leading `#`), or `"auto"` for theme default. */
  readonly color?: string;
  /** Padding between border and content, in points. */
  readonly space?: number;
}

/** Modeled subset of `<w:tblBorders>` (and `<w:tcBorders>`). */
export interface TableBorders {
  readonly top?: BorderSide;
  readonly left?: BorderSide;
  readonly bottom?: BorderSide;
  readonly right?: BorderSide;
  readonly insideH?: BorderSide;
  readonly insideV?: BorderSide;
}

/** Per-side spacing in twips. Used for cell padding (`<w:tcMar>`/`<w:tblCellMar>`). */
export interface BoxSides {
  readonly top?: Twips;
  readonly left?: Twips;
  readonly bottom?: Twips;
  readonly right?: Twips;
}

/**
 * Modeled subset of `<w:shd>`. `fill` is the background color (hex RGB);
 * `color` is the foreground/pattern color; `pattern` is the OOXML
 * `w:val` (e.g. `clear`, `pct25`, `solid`). The renderer only honours
 * `fill` for now; `pattern` round-trips through the raw cache.
 */
export interface Shading {
  readonly fill?: string;
  readonly color?: string;
  readonly pattern?: string;
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
  /** `<w:shd>` cell shading (background fill, pattern). */
  readonly shd?: Shading;
  /** `<w:tcBorders>` — per-cell border overrides. */
  readonly tcBorders?: TableBorders;
  /** `<w:vAlign w:val="…"/>` — vertical alignment of cell content. */
  readonly vAlign?: "top" | "center" | "bottom";
  /** `<w:tcMar>` — per-cell padding overrides (twips). */
  readonly tcMar?: BoxSides;
  readonly opaqueProps?: ReadonlyArray<OpaqueXml>;
}

/* ── Page geometry / section properties (P3.2) ───────────────────────────── */

/** OOXML twips (1/20 pt). Used for page sizes, margins, indents, etc. */
export type Twips = number;

/**
 * Typed projection of `<w:pgSz>`. `w` and `h` are page width/height in
 * twips. `orient` is the explicit `w:orient` attribute when present;
 * Word will also write `w:code` (size enum) which we don't model — it
 * round-trips through the parent section's `opaqueProps`.
 */
export interface PageSize {
  readonly w: Twips;
  readonly h: Twips;
  readonly orient?: "portrait" | "landscape";
}

/** Typed projection of `<w:pgMar>`. */
export interface PageMargins {
  readonly top: Twips;
  readonly right: Twips;
  readonly bottom: Twips;
  readonly left: Twips;
  readonly header: Twips;
  readonly footer: Twips;
  readonly gutter?: Twips;
}

/** Typed projection of `<w:cols>` (multi-column body layout). */
export interface PageColumns {
  readonly num: number;
  readonly sep?: boolean;
  readonly equalWidth?: boolean;
  readonly space?: Twips;
}

/**
 * Typed projection of `<w:headerReference>` / `<w:footerReference>`.
 * `relationshipId` resolves through `word/_rels/document.xml.rels` to a
 * concrete header/footer part stored in `DocxDocument.headersAndFooters`.
 */
export interface HeaderFooterRef {
  readonly type: "default" | "first" | "even";
  readonly relationshipId: string;
}

/**
 * Typed projection of `<w:sectPr>`. The renderer reads this to draw page
 * frames at the correct size, margins, and to pick the right
 * header/footer slot per page. Mutating commands (P3.4 / P3.6) edit
 * these fields in place; untouched sections re-emit `SectionBreak.raw`
 * verbatim for byte-identical round-trip.
 */
export interface SectionProperties {
  readonly pgSz?: PageSize;
  readonly pgMar?: PageMargins;
  readonly cols?: PageColumns;
  readonly headerRefs: ReadonlyArray<HeaderFooterRef>;
  readonly footerRefs: ReadonlyArray<HeaderFooterRef>;
  /** `<w:titlePg/>` — section uses the `first` header/footer on page 1. */
  readonly titlePg?: boolean;
  /**
   * `<w:type w:val>`. Drives Word's flow at the section boundary:
   * `continuous` keeps text on the same page; `nextPage` (default) starts
   * a fresh page; `oddPage`/`evenPage` skip to the next odd/even page.
   */
  readonly sectionType?: "continuous" | "nextPage" | "oddPage" | "evenPage" | "nextColumn";
  /**
   * Catch-all for `<w:sectPr>` children we don't model yet
   * (`<w:lineNumType>`, `<w:pgNumType>`, `<w:formProt>`, `<w:vAlign>`,
   * `<w:rtlGutter>`, `<w:docGrid>`, `<w:bidi>`, …). Captured in original
   * order so the serializer rebuild path preserves them.
   */
  readonly opaqueProps?: ReadonlyArray<OpaqueXml>;
}

export interface SectionBreak {
  readonly kind: "section-break";
  readonly id: NodeId;
  /**
   * Typed projection of the `<w:sectPr>`. Always present from P3.2
   * onwards; synthetic snapshots constructed without a `sectPr` set
   * `headerRefs`/`footerRefs` to `[]` and leave the geometry fields
   * undefined.
   */
  readonly properties: SectionProperties;
  /**
   * Original `<w:sectPr>` subtree captured at parse time. When present
   * AND no field of `properties` has been mutated, the serializer
   * re-emits `raw` verbatim (byte-preservation fast path, mirrors
   * `Table.raw` / `InlineImageDrawing.raw`). Mutating commands MUST
   * drop `raw` on the new `SectionBreak` they produce.
   */
  readonly raw?: OpaqueXml;
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
  /**
   * Literal Latin typeface from `<w:rFonts w:ascii="…"/>`. Wins over
   * {@link fontFamilyAsciiTheme} when both are present at the same
   * cascade level — that is Word's resolution rule, mirrored in
   * `agent/style-resolver.ts`.
   */
  readonly fontFamily?: string;
  /**
   * Theme reference from `<w:rFonts w:asciiTheme="…"/>`. Typical
   * values: `majorHAnsi`, `minorHAnsi`, `majorAscii`, `minorAscii`.
   * Resolved through {@link ThemePart} (or the Word-default fallback
   * map in `style-resolver.ts`) to a literal typeface. We keep the
   * raw value rather than eagerly projecting at parse time so the
   * cascade resolver can re-project against an updated theme part
   * without re-parsing.
   */
  readonly fontFamilyAsciiTheme?: string;
  /** Companion to {@link fontFamilyAsciiTheme}; from `w:hAnsiTheme`. */
  readonly fontFamilyHAnsiTheme?: string;
  /**
   * Literal high-ANSI typeface from `<w:rFonts w:hAnsi="…"/>`. Excel
   * mirrors this from `w:ascii` for most western documents; it is
   * tracked separately so {@link fontFamily} edits do not destroy
   * source documents that set it explicitly.
   */
  readonly fontFamilyHAnsi?: string;
  /**
   * Literal East-Asian typeface from `<w:rFonts w:eastAsia="…"/>`.
   * Required for CJK documents; the default mapping rule (per the
   * OOXML spec) is "use this font for any character whose Unicode
   * block falls under the East Asian set". Editing `fontFamily`
   * should NOT silently overwrite this slot.
   */
  readonly fontFamilyEastAsia?: string;
  /**
   * Literal complex-script typeface from `<w:rFonts w:cs="…"/>`.
   * Used for Arabic, Hebrew, Indic, etc.
   */
  readonly fontFamilyComplexScript?: string;
  /** Theme reference from `<w:rFonts w:eastAsiaTheme="…"/>`. */
  readonly fontFamilyEastAsiaTheme?: string;
  /** Theme reference from `<w:rFonts w:cstheme="…"/>`. */
  readonly fontFamilyComplexScriptTheme?: string;
  readonly fontSize?: number;
  readonly color?: string;
  readonly highlight?: string;
  readonly opaqueProps?: ReadonlyArray<OpaqueXml>;
}

export type RunChild =
  | TextLeaf
  | BreakLeaf
  | TabLeaf
  | DrawingLeaf
  | PageBreakLeaf
  | LastRenderedPageBreakLeaf
  | PageNumberFieldLeaf
  | FootnoteReferenceLeaf
  | EmbeddedSpreadsheet
  | OpaqueRunChild;

/**
 * `<w:object>` whose `<o:OLEObject ProgID>` matches `Excel.Sheet.*`
 * (typically `Excel.Sheet.12`). The "live" Excel embed in Word —
 * double-click in Word pops the embedded `.xlsx` open in Excel for
 * editing. Symmetric to PPTX's `OleSpreadsheetShape`.
 *
 * We type just enough to (a) detect the embed, (b) round-trip the
 * embedded workbook + preview image relationships, and (c) author
 * fresh embeds via `docx:insert-spreadsheet`. Everything else (VML
 * shape attrs, `<o:OLEObject>` tail children, `<w:objectEmbed>`
 * variants) is captured opaquely so existing files survive
 * byte-identical no-touch saves.
 */
export interface EmbeddedSpreadsheet {
  readonly kind: "embedded-spreadsheet";
  readonly id: NodeId;
  /** `<o:OLEObject r:id>` — relationship pointing at the embedded part. */
  readonly oleRelId: string;
  /** Resolved package-absolute path of the embedded `.xlsx` (or `.bin`). */
  readonly embeddingPartPath: string;
  /** `<o:OLEObject ProgID>` — e.g. `Excel.Sheet.12`. */
  readonly progId: string;
  /** Embedded part kind: `xlsx` (true Excel package) or `bin` (legacy CFB). */
  readonly embeddingKind: "xlsx" | "bin";
  /** `<v:imagedata r:id>` — preview image relationship id, if known. */
  readonly previewImageRelId?: string;
  /** Resolved preview image part path, if `previewImageRelId` was wired. */
  readonly previewImagePartPath?: string;
  /** `<o:OLEObject>` attribute bag (Type, ShapeID, DrawAspect, ObjectID, …). */
  readonly oleObjectAttrs: Readonly<Record<string, string>>;
  /**
   * Original `<w:object>` subtree captured at parse time. When present
   * AND no typed field has changed, the serializer re-emits these
   * bytes verbatim (byte-preservation fast path).
   */
  readonly raw?: OpaqueXml;
}

/**
 * `<w:fldSimple w:instr=" PAGE \* MERGEFORMAT "/>` and the equivalent
 * `<w:fldSimple w:instr=" NUMPAGES "/>`. Promoted to a typed leaf in
 * P3.4 / W15 so the toolbar can produce page-number fields and the
 * paged-renderer can render the live page index.
 *
 * The complex `<w:fldChar>`-bracketed multi-run form stays as
 * {@link OpaqueRunChild} for now; it requires multi-run reassembly
 * across siblings that the parser does not yet do.
 *
 * Round-trip invariant: parse → serialize → parse produces the same
 * `instr` string (including switches like `\* MERGEFORMAT`).
 */
export interface PageNumberFieldLeaf {
  readonly kind: "page-number-field";
  readonly id: NodeId;
  /** Variant of the field. Determines what the renderer substitutes. */
  readonly field: "PAGE" | "NUMPAGES";
  /**
   * The literal `w:instr` attribute as it appeared in the source XML
   * (e.g. `" PAGE \\* MERGEFORMAT "`). Captured verbatim so the
   * serializer can re-emit byte-identical bytes when the leaf is
   * untouched.
   */
  readonly instr: string;
  /**
   * Optional cached display value Word writes inside the field
   * (`<w:t>3</w:t>`). Carried through for byte round-trip. The
   * runtime renderer ignores this and substitutes the live page
   * index from the page chunker.
   */
  readonly cachedText?: string;
}

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
  /**
   * `column` and `textWrapping` are kept on this legacy leaf. The page
   * break case (`<w:br w:type="page"/>`) is promoted to the typed
   * {@link PageBreakLeaf} at parse time so the page chunker can switch
   * on `kind` directly.
   */
  readonly breakType?: "column" | "textWrapping";
}

/**
 * A `<w:br w:type="page"/>` inside a run. Promoted to a typed leaf in
 * P3.2 / W6 so the page-chunker can split the body without reaching into
 * opaque carriers. Round-trips back to `<w:br w:type="page"/>`.
 */
export interface PageBreakLeaf {
  readonly kind: "page-break";
  readonly id: NodeId;
}

/**
 * `<w:lastRenderedPageBreak/>`. Word writes this hint at the position
 * where pagination broke during the last save. Layout-only — no
 * formatting, no content. The page chunker uses it as a cheap heuristic
 * for picking initial page breaks but never treats it as authoritative
 * (Word may not have written it; the geometry may have changed since).
 */
export interface LastRenderedPageBreakLeaf {
  readonly kind: "last-rendered-page-break";
  readonly id: NodeId;
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
export type DrawingLeaf = InlineImageDrawing | ChartDrawing | OpaqueDrawing;

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
