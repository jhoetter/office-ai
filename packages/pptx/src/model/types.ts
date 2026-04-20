/**
 * PPTX in-memory model.
 *
 * See spec/pptx/document-model.md and spec/pptx/ooxml-mapping.md.
 */

import type { DocumentSnapshot, NodeId, ooxml as _ooxml } from "@officeai/core";
import type { ThemeColorScheme } from "../renderer/layout/color.js";

export type { ThemeColorScheme };

// Re-imported as a type so we don't depend on the runtime namespace.
type OoxmlContainer = InstanceType<typeof _ooxml.OoxmlContainer>;

// ─── Snapshot + dirty flags ───────────────────────────────────────────────

export interface PptxSnapshot extends DocumentSnapshot<PptxPresentation> {
  readonly format: "pptx";
  /** Per-part dirty flags. Drives the serializer's per-part decision. */
  readonly dirty: PptxDirty;
  /** Set of part paths to remove on serialize (delete-slide etc.). */
  readonly removedParts: ReadonlySet<string>;
  /**
   * The OOXML container, attached at parse time so the serializer can
   * reach untouched-part bytes. Treated as opaque from the model's POV.
   */
  readonly container: OoxmlContainer;
  /** Per-part-path RelationshipGraph snapshots. */
  readonly relationships: ReadonlyMap<string, RelationshipsSnap>;
  /** Content-types snapshot. */
  readonly contentTypes: ContentTypesSnap;
}

export interface PptxDirty {
  readonly presentation: boolean;
  readonly slides: ReadonlySet<string>;
  readonly notesSlides: ReadonlySet<string>;
  readonly masters: ReadonlySet<string>;
  readonly layouts: ReadonlySet<string>;
  readonly theme: ReadonlySet<string>;
  readonly media: ReadonlySet<string>;
  /** F3: chart parts to re-emit from the typed model rather than from container bytes. */
  readonly charts: ReadonlySet<string>;
  /**
   * Embedded-binary part paths (under `ppt/embeddings/`) that have
   * been added or mutated since load. Used by `pptx:insert-spreadsheet`
   * for OLE-Excel `.xlsx` packages; untouched embeddings ride the
   * container's part cache.
   */
  readonly embeddings: ReadonlySet<string>;
  /** Per-slide-comments-part dirty set — keyed by `ppt/comments/commentN.xml`. */
  readonly comments: ReadonlySet<string>;
  /** True when `ppt/commentAuthors.xml` needs to be rebuilt. */
  readonly commentAuthors: boolean;
  readonly relationships: ReadonlySet<string>;
  readonly contentTypes: boolean;
}

export const emptyDirty = (): PptxDirty => ({
  presentation: false,
  slides: new Set<string>(),
  notesSlides: new Set<string>(),
  masters: new Set<string>(),
  layouts: new Set<string>(),
  theme: new Set<string>(),
  media: new Set<string>(),
  charts: new Set<string>(),
  embeddings: new Set<string>(),
  comments: new Set<string>(),
  commentAuthors: false,
  relationships: new Set<string>(),
  contentTypes: false,
});

export interface RelationshipSnap {
  readonly id: string;
  readonly type: string;
  readonly target: string;
  readonly targetMode?: "External" | "Internal";
}

export interface RelationshipsSnap {
  readonly relsPath: string;
  readonly entries: ReadonlyArray<RelationshipSnap>;
}

export interface ContentTypesDefault {
  readonly extension: string;
  readonly contentType: string;
}

export interface ContentTypesOverride {
  readonly partName: string;
  readonly contentType: string;
}

export interface ContentTypesSnap {
  readonly defaults: ReadonlyArray<ContentTypesDefault>;
  readonly overrides: ReadonlyArray<ContentTypesOverride>;
}

// ─── Presentation root ────────────────────────────────────────────────────

export interface PptxPresentation {
  readonly id: NodeId;
  readonly slides: ReadonlyArray<Slide>;
  readonly slideSize: SlideSize;
  readonly notesSize?: SlideSize;
  readonly masters: ReadonlyMap<string, OpaquePart>;
  readonly layouts: ReadonlyMap<string, SlideLayout>;
  readonly theme: ReadonlyMap<string, OpaquePart>;
  /**
   * Resolved color scheme from the first theme part (`a:clrScheme`). Used
   * by the renderer to draw `<a:schemeClr>` references with the correct
   * palette. Falls back to `DEFAULT_THEME` if no theme part is present
   * or the scheme can't be parsed.
   */
  readonly themeDefault: ThemeColorScheme;
  readonly notesSlides: ReadonlyMap<string, NotesSlide>;
  /**
   * Per-slide-comments map keyed by `ppt/comments/commentN.xml`. Each
   * entry holds a typed list of comments belonging to that slide.
   * Slides reference their comments part via `Slide.commentsPartPath`.
   */
  readonly commentsByPart: ReadonlyMap<string, PptxCommentsPart>;
  /**
   * Single workbook-wide author registry from `ppt/commentAuthors.xml`.
   * Comments reference authors by `authorId`; `null` means the deck
   * has no comments yet (no author part on disk).
   */
  readonly commentAuthors: PptxCommentAuthorsPart | null;
  readonly media: ReadonlyMap<string, MediaPart>;
  /**
   * F3: typed chart parts keyed by part path. Each `ChartShape` resolves
   * its `chartPartPath` here. Anything in the chart XML we don't model
   * is preserved as `OpaqueXml` inside the part.
   */
  readonly charts: ReadonlyMap<string, ChartPart>;
  /**
   * Embedded binary parts under `ppt/embeddings/` keyed by part path.
   * Currently the home of OLE-Excel `.xlsx` workbooks shipped by
   * `pptx:insert-spreadsheet`. Mirrors `media` but for embeddings;
   * untouched parts round-trip byte-identical via the container.
   */
  readonly embeddings: ReadonlyMap<string, EmbeddedBinaryPart>;
  readonly presentationRootAttrs: Readonly<Record<string, string>>;
  readonly presentationOpaqueTail: ReadonlyArray<OpaqueXml>;
  /**
   * Raw <p:sldIdLst> attrs (id is mostly empty; preserved for round-trip).
   */
  readonly sldIdLstAttrs: Readonly<Record<string, string>>;
  readonly idGen: PptxIdGen;
}

export interface SlideSize {
  readonly cxEmu: number;
  readonly cyEmu: number;
  readonly type?: string;
}

export interface PptxIdGen {
  readonly nextSlideId: number;
  readonly nextSlidePartIndex: number;
  readonly nextMediaPartIndex: number;
}

export interface OpaquePart {
  readonly partPath: string;
  readonly raw: OpaqueXml;
}

// ─── Slide layouts ────────────────────────────────────────────────────────

/**
 * The 11 PowerPoint-standard layouts plus an `unknown` escape hatch for
 * layouts the classifier can't pigeonhole. Naming mirrors the names used
 * in the desktop app's layout picker.
 */
export type LayoutKind =
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
  | "bigNumber"
  | "unknown";

/**
 * Subset of `<p:ph>` attributes we care about for layout cloning. Every
 * placeholder also remembers the (idx, type) pair so when the user
 * switches layouts we can preserve any user content that maps cleanly
 * to a placeholder of the same idx in the new layout.
 */
export interface PlaceholderSpec {
  /** `<p:ph type="...">`. Defaults to `"body"` when absent (PowerPoint's default). */
  readonly type: string;
  /** `<p:ph idx="...">`. Defaults to 0. */
  readonly idx: number;
  /** Optional `<p:ph sz="...">` (e.g. `"quarter"`, `"half"`, `"full"`). */
  readonly sz?: string;
  /** Optional position from `<a:xfrm>/<a:off>` if present in the layout. */
  readonly position?: Position;
  /** Optional size from `<a:xfrm>/<a:ext>` if present in the layout. */
  readonly size?: Size;
}

// ─── Speaker notes ────────────────────────────────────────────────────────

/**
 * Typed speaker-notes slide. Promoted from `OpaquePart` so we can edit
 * the body placeholder text directly without having to rewrite raw XML.
 *
 * Every notes part has at most one "body" placeholder (PowerPoint's
 * convention for the prompt that holds the actual notes); the parser
 * extracts that into `body`. Anything else inside the notes part — the
 * slide-image placeholder, header/footer placeholders, formatting,
 * extLst — stays in `raw` for byte-faithful round-trip when the body
 * isn't dirtied.
 */
export interface NotesSlide {
  readonly partPath: string;
  /** Typed text-body of the notes "body" placeholder, when present. */
  readonly body: TextBody;
  /** Verbatim `<p:notes>` blob preserved for round-trip. */
  readonly raw: OpaqueXml;
}

// ─── Comments ─────────────────────────────────────────────────────────────

/**
 * One author record from `ppt/commentAuthors.xml`. PowerPoint indexes
 * authors by a numeric `id` and references it from each comment's
 * `authorId`. We keep the XML's optional `initials`, `lastIdx`, and
 * `clrIdx` for round-trip even though only `name` matters to the UI.
 */
export interface PptxCommentAuthor {
  readonly id: number;
  readonly name: string;
  readonly initials?: string;
  /** Highest `idx` PowerPoint has minted for this author so far. */
  readonly lastIdx?: number;
  readonly clrIdx?: number;
}

export interface PptxCommentAuthorsPart {
  readonly partPath: string;
  readonly authors: ReadonlyArray<PptxCommentAuthor>;
  /** Verbatim `<p:cmAuthorLst>` blob preserved for round-trip. */
  readonly raw?: OpaqueXml;
}

/**
 * One PowerPoint comment. Position is in EMU on the slide (PowerPoint
 * uses a fixed 914400 EMU per inch). When `parentId` is set, the
 * comment is a reply in the thread rooted at that id. `resolved` is a
 * synthetic flag we serialise into a `<p:extLst>` extension so it
 * survives a round-trip without losing the rest of the comment's data.
 */
export interface PptxComment {
  /** Stable, deck-unique id minted by the parser/agent. */
  readonly id: string;
  /** Reference into `commentAuthors.authors`. */
  readonly authorId: number;
  /** Per-author monotonic index — required by the OOXML schema. */
  readonly idx: number;
  /** ISO-8601 creation timestamp (`<p:cm dt="…">`). */
  readonly createdAt?: string;
  /** Pin position on the slide, in EMU. */
  readonly xEmu: number;
  readonly yEmu: number;
  /** Plain-text body. PowerPoint comments don't carry rich formatting. */
  readonly text: string;
  /** Reply-to id (top-level comments leave this undefined). */
  readonly parentId?: string;
  /**
   * Optional anchored shape id. When set, the comment is conceptually
   * attached to that shape (the canvas paints a yellow indicator over
   * it). Persisted via an `officeai:shapeAnchor` extension on `<p:cm>`
   * so it round-trips through PowerPoint without losing data.
   */
  readonly shapeId?: string;
  /** Synthetic resolved flag — round-tripped via an extLst extension. */
  readonly resolved?: boolean;
}

export interface PptxCommentsPart {
  readonly partPath: string;
  readonly comments: ReadonlyArray<PptxComment>;
}

export interface SlideLayout {
  readonly partPath: string;
  /** Layout classification — drives the picker grid + add-slide cloning. */
  readonly kind: LayoutKind;
  /** `<p:cSld name="...">` if present, falls back to a friendly default. */
  readonly name: string;
  /** Every placeholder declared on the layout, in document order. */
  readonly placeholders: ReadonlyArray<PlaceholderSpec>;
  /**
   * Verbatim `<p:sldLayout>` blob — used by the serializer to write the
   * part back unchanged when the layout itself wasn't edited. For
   * built-in layouts we synthesise on demand this comes from the
   * builtin-layouts factory.
   */
  readonly raw: OpaqueXml;
}

export interface MediaPart {
  readonly partPath: string;
  readonly contentType: string;
  readonly bytes: Uint8Array;
  readonly sha256: string;
}

/**
 * Embedded binary part stored under `ppt/embeddings/`. Currently used
 * for OLE-Excel `.xlsx` packages (`progId` = `Excel.Sheet.12`) but kept
 * generic so future OLE binaries can ride the same map.
 */
export interface EmbeddedBinaryPart {
  readonly partPath: string;
  readonly contentType: string;
  /**
   * Materialised bytes. Always present for parts loaded from an
   * existing package; absent for fresh parts authored by
   * `pptx:insert-spreadsheet`, where the serializer builds the bytes
   * lazily from `pendingGrid` so the command handler stays sync.
   */
  readonly bytes?: Uint8Array;
  /** Source 2D grid for a freshly-authored OLE-Excel embed. */
  readonly pendingGrid?: ReadonlyArray<ReadonlyArray<string | number | null | undefined>>;
  /** Worksheet name used when materialising `pendingGrid`. */
  readonly pendingSheetName?: string;
}

// ─── Chart parts (F3) ─────────────────────────────────────────────────────

export type ChartType = "bar" | "line" | "pie" | "area" | "unsupported";

export interface ChartSeries {
  readonly id: NodeId;
  /** `<c:ser><c:idx val>`. */
  readonly idx: number;
  readonly name?: string;
  readonly values: ReadonlyArray<number>;
}

/**
 * F3: typed chart part referenced from a `ChartShape`. Only the four typed
 * fields (`title`, `chartType`, `categories`, `series`) are model-driven;
 * everything else (axes, legend, plot-area styling, embedded xlsx) is
 * preserved verbatim and re-emitted byte-for-byte unless the part is
 * marked dirty.
 */
export interface ChartPart {
  readonly partPath: string;
  readonly contentType: string;
  readonly chartType: ChartType;
  readonly title?: string;
  readonly categories: ReadonlyArray<string>;
  readonly series: ReadonlyArray<ChartSeries>;
  /** `<c:chartSpace>` head + tail (sans the children we model). Verbatim. */
  readonly chartSpaceRaw: OpaqueXml;
  /** `<c:plotArea>` siblings (axes, legend, …). Verbatim. */
  readonly plotAreaTailRaw: ReadonlyArray<OpaqueXml>;
  /** Verbatim `<c:ser>` head/tail keyed by series idx. */
  readonly seriesRaw: ReadonlyMap<number, OpaqueXml>;
  /**
   * Package-absolute path of the embedded `.xlsx` package backing this
   * chart's `c:externalData`, if any. Word/PowerPoint always author one
   * (e.g. `ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx`); we
   * surface it so the serializer can refresh the bytes when the typed
   * data model changes. Charts authored without an embedded workbook
   * leave this `undefined` and the serializer mints one on first save.
   */
  readonly embeddingPartPath?: string;
  /**
   * Relationship id (relative to the chart part's `_rels/chartN.xml.rels`)
   * pointing at the embedded workbook. Required for `<c:externalData>`.
   * Synthesised by the serializer when missing.
   */
  readonly embeddingRelId?: string;
  /**
   * Worksheet name inside the embedded workbook the chart references in
   * its `<c:f>` strings (e.g. `"Sheet1"` or PowerPoint's German
   * `"Tabelle1"`). Defaults to `"Sheet1"` for newly-authored charts.
   */
  readonly embeddingSheetName?: string;
}

// ─── Slide ────────────────────────────────────────────────────────────────

export interface Slide {
  readonly id: NodeId;
  readonly partPath: string;
  readonly slideId: number;
  readonly relId: string;
  readonly layoutPartPath?: string;
  readonly notesSlidePartPath?: string;
  /**
   * Per-slide comments part path (`ppt/comments/commentN.xml`). Resolved
   * via the slide's rels and pointed at `commentsByPart`. Absent slides
   * have no comments.
   */
  readonly commentsPartPath?: string;
  readonly shapes: ReadonlyArray<Shape>;
  /**
   * F4: typed slide transition (`<p:transition>`). Optional — slides
   * without a transition have no field. Unsupported transitions still
   * land here with `kind: "unsupported"` plus their verbatim raw blob.
   */
  readonly transition?: SlideTransition;
  /**
   * F4: typed simple per-shape entrance animations parsed from
   * `<p:timing>`. Anything we don't model stays in `timingTailRaw`.
   */
  readonly animations: ReadonlyArray<EntranceAnimation>;
  /**
   * F4: opaque remainder of `<p:timing>` we don't introspect (build,
   * advanced sequences, custom paths, behaviors). Re-emitted verbatim.
   */
  readonly timingTailRaw?: OpaqueXml;
  /**
   * Tail children of `<p:sld>` we don't introspect (clrMapOvr, extLst, …).
   * `<p:transition>` and `<p:timing>` are removed iff promoted above;
   * otherwise they remain here verbatim.
   */
  readonly slideOpaqueTail: ReadonlyArray<OpaqueXml>;
  readonly slideRootAttrs: Readonly<Record<string, string>>;
  readonly cSldAttrs: Readonly<Record<string, string>>;
  /** Head children of <p:cSld> + <p:spTree> we don't introspect. */
  readonly spTreeHead: ReadonlyArray<OpaqueXml>;
  /** Optional <p:bg> captured if present (sits inside <p:cSld>). */
  readonly cSldHead: ReadonlyArray<OpaqueXml>;
}

// ─── Animations / transitions (F4) ───────────────────────────────────────

export type TransitionKind = "none" | "fade" | "push" | "wipe" | "split" | "cut" | "unsupported";

export type TransitionSpeed = "slow" | "med" | "fast";

/**
 * F4: typed slide transition. Only the six listed kinds round-trip
 * cleanly; anything else stays as `kind: "unsupported"` with its
 * verbatim `<p:transition>` blob preserved in `raw`.
 */
export interface SlideTransition {
  readonly id: NodeId;
  readonly kind: TransitionKind;
  readonly speed?: TransitionSpeed;
  /** Verbatim `<p:transition>` blob for unsupported kinds and tail. */
  readonly raw?: OpaqueXml;
}

// ─── Expanded animation gallery (F4 v2) ──────────────────────────────────
//
// The first iteration of F4 only modelled four entrance effects with a
// single `<p:set>` body. PowerPoint's animation gallery covers four
// categories — Entrance, Emphasis, Exit and Motion Paths — driven by
// the same `<p:cTn @presetClass @presetID>` mechanism. The expanded
// model below is a discriminated union over all four so the parser /
// serializer / renderer can dispatch through a single registry.

export type AnimationCategory = "entrance" | "emphasis" | "exit" | "motionPath";

/**
 * Triggering behaviour of an animation step relative to its predecessor.
 * Maps directly to OOXML `<p:par><p:cTn nodeType="…">`:
 *   - `onClick`        → `clickEffect` (default; advances on click)
 *   - `withPrevious`   → `withEffect` (starts simultaneously)
 *   - `afterPrevious`  → `afterEffect` (starts when previous finishes)
 */
export type AnimationTrigger = "onClick" | "withPrevious" | "afterPrevious";

/**
 * Effect direction. Mirrors PowerPoint's "Effect Options" submenu and is
 * preset-specific: only certain presets accept directions and the set of
 * accepted values varies (FlyIn allows all 4 cardinals, Wipe accepts
 * `left|right|up|down`, Spin accepts the two rotational directions, …).
 *
 * The registry in `animation/presets.ts` declares per-preset valid
 * directions; the UI hides the picker for presets that don't accept any.
 */
export type AnimationDirection =
  | "left"
  | "right"
  | "up"
  | "down"
  | "in"
  | "out"
  | "horizontal"
  | "vertical"
  | "clockwise"
  | "counterclockwise";

export type EntrancePreset =
  | "appear"
  | "fade"
  | "flyIn"
  | "floatIn"
  | "split"
  | "wipe"
  | "shape"
  | "wheel"
  | "randomBars"
  | "growAndTurn"
  | "zoom"
  | "swivel"
  | "bounce";

export type EmphasisPreset =
  | "pulse"
  | "colorPulse"
  | "teeter"
  | "spin"
  | "growShrink"
  | "desaturate"
  | "fontColor"
  | "lineColor";

export type ExitPreset =
  | "disappear"
  | "fade"
  | "flyOut"
  | "floatOut"
  | "split"
  | "wipe"
  | "shape"
  | "wheel"
  | "zoom";

export type MotionPathPreset = "line" | "arc" | "turn" | "loops" | "custom";

/** Discriminated alias used by the renderer / picker. */
export type AnimationPreset = EntrancePreset | EmphasisPreset | ExitPreset | MotionPathPreset;

/**
 * F4 v2 — typed shape animation. Replaces the earlier `EntranceAnimation`
 * (which is now a thin alias for backward-compat). Targets a shape by
 * `cNvPrId` (what `<p:spTgt @spid>` references). Only effects listed in
 * the preset registry round-trip via the typed serializer; anything
 * else is preserved through `timingTailRaw` (slide-level) or `raw`
 * (per-animation) blob capture.
 */
export interface ShapeAnimation {
  readonly id: NodeId;
  readonly targetCNvPrId: number;
  readonly category: AnimationCategory;
  readonly preset: AnimationPreset;
  /** Defaults to `"onClick"`. */
  readonly trigger: AnimationTrigger;
  /** Optional direction (only for direction-aware presets). */
  readonly direction?: AnimationDirection;
  /** `<p:cTn @dur>` in milliseconds. Defaults to per-preset value. */
  readonly durationMs?: number;
  /** `<p:cTn @delay>` in milliseconds. */
  readonly delayMs?: number;
  /** Trigger order in the main entrance/emphasis/exit sequence (0-based). */
  readonly order: number;
  /**
   * MotionPath only: SVG-style command string in slide-relative
   * coordinates (origin = top-left of the bounding rect at start;
   * 1 unit = full slide width / height). Translates to OOXML's
   * compact path syntax (`M`, `L`, `C`, `E` for end) when emitted.
   */
  readonly motionPath?: string;
  /**
   * Captured raw `<p:par>` blob. When present, the serializer re-emits
   * this verbatim instead of synthesising from the typed fields — used
   * to preserve byte-perfect round-trip for animations imported from
   * an existing PPTX. Cleared whenever the typed fields are mutated.
   */
  readonly raw?: OpaqueXml;
}

/**
 * @deprecated Use `ShapeAnimation`. Retained as a structural alias so
 * external imports of `EntranceAnimation` keep compiling. New code
 * should import `ShapeAnimation` directly.
 */
export type EntranceAnimation = ShapeAnimation;

/**
 * @deprecated Use the discriminated `EntrancePreset` union. The legacy
 * 4-value type is kept for external imports; new code should pull
 * `EntrancePreset` from this module.
 */
export type EntranceEffect = "appear" | "fade" | "fly-in" | "wipe";

// ─── Shapes ───────────────────────────────────────────────────────────────

export type Shape =
  | TextShape
  | Picture
  | TableShape
  | ChartShape
  | OleSpreadsheetShape
  | GroupShape
  | ConnectorShape
  | OpaqueShape;

export type ShapeKind = Shape["kind"];

// ─── Connectors ───────────────────────────────────────────────────────────

/**
 * One end of a connector. Anchored endpoints reference a target shape's
 * `cNvPrId` (the OOXML numeric id used by `<a:stCxn>` / `<a:endCxn>`)
 * plus a side; free endpoints carry an absolute slide-coordinate point.
 *
 * Side strings reuse the existing `AnchorSide` from the renderer so the
 * canvas's snap-to-anchor result can be persisted directly.
 */
export type ConnectorSide = "n" | "s" | "e" | "w" | "center";

export type ConnectorEndpoint =
  | { readonly kind: "free"; readonly xEmu: number; readonly yEmu: number }
  | {
      readonly kind: "anchored";
      readonly targetCNvPrId: number;
      readonly side: ConnectorSide;
      /**
       * Optional position along the side, in [0, 1]. 0 is the left end
       * of n/s edges (resp. top of w/e), 1 the right (resp. bottom),
       * 0.5 the midpoint. Omitted (treated as 0.5) for plain cardinal
       * anchors. Ignored for `center`. OOXML round-trip collapses `t`
       * to the nearest cardinal connection-site index on save.
       */
      readonly t?: number;
    };

export type ConnectorType = "straight" | "elbow" | "curved" | "unsupported";

export type ConnectorEndShape = "none" | "arrow" | "triangle" | "oval";

/**
 * Editor-facing connector dash style. Mirrors PowerPoint's
 * `ST_PresetLineDashVal` 1:1 so we never lose a preset on round-trip.
 *
 * Naming preserves the OOXML token spelling (sysDot, lgDashDotDot, …)
 * so the parser/serializer mapping is the identity function. The
 * legacy short names (`dashed`, `dotted`, `longDash`, `dashDot`) are
 * retained as additional aliases for callers / fixtures that
 * predated the full preset coverage; the serializer collapses them
 * to their canonical OOXML form.
 */
export type ConnectorDashStyle =
  | "solid"
  | "dot"
  | "dash"
  | "lgDash"
  | "dashDot"
  | "lgDashDot"
  | "lgDashDotDot"
  | "sysDash"
  | "sysDot"
  | "sysDashDot"
  | "sysDashDotDot"
  // Legacy short aliases (kept for callers that predate the full enum).
  | "dashed"
  | "dotted"
  | "longDash";

export interface ConnectorStroke {
  /**
   * 6-character RRGGBB hex (no `#`). Always present so the renderer
   * has a sensible fallback even when {@link colorTheme} drives
   * the actual paint at render time.
   */
  readonly color: string;
  /**
   * Theme color reference from `<a:schemeClr val="…"/>`. When set,
   * the serializer emits a `<a:schemeClr>` element instead of the
   * literal `<a:srgbClr>`, preserving theme indirection across
   * round-trips. Common values include `accent1`…`accent6`, `tx1`,
   * `tx2`, `bg1`, `bg2`, `dk1`, `dk2`, `lt1`, `lt2`, and
   * `phClr`/`folHlink`/`hlink` for hyperlinks.
   */
  readonly colorTheme?: string;
  readonly widthEmu: number;
  /** Optional dash pattern. Defaults to "solid" when omitted. */
  readonly dash?: ConnectorDashStyle;
}

/**
 * A connector (`<p:cxnSp>`) is a line that "remembers" what it was
 * anchored to. When either endpoint moves, anchored connectors re-route
 * automatically. The `position`/`size` on `ShapeBase` reflect the
 * derived bounding box of the resolved endpoints — they're rebuilt by
 * the model on every endpoint update, not authored independently.
 */
export interface ConnectorShape extends ShapeBase {
  readonly kind: "connector";
  readonly connectorType: ConnectorType;
  readonly start: ConnectorEndpoint;
  readonly end: ConnectorEndpoint;
  readonly stroke?: ConnectorStroke;
  readonly headEnd?: ConnectorEndShape;
  readonly tailEnd?: ConnectorEndShape;
  /**
   * Optional user-supplied bend offsets for elbow connectors. Each
   * value is a perpendicular displacement (in EMU) added to one of the
   * routed segments — see `routeElbow` for which segment each entry
   * controls. Authoring intent only; ignored for non-elbow types and
   * lossy through OOXML round-trip (we don't currently emit/parse
   * <a:bentConnectorAdjust> overrides).
   */
  readonly waypoints?: ReadonlyArray<number>;
  /** `<p:nvCxnSpPr>` head/tail (sans `<p:cNvPr>` we type). Verbatim. */
  readonly nvCxnSpPrTail: ReadonlyArray<OpaqueXml>;
  /** `<p:spPr>` tail (sans `<a:xfrm>` and `<a:prstGeom>` we rebuild). */
  readonly spPrTail: ReadonlyArray<OpaqueXml>;
}

export interface ShapeBase {
  readonly id: NodeId;
  readonly cNvPrId: number;
  readonly name: string;
  readonly position?: Position;
  readonly size?: Size;
  /**
   * Optional rotation, in degrees, clockwise around the shape's centre.
   * Mirrors `<a:xfrm rot="…">` (which OOXML stores in 60000ths of a
   * degree). `undefined` and `0` are equivalent — the renderer skips
   * the rotation transform in both cases and the serializer omits the
   * attribute, matching what PowerPoint emits for an unrotated shape.
   * Values are normalised to `[0, 360)` on commit but the renderer
   * accepts any finite number.
   */
  readonly rotation?: number;
}

export interface Position {
  readonly xEmu: number;
  readonly yEmu: number;
}

export interface Size {
  readonly cxEmu: number;
  readonly cyEmu: number;
}

export interface TextShape extends ShapeBase {
  readonly kind: "text";
  readonly placeholder?: { type?: string; idx?: number };
  readonly txBody: TextBody;
  readonly nvSpPrTail: ReadonlyArray<OpaqueXml>;
  readonly spPrTail: ReadonlyArray<OpaqueXml>;
  readonly styleRaw?: OpaqueXml;
}

export interface Picture extends ShapeBase {
  readonly kind: "pic";
  readonly mediaRelId: string;
  /** Resolved package-absolute media part path. */
  readonly mediaPartPath: string;
  readonly nvPicPrTail: ReadonlyArray<OpaqueXml>;
  readonly blipFillTail: ReadonlyArray<OpaqueXml>;
  readonly spPrTail: ReadonlyArray<OpaqueXml>;
  readonly styleRaw?: OpaqueXml;
}

/**
 * `<p:graphicFrame>` whose `<a:graphicData>` hosts an `<a:tbl>`.
 * The only typed `graphicFrame` payload — chart and SmartArt frames
 * remain `OpaqueShape`. Cell text is fully typed; cell visual properties
 * (`<a:tcPr>`) and the surrounding `<a:tblPr>` stay opaque.
 */
export interface TableShape extends ShapeBase {
  readonly kind: "table";
  /** Per-column widths in EMU. `<a:tblGrid> <a:gridCol w=…>`. */
  readonly columnWidths: ReadonlyArray<number>;
  /** Row data; rows[i].cells.length === columnWidths.length. */
  readonly rows: ReadonlyArray<TableRow>;
  /** `<a:tblPr>` — verbatim. */
  readonly tblPrRaw?: OpaqueXml;
  /** `<p:nvGraphicFramePr>` tail (sans `<p:cNvPr>` we typed). Verbatim. */
  readonly nvGraphicFramePrTail: ReadonlyArray<OpaqueXml>;
  /** `<a:graphicData @uri>`. Always the table URI for a TableShape. */
  readonly graphicDataUri: string;
}

export interface TableRow {
  readonly id: NodeId;
  /** Row height (EMU). `<a:tr @h>`. */
  readonly height: number;
  readonly cells: ReadonlyArray<TableCell>;
  /** Other `<a:tr>` attributes (e.g. `@thStr`). Verbatim. */
  readonly trAttrs: Readonly<Record<string, string>>;
}

export interface TableCell {
  readonly id: NodeId;
  readonly txBody: TextBody;
  /** `<a:tcPr>`. Verbatim. */
  readonly tcPrRaw?: OpaqueXml;
  /** `<a:tc>` attrs (gridSpan, hMerge, vMerge, …). Verbatim. */
  readonly tcAttrs: Readonly<Record<string, string>>;
}

/**
 * F3: typed chart graphic frame. Position/size are model-driven; the
 * actual chart data lives in a referenced `ChartPart` (resolved via
 * the slide's relationships file).
 */
export interface ChartShape extends ShapeBase {
  readonly kind: "chart";
  /** `<c:chart r:id>` relationship id (relative to the slide rels file). */
  readonly chartRelId: string;
  /** Resolved chart part path, e.g. `ppt/charts/chart1.xml`. */
  readonly chartPartPath: string;
  /** `<p:nvGraphicFramePr>` tail (sans `<p:cNvPr>` we typed). Verbatim. */
  readonly nvGraphicFramePrTail: ReadonlyArray<OpaqueXml>;
  /** `<a:graphicData @uri>`. Always the chart uri for a ChartShape. */
  readonly graphicDataUri: string;
}

/**
 * `<p:graphicFrame>` whose `<a:graphicData>` hosts a `<p:oleObj>` with
 * `progId` matching `Excel.Sheet.*` (typically `Excel.Sheet.12`).
 *
 * This is the "live" Excel embed: double-click in PowerPoint pops the
 * embedded `.xlsx` open in Excel for editing. We type just enough of
 * the frame to (a) detect the embed, (b) keep the embedded workbook
 * + preview image rels round-trippable, and (c) author fresh embeds
 * via `pptx:insert-spreadsheet`. Everything else (the original `<p:pic>`
 * preview subtree, `<p:embed>` flags, OLE follow content) is captured
 * opaquely so existing files survive byte-identical no-touch saves.
 */
export interface OleSpreadsheetShape extends ShapeBase {
  readonly kind: "ole-spreadsheet";
  /** `<p:oleObj r:id>` — relationship id pointing at the embedded part. */
  readonly oleRelId: string;
  /** Resolved package-absolute path of the embedded `.xlsx` (or `.bin`). */
  readonly embeddingPartPath: string;
  /** `<p:oleObj progId>` — e.g. `Excel.Sheet.12`. */
  readonly progId: string;
  /** Embedded part kind: `xlsx` (true Excel package) or `bin` (legacy CFB). */
  readonly embeddingKind: "xlsx" | "bin";
  /** `<p:pic>` preview image rel (absolute media path), if known. */
  readonly previewMediaRelId?: string;
  readonly previewMediaPartPath?: string;
  /** Captured `<p:oleObj>` attribute bag (spid, name, imgW, imgH, …). */
  readonly oleObjAttrs: Readonly<Record<string, string>>;
  /**
   * Raw `<p:oleObj>` child subtree (`<p:embed>`, `<p:link>`, `<p:pic>`,
   * follow content). Preserved verbatim so untouched files round-trip
   * byte-identically.
   */
  readonly oleObjChildrenRaw: ReadonlyArray<OpaqueXml>;
  /** `<p:nvGraphicFramePr>` tail (sans `<p:cNvPr>`). Verbatim. */
  readonly nvGraphicFramePrTail: ReadonlyArray<OpaqueXml>;
  /** `<a:graphicData @uri>`. Always the OLE uri for an OleSpreadsheetShape. */
  readonly graphicDataUri: string;
}

export interface GroupShape extends ShapeBase {
  readonly kind: "group";
  /** Captured <a:chOff>+<a:chExt> slice for verbatim re-emit. */
  readonly chOffExtRaw: ReadonlyArray<OpaqueXml>;
  readonly children: ReadonlyArray<Shape>;
  readonly grpSpPrTail: ReadonlyArray<OpaqueXml>;
  readonly nvGrpSpPrTail: ReadonlyArray<OpaqueXml>;
}

export interface OpaqueShape extends ShapeBase {
  readonly kind: "opaque";
  readonly tag: string;
  readonly raw: OpaqueXml;
}

// ─── Text body ────────────────────────────────────────────────────────────

export interface TextBody {
  readonly bodyPrRaw?: OpaqueXml;
  readonly lstStyleRaw?: OpaqueXml;
  readonly paragraphs: ReadonlyArray<TextParagraph>;
}

export interface TextParagraph {
  readonly id: NodeId;
  readonly properties: TextParagraphProperties;
  readonly runs: ReadonlyArray<TextRun>;
  readonly endParaRPrRaw?: OpaqueXml;
}

export interface TextParagraphProperties {
  readonly level?: number;
  readonly alignment?: "left" | "center" | "right" | "justify";
  /** <a:pPr> attrs we don't introspect. */
  readonly opaqueAttrs?: Readonly<Record<string, string>>;
  /** Children of <a:pPr> we don't model. */
  readonly opaqueChildren?: ReadonlyArray<OpaqueXml>;
}

export interface TextRun {
  readonly id: NodeId;
  readonly properties: TextRunProperties;
  readonly text: string;
  readonly isLineBreak?: boolean;
}

export interface TextRunProperties {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean | string;
  readonly strike?: boolean;
  readonly fontSizeHundredths?: number;
  readonly fontFamily?: string;
  /**
   * East Asian typeface from `<a:ea typeface="…"/>`. Required for
   * CJK presentations. Editing {@link fontFamily} preserves this
   * slot — Asian glyphs continue to render in their dedicated face.
   */
  readonly fontFamilyEastAsia?: string;
  /**
   * Complex-script typeface from `<a:cs typeface="…"/>` (Arabic,
   * Hebrew, Indic, etc.).
   */
  readonly fontFamilyComplexScript?: string;
  /**
   * Symbol typeface from `<a:sym typeface="…"/>`. Used for special
   * glyph maps (Wingdings-class fonts).
   */
  readonly fontFamilySymbol?: string;
  /**
   * Theme reference for the Latin face. Carried verbatim — the
   * canonical PowerPoint values are `+mj-lt` (major Latin) and
   * `+mn-lt` (minor Latin); we don't validate so unknown values
   * round-trip too.
   */
  readonly fontFamilyLatinTheme?: string;
  /** Theme reference for the East Asian face (`+mj-ea` / `+mn-ea`). */
  readonly fontFamilyEastAsiaTheme?: string;
  /** Theme reference for the complex-script face (`+mj-cs` / `+mn-cs`). */
  readonly fontFamilyComplexScriptTheme?: string;
  readonly color?: string;
  /**
   * Character highlight (background colour behind glyphs). Stored as
   * a 6-character RRGGBB hex string, matching `color`. Round-trips
   * through `<a:highlight><a:srgbClr val="…"/></a:highlight>`.
   * `undefined` means "inherit"; the empty string is rejected by the
   * format-text command (use `undefined` to clear via patch).
   */
  readonly highlight?: string;
  /** <a:rPr> attrs we don't introspect (lang, dirty, kern, …). */
  readonly opaqueAttrs?: Readonly<Record<string, string>>;
  /** Children of <a:rPr> we don't model (schemeClr, hlinkClick, …). */
  readonly opaqueChildren?: ReadonlyArray<OpaqueXml>;
}

// ─── OpaqueXml ────────────────────────────────────────────────────────────

export interface OpaqueXml {
  readonly tag: string;
  readonly attrs: Readonly<Record<string, string>>;
  /** Original prefixed attrs (e.g. "@_id") for byte-faithful re-emit. */
  readonly rawAttrs: Readonly<Record<string, string>>;
  readonly subtree: ReadonlyArray<unknown>;
}
