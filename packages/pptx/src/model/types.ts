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
  readonly layouts: ReadonlyMap<string, OpaquePart>;
  readonly theme: ReadonlyMap<string, OpaquePart>;
  /**
   * Resolved color scheme from the first theme part (`a:clrScheme`). Used
   * by the renderer to draw `<a:schemeClr>` references with the correct
   * palette. Falls back to `DEFAULT_THEME` if no theme part is present
   * or the scheme can't be parsed.
   */
  readonly themeDefault: ThemeColorScheme;
  readonly notesSlides: ReadonlyMap<string, OpaquePart>;
  readonly media: ReadonlyMap<string, MediaPart>;
  /**
   * F3: typed chart parts keyed by part path. Each `ChartShape` resolves
   * its `chartPartPath` here. Anything in the chart XML we don't model
   * is preserved as `OpaqueXml` inside the part.
   */
  readonly charts: ReadonlyMap<string, ChartPart>;
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

export interface MediaPart {
  readonly partPath: string;
  readonly contentType: string;
  readonly bytes: Uint8Array;
  readonly sha256: string;
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
}

// ─── Slide ────────────────────────────────────────────────────────────────

export interface Slide {
  readonly id: NodeId;
  readonly partPath: string;
  readonly slideId: number;
  readonly relId: string;
  readonly layoutPartPath?: string;
  readonly notesSlidePartPath?: string;
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

export type EntranceEffect = "appear" | "fade" | "fly-in" | "wipe";

/**
 * F4: typed per-shape entrance animation. Targets a shape by `cNvPrId`
 * (which is what `<p:spTgt @spid>` references). Only the four named
 * effects are typed; anything else stays in `timingTailRaw` verbatim.
 */
export interface EntranceAnimation {
  readonly id: NodeId;
  readonly targetCNvPrId: number;
  readonly effect: EntranceEffect;
  /** `<p:cTn @dur>` in milliseconds. */
  readonly durationMs?: number;
  /** Trigger order in the main entrance sequence (0-based). */
  readonly order: number;
}

// ─── Shapes ───────────────────────────────────────────────────────────────

export type Shape = TextShape | Picture | TableShape | ChartShape | GroupShape | OpaqueShape;

export type ShapeKind = Shape["kind"];

export interface ShapeBase {
  readonly id: NodeId;
  readonly cNvPrId: number;
  readonly name: string;
  readonly position?: Position;
  readonly size?: Size;
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
  readonly color?: string;
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
