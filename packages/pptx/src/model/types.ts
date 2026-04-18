/**
 * PPTX in-memory model.
 *
 * See spec/pptx/document-model.md and spec/pptx/ooxml-mapping.md.
 */

import type { DocumentSnapshot, NodeId, ooxml as _ooxml } from "@officeai/core";

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
  readonly notesSlides: ReadonlyMap<string, OpaquePart>;
  readonly media: ReadonlyMap<string, MediaPart>;
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

// ─── Slide ────────────────────────────────────────────────────────────────

export interface Slide {
  readonly id: NodeId;
  readonly partPath: string;
  readonly slideId: number;
  readonly relId: string;
  readonly layoutPartPath?: string;
  readonly notesSlidePartPath?: string;
  readonly shapes: ReadonlyArray<Shape>;
  readonly slideOpaqueTail: ReadonlyArray<OpaqueXml>;
  readonly slideRootAttrs: Readonly<Record<string, string>>;
  readonly cSldAttrs: Readonly<Record<string, string>>;
  /** Head children of <p:cSld> + <p:spTree> we don't introspect. */
  readonly spTreeHead: ReadonlyArray<OpaqueXml>;
  /** Optional <p:bg> captured if present (sits inside <p:cSld>). */
  readonly cSldHead: ReadonlyArray<OpaqueXml>;
}

// ─── Shapes ───────────────────────────────────────────────────────────────

export type Shape = TextShape | Picture | GroupShape | OpaqueShape;

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
