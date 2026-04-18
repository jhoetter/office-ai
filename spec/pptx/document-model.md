# PPTX — In-Memory Model

## Roots

```typescript
import type { NodeId, DocumentSnapshot } from "@officeai/core";

export interface PptxSnapshot extends DocumentSnapshot<PptxPresentation> {
  readonly format: "pptx";
  /**
   * Per-part dirty flags. Cleared by the parser; set by command handlers
   * via the bus' producer step. Drives the serializer's per-part decision
   * (re-emit from typed model vs copy bytes from container cache).
   */
  readonly dirty: PptxDirty;
  /**
   * Attached at parse time so the serializer can reach untouched-part
   * bytes. Not part of the value identity (excluded from snapshot
   * equality in tests).
   */
  readonly container: import("@officeai/core").OoxmlContainer;
}

export interface PptxDirty {
  readonly presentation: boolean;       // ppt/presentation.xml
  readonly slides: ReadonlySet<string>; // slide part paths e.g. "ppt/slides/slide3.xml"
  readonly notesSlides: ReadonlySet<string>;
  readonly masters: ReadonlySet<string>;
  readonly layouts: ReadonlySet<string>;
  readonly theme: ReadonlySet<string>;
  readonly media: ReadonlySet<string>;
  readonly relationships: ReadonlySet<string>; // rels part paths
  readonly contentTypes: boolean;       // [Content_Types].xml
}

export interface PptxPresentation {
  readonly id: NodeId;
  /** Slides in display order. The natural address of a slide is its index. */
  readonly slides: ReadonlyArray<Slide>;
  /** Slide canvas size in EMU (typical: 9144000×6858000 or 12192000×6858000). */
  readonly slideSize: SlideSize;
  /** Notes canvas size in EMU. */
  readonly notesSize?: SlideSize;
  /**
   * Slide masters keyed by part path. Treated as opaque this session —
   * we read them only for placeholder / theme resolution at render time.
   */
  readonly masters: ReadonlyMap<string, OpaquePart>;
  /** Slide layouts keyed by part path. Opaque, render-time only. */
  readonly layouts: ReadonlyMap<string, OpaquePart>;
  /** Theme parts keyed by part path. Opaque, render-time only. */
  readonly theme: ReadonlyMap<string, OpaquePart>;
  /** Notes slides keyed by part path. Opaque, attached to a slide via rels. */
  readonly notesSlides: ReadonlyMap<string, OpaquePart>;
  /** Media parts keyed by part path. Carry SHA-256 for dedup. */
  readonly media: ReadonlyMap<string, MediaPart>;
  /** Original `<p:presentation>` root attributes. Re-emitted verbatim. */
  readonly presentationRootAttrs: Readonly<Record<string, string>>;
  /**
   * Tail children of `<p:presentation>` we don't introspect
   * (defaultTextStyle, custShowLst, etc.). Re-emitted verbatim.
   */
  readonly presentationOpaqueTail: ReadonlyArray<OpaqueXml>;
  /**
   * Stable counter for minting new slide-ids / rIds / part paths. The
   * parser seeds this from the highest values it sees; commands bump it
   * monotonically and never re-use ids.
   */
  readonly idGen: PptxIdGen;
}

export interface SlideSize {
  readonly cxEmu: number; // width
  readonly cyEmu: number; // height
  /** OOXML slide-size type: "screen4x3", "screen16x9", "custom", etc. */
  readonly type?: string;
}

export interface PptxIdGen {
  readonly nextSlideId: number;       // <p:sldId @id>; PowerPoint starts at 256
  readonly nextRelId: number;         // global per-rels-part counter; serializer scopes per rels file
  readonly nextSlidePartIndex: number;// next "slideN.xml" suffix
  readonly nextMediaPartIndex: number;// next "imageN.ext" suffix
}

export interface OpaquePart {
  readonly partPath: string;
  /**
   * Captured for the renderer's read-only inspection. The serializer
   * does NOT re-stringify this — it copies bytes from the container
   * cache. (`raw` is for code paths that need to *read* the part
   * structurally, e.g. theme color resolution.)
   */
  readonly raw: OpaqueXml;
}

export interface MediaPart {
  readonly partPath: string;
  readonly contentType: string; // e.g. "image/png"
  readonly bytes: Uint8Array;   // referenced by inserted images for re-emit
  readonly sha256: string;      // dedup key
}
```

## Slide

```typescript
export interface Slide {
  readonly id: NodeId;
  /** OOXML part path, e.g. "ppt/slides/slide1.xml". Stable across reorders. */
  readonly partPath: string;
  /** OOXML slide id from `<p:sldId @id>`. Stable across reorders. */
  readonly slideId: number;
  /** Relationship id from `<p:sldId @r:id>`. Resolves to partPath. */
  readonly relId: string;
  /** Reference to the layout this slide uses, by partPath. */
  readonly layoutPartPath?: string;
  /** Reference to the attached notes slide partPath, if present. */
  readonly notesSlidePartPath?: string;
  /** Top-level shapes inside `<p:cSld><p:spTree>`, in z-order. */
  readonly shapes: ReadonlyArray<Shape>;
  /**
   * Tail children of `<p:sld>` we don't introspect (clrMapOvr, transition,
   * timing, extLst, …). Re-emitted verbatim on serialize.
   */
  readonly slideOpaqueTail: ReadonlyArray<OpaqueXml>;
  /**
   * Original root attributes of `<p:sld>`. Re-emitted verbatim.
   */
  readonly slideRootAttrs: Readonly<Record<string, string>>;
  /**
   * Original `<p:cSld>` attributes. Re-emitted verbatim.
   */
  readonly cSldAttrs: Readonly<Record<string, string>>;
  /**
   * Original `<p:spTree>` head (NonVisualGroupShapeProperties +
   * GroupShapeProperties for the spTree). Re-emitted verbatim.
   */
  readonly spTreeHead: ReadonlyArray<OpaqueXml>;
}
```

## Shapes

```typescript
export type Shape = TextShape | Picture | TableShape | GroupShape | OpaqueShape;

export type ShapeKind = Shape["kind"];

export interface ShapeBase {
  readonly id: NodeId;
  /** OOXML `<p:cNvPr @id>`. Per-slide-unique integer minted by PowerPoint. */
  readonly cNvPrId: number;
  /** OOXML `<p:cNvPr @name>`. Free-text label. */
  readonly name: string;
  /** Resolved position (EMU). May be undefined if inherited from layout. */
  readonly position?: Position;
  /** Resolved size (EMU). May be undefined if inherited from layout. */
  readonly size?: Size;
}

export interface Position { readonly xEmu: number; readonly yEmu: number; }
export interface Size     { readonly cxEmu: number; readonly cyEmu: number; }

/** A `<p:sp>` shape carrying a `<p:txBody>`. P0-editable text + xfrm. */
export interface TextShape extends ShapeBase {
  readonly kind: "text";
  /** Placeholder type if the shape is a placeholder (`<p:ph type="…" idx="…">`). */
  readonly placeholder?: { type?: string; idx?: number };
  /** Typed text body. May be empty (txBody with a single empty paragraph). */
  readonly txBody: TextBody;
  /**
   * Captured `<p:nvSpPr>` (minus the parts we typed: cNvPr id+name, ph) so
   * the serializer can re-emit non-introspected metadata verbatim.
   */
  readonly nvSpPrTail: ReadonlyArray<OpaqueXml>;
  /**
   * Captured `<p:spPr>` children other than `<a:xfrm>` (geometry,
   * fill, line, etc.). Re-emitted verbatim. xfrm is rebuilt from
   * position+size.
   */
  readonly spPrTail: ReadonlyArray<OpaqueXml>;
  /**
   * Captured `<p:style>` if present. Re-emitted verbatim.
   */
  readonly styleRaw?: OpaqueXml;
}

/** A `<p:pic>` picture shape, referencing a media part by its slide-rel. */
export interface Picture extends ShapeBase {
  readonly kind: "pic";
  /** Slide-rel id (`r:embed` on `<a:blip>`). Resolves to a `MediaPart`. */
  readonly mediaRelId: string;
  /**
   * The dereferenced media part path (cached on the model so commands
   * don't have to rewalk the rels graph). Equal to
   * `relationships.get(slideRelsPath)!.find(r => r.id === mediaRelId)!.target`,
   * normalized to a package-absolute path.
   */
  readonly mediaPartPath: string;
  /**
   * Captured `<p:nvPicPr>` tail and `<p:blipFill>` (sans the `<a:blip>`
   * we typed) and `<p:spPr>` tail (sans `<a:xfrm>`). Re-emitted verbatim.
   */
  readonly nvPicPrTail: ReadonlyArray<OpaqueXml>;
  readonly blipFillTail: ReadonlyArray<OpaqueXml>;
  readonly spPrTail: ReadonlyArray<OpaqueXml>;
  readonly styleRaw?: OpaqueXml;
}

/**
 * A `<p:graphicFrame>` whose `<a:graphicData>` hosts an `<a:tbl>`
 * (the only `graphicFrame` payload typed in F2 — charts and SmartArt
 * stay `OpaqueShape`).
 *
 * Cell text is fully typed (paragraphs → runs) so `set-cell-text` can
 * roundtrip cleanly. Cell visual properties (`<a:tcPr>`, fills, borders)
 * stay opaque, mirroring the conservative posture taken for picture
 * `<p:blipFill>` tails.
 */
export interface TableShape extends ShapeBase {
  readonly kind: "table";
  /** `<a:tblGrid>` column widths (EMU). One entry per column. */
  readonly columnWidths: ReadonlyArray<number>;
  /** Row data; rows[i].cells.length === columnWidths.length. */
  readonly rows: ReadonlyArray<TableRow>;
  /** `<a:tblPr>` — visual style ref, fills, banded settings. Verbatim. */
  readonly tblPrRaw?: OpaqueXml;
  /**
   * `<p:nvGraphicFramePr>` and `<p:graphicFramePr>` heads — verbatim
   * (cNvPr id+name are the only fields typed via `ShapeBase`).
   */
  readonly nvGraphicFramePrTail: ReadonlyArray<OpaqueXml>;
  /** `<a:graphicData @uri>` — always `"…/drawingml/2006/table"` for tables. */
  readonly graphicDataUri: string;
}

export interface TableRow {
  readonly id: NodeId;
  /** Row height (EMU). `<a:tr @h>`. */
  readonly height: number;
  readonly cells: ReadonlyArray<TableCell>;
  /** Anything on `<a:tr>` we don't model (e.g. `@thStr`). Verbatim. */
  readonly trAttrs: Readonly<Record<string, string>>;
}

export interface TableCell {
  readonly id: NodeId;
  /** Typed cell text body. Same shape as `TextShape.txBody`. */
  readonly txBody: TextBody;
  /** `<a:tcPr>` — borders, fill, anchor. Verbatim. */
  readonly tcPrRaw?: OpaqueXml;
  /**
   * `<a:tc>` attributes (`@gridSpan`, `@rowSpan`, `@hMerge`, `@vMerge`).
   * Re-emitted verbatim. F2 commands do NOT model merges yet — splits/
   * merges raise `not-applicable`.
   */
  readonly tcAttrs: Readonly<Record<string, string>>;
}

/** A `<p:grpSp>` group. Children are typed but P0 only moves the group as a unit. */
export interface GroupShape extends ShapeBase {
  readonly kind: "group";
  /**
   * Group child coordinate-system offset+extent (`<a:chOff>`/`<a:chExt>`).
   * Re-emitted verbatim — moving a group changes its xfrm `off/ext`,
   * not chOff/chExt (PowerPoint's behavior).
   */
  readonly chOffExtRaw: OpaqueXml;
  readonly children: ReadonlyArray<Shape>;
  readonly grpSpPrTail: ReadonlyArray<OpaqueXml>;
}

/**
 * Anything not above (cxnSp, graphicFrame, contentPart, AlternateContent, …).
 * Position/size parsed best-effort from a leading `<a:xfrm>` if present;
 * otherwise both are undefined and the renderer falls back to layout/master.
 */
export interface OpaqueShape extends ShapeBase {
  readonly kind: "opaque";
  /** Tag of the opaque element so renderers can switch on type. */
  readonly tag: string; // "p:cxnSp" | "p:graphicFrame" | "mc:AlternateContent" | …
  /** The full element subtree, captured for verbatim re-emit. */
  readonly raw: OpaqueXml;
}
```

## Text body

```typescript
export interface TextBody {
  /** The OOXML `<a:bodyPr>` carrier (margins, autofit, anchor). Opaque this session. */
  readonly bodyPrRaw?: OpaqueXml;
  /** The OOXML `<a:lstStyle>` carrier. Opaque this session. */
  readonly lstStyleRaw?: OpaqueXml;
  readonly paragraphs: ReadonlyArray<TextParagraph>;
}

export interface TextParagraph {
  readonly id: NodeId;
  /** `<a:pPr>` carrier — typed properties + an opaque tail for unmodeled children. */
  readonly properties: TextParagraphProperties;
  readonly runs: ReadonlyArray<TextRun>;
  /** A trailing `<a:endParaRPr>` carrier preserved verbatim. */
  readonly endParaRPrRaw?: OpaqueXml;
}

export interface TextParagraphProperties {
  /** `<a:pPr @lvl>` indent level (0..8). */
  readonly level?: number;
  /** Alignment: l/ctr/r/just (mapped to "left"/"center"/"right"/"justify"). */
  readonly alignment?: "left" | "center" | "right" | "justify";
  /** Anything in `<a:pPr>` we don't model. Re-emitted verbatim. */
  readonly opaqueProps?: ReadonlyArray<OpaqueXml>;
}

export interface TextRun {
  readonly id: NodeId;
  readonly properties: TextRunProperties;
  /** Plain text from `<a:t>`. Line breaks (`<a:br>`) are modeled as separate runs. */
  readonly text: string;
  /** Whether this run came from a `<a:br>` (renders as line break, no text). */
  readonly isLineBreak?: boolean;
}

export interface TextRunProperties {
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly underline?: boolean | string;
  readonly strike?: boolean;
  /** EMU font size (OOXML stores in hundredths of a point: `sz="3200"` → 32pt). */
  readonly fontSizeHundredths?: number;
  /** Typeface family from `<a:latin typeface="…"/>`. East-asian/cs preserved opaquely. */
  readonly fontFamily?: string;
  /** Resolved color (hex without `#`) when `<a:srgbClr>` is used directly. */
  readonly color?: string;
  /**
   * Anything in `<a:rPr>` we don't model (theme color references, hyperlinks,
   * lang, smartTag, lineSpacing, …). Re-emitted verbatim.
   */
  readonly opaqueProps?: ReadonlyArray<OpaqueXml>;
}
```

## OpaqueXml

The shared `OpaqueXml` type — already used by DOCX — is reused as-is:

```typescript
export interface OpaqueXml {
  /** Tag including namespace prefix, e.g. "p:sp", "a:xfrm". */
  readonly tag: string;
  readonly attrs: Readonly<Record<string, string>>;
  /**
   * Verbatim slice from the parser's preserveOrder representation. The
   * serializer re-emits this slice byte-for-byte.
   */
  readonly subtree: ReadonlyArray<unknown>;
}
```

## Identity & freezing

- The parser mints `NodeId`s for every typed node (`id` field).
- `cNvPrId` / `slideId` / `relId` are PowerPoint's identifiers and are stable across reorders.
- After `delete-slide` the freed `slideId` and `relId` are NOT re-used (consistent with PowerPoint).
- `freezeSnapshot(snapshot)` from `@officeai/core` deep-freezes the snapshot in dev/test for catch-on-write protection.

## What is NOT in the model

- Slide masters / layouts / theme contents — opaque parts only (we do read them at render time, but we never type them).
- Notes-slide contents — opaque per-slide.
- Animation timing tree, transitions, color map overrides — captured under `slideOpaqueTail`.
- Connector geometry, `graphicFrame` charts/SmartArt/tables — opaque shapes only.
