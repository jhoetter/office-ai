# PPTX — Agent Commands

> Complete typed list. Ten commands ship in P0/P1 of this PPTX phase
> (per [`prompt.md`](../../prompt.md) lines 414–425, plus `format-text`
> from line 420).

## Common types

```typescript
import type { Command, NodeId } from "@officeai/core";

export interface PptxCommandBase<TType extends string, TPayload> extends Command<TType, TPayload> {
  // type & payload from Command<>; convenience-typed here per command.
}

/** A flat-text range inside a TextShape's TextBody. */
export interface PptxTextRange {
  /** 0-based paragraph index inside the txBody. */
  readonly paragraph: number;
  /** Inclusive char offset within the paragraph's flat text. */
  readonly start: number;
  /** Exclusive char offset within the paragraph's flat text.
   *  start === end is legal but a no-op (returns invalid-payload). */
  readonly end: number;
}

export interface TextFormatPayload {
  /** Toggle marks; absence = leave unchanged. */
  bold?: boolean;
  italic?: boolean;
  underline?: boolean | string;
  strike?: boolean;
  fontFamily?: string;
  /** Hundredths of a point — what OOXML stores in `<a:rPr @sz>`. */
  fontSizeHundredths?: number;
  /** Hex without leading "#". */
  color?: string;
}
```

All payloads are JSON-serializable. All shape addresses are
`(slideIndex, shapeId)` — `slideIndex` is the user-facing 0-based
ordinal, `shapeId` is the typed-model `NodeId`.

## Commands (P0)

### `pptx:add-slide`

```typescript
type AddSlidePayload = {
  /** 0-based insert position. Defaults to `slides.length` (append). */
  at?: number;
  /**
   * Layout part path to clone placeholders from. If omitted, the slide
   * starts blank (no shapes; only `<p:cSld><p:spTree>` group head).
   * The layout MUST exist in `presentation.layouts`.
   */
  layoutPartPath?: string;
};
```

Mint a fresh slide:

1. Pick `nextSlideId` and `nextSlidePartIndex` from `idGen`. Bump `idGen`.
2. Pick `nextRelId` for `ppt/_rels/presentation.xml.rels`. Bump.
3. Construct an empty (or placeholder-cloned) `Slide` with empty
   `slideOpaqueTail`, default root attrs, the new `partPath`
   (`ppt/slides/slide{N}.xml`), the new `slideId`, the new `relId`, and
   the requested `layoutPartPath` (if provided).
4. Insert into `slides` at `at`.
5. Add `<Relationship Id="rIdN" Type=".../slide" Target="slides/slide{N}.xml"/>` to
   `ppt/_rels/presentation.xml.rels`.
6. If a layout was specified, add a layout relationship in the new
   `ppt/slides/_rels/slide{N}.xml.rels` file.
7. Add `<Override PartName="/ppt/slides/slide{N}.xml"
   ContentType=".../slide+xml"/>` to `[Content_Types].xml`.
8. Set `dirty.presentation = true`, `dirty.slides.add(newPath)`,
   `dirty.relationships.add(presRelsPath)`,
   `dirty.relationships.add(newSlideRelsPath)` if layout was set,
   `dirty.contentTypes = true`.

Errors:

- `invalid-position` — `at < 0` or `at > slides.length`.
- `unknown-target` — `layoutPartPath` provided but not present in `presentation.layouts`.

Diff: `{ kind: "node-inserted", path: ["slides", at], summary: "slide" }`.

### `pptx:delete-slide`

```typescript
type DeleteSlidePayload = { slideIndex: number };
```

Drop the slide:

1. Resolve the slide at `slideIndex`. `unknown-target` if OOB.
2. Remove from `slides[]`.
3. Remove its `<p:sldId>` entry from `presentation.xml`.
4. Remove its `<Relationship>` from `ppt/_rels/presentation.xml.rels`.
5. Remove its `<Override>` from `[Content_Types].xml`.
6. If the slide had `notesSlidePartPath`, drop the notes-slide part,
   its rels file, and its `<Override>` too.
7. Mark the slide part path (and its rels file) as `removedParts`.
8. Set `dirty.presentation = true`, `dirty.relationships.add(presRelsPath)`,
   `dirty.contentTypes = true`.

Errors:

- `unknown-target` — `slideIndex < 0` or `>= slides.length`.

Diff: `{ kind: "node-deleted", path: ["slides", slideIndex], summary: "slide" }`.

### `pptx:duplicate-slide`

```typescript
type DuplicateSlidePayload = { slideIndex: number };
```

Deep-clone the slide at `slideIndex` and insert at `slideIndex + 1`:

1. Pick `nextSlideId`, `nextRelId`, `nextSlidePartIndex` for the clone.
2. Deep-clone the typed `Slide`. Re-mint every `cNvPrId` inside the
   clone to be unique within the new slide (max-existing-on-clone + 1
   per shape kind; per-slide scope, so cross-slide collisions are fine).
3. Copy the source slide's `_rels/slide{srcN}.xml.rels` file content to
   `_rels/slide{newN}.xml.rels` (so picture references continue to
   resolve to the same media parts — no media duplication).
4. Add `<p:sldId>` and `<Relationship>` in the right places.
5. Add `<Override>` for the new slide.
6. Mark `dirty.presentation`, both rels parts, content-types, and the
   new slide part.

Errors:

- `unknown-target` — `slideIndex` out of range.

Diff: `{ kind: "node-inserted", path: ["slides", slideIndex + 1], summary: "slide (duplicate)" }`.

### `pptx:move-slide`

```typescript
type MoveSlidePayload = { from: number; to: number };
```

Reorder slides:

1. Validate `from` and `to` are within `[0, slides.length)`.
2. If `from === to`: no-op (returns success but bumps revision so the
   bus history records the attempt).
3. Splice `slides[]` to put the slide currently at `from` at `to`.
4. Re-emit `<p:sldIdLst>` with the new order. **No part renaming, no
   rId changes.**
5. Set `dirty.presentation = true`.

Errors:

- `invalid-position` — either index OOB.

Diff: `{ kind: "node-moved", path: from, to, summary: "slide" }`.

### `pptx:set-text`

```typescript
type SetTextPayload = {
  slideIndex: number;
  shapeId: NodeId;
  /** New plain-text content of the shape's TextBody.
   *  "\n" separates paragraphs. */
  text: string;
};
```

Replace the text of a `TextShape`:

1. Resolve the shape via `(slideIndex, shapeId)`. `unknown-target` if
   missing or not a `TextShape` (Picture / Group / Opaque shapes
   reject with `not-applicable`).
2. Split `payload.text` by `\n` into paragraph strings.
3. For each paragraph, build a `TextParagraph` whose `properties` are
   inherited from the existing first paragraph's `properties` (so
   alignment/level survive a text-only edit). Each paragraph gets a
   single `TextRun` whose `properties` are inherited from the existing
   first run's `properties`. If the original was empty, start from
   default empty properties.
4. Preserve `bodyPrRaw`, `lstStyleRaw`, and the original first
   paragraph's `endParaRPrRaw`.
5. Set `dirty.slides.add(slidePartPath)`.

Errors:

- `unknown-target` — slide or shape not found.
- `not-applicable` — shape is not a `TextShape`.

Diff: `{ kind: "node-updated", path: ["slides", slideIndex, "shapes", shapeIdx, "txBody"], field: "text", summary: "text" }`.

### `pptx:set-position`

```typescript
type SetPositionPayload = {
  slideIndex: number;
  shapeId: NodeId;
  /** Position in EMU. */
  x: number;
  y: number;
};
```

Set a shape's `<a:off>`:

1. Resolve the shape. `unknown-target` if missing.
2. Validate `x`, `y` are finite integers (we round non-integers).
3. Replace `shape.position` with `{ xEmu: x, yEmu: y }`.
4. Set `dirty.slides.add(slidePartPath)`.

Works on `TextShape`, `Picture`, `GroupShape`. For `OpaqueShape` we
**reject `not-applicable`** because the shape's geometry is held inside
its `raw` slice and we don't introspect arbitrary opaque geometry this
session.

Errors:

- `unknown-target` — slide or shape not found.
- `invalid-payload` — `x` or `y` not finite.
- `not-applicable` — opaque shape.

Diff: `{ kind: "node-updated", path: ..., field: "position", summary: "(${x},${y})" }`.

### `pptx:set-size`

```typescript
type SetSizePayload = {
  slideIndex: number;
  shapeId: NodeId;
  width: number;   // EMU
  height: number;  // EMU
};
```

Mirror of `set-position` for `<a:ext>`. `width > 0`, `height > 0`
required (`invalid-payload` otherwise). Same shape-kind eligibility
rules.

## Commands (P1)

### `pptx:format-text`

```typescript
type FormatTextPayload = {
  slideIndex: number;
  shapeId: NodeId;
  range: PptxTextRange;
  format: TextFormatPayload;
};
```

Apply `format` to every run intersecting `range` inside the
`TextShape`'s body:

1. Resolve the shape (`TextShape` only; reject `not-applicable`).
2. Validate `range.paragraph` is in bounds, `range.start <= range.end`,
   and `range.end <= paragraph.flatTextLength`.
3. Walk the runs of `paragraph.runs`. Split runs at `range.start` and
   `range.end` boundaries (mirrors DOCX's `format-range` algorithm).
4. For every run fully inside the range, merge `format` into
   `properties` (toggle semantics: absence = leave unchanged).
5. Re-coalesce adjacent runs whose `properties` are now equal.
6. Set `dirty.slides.add(slidePartPath)`.

Errors:

- `unknown-target` — slide or shape not found.
- `not-applicable` — shape is not a `TextShape`.
- `invalid-payload` — range invalid or paragraph index out of range.

### `pptx:insert-image`

```typescript
type InsertImagePayload = {
  slideIndex: number;
  data: Uint8Array | ArrayBuffer;
  mimeType: string; // "image/png" | "image/jpeg" | "image/gif" | "image/bmp" | "image/tiff" | "image/webp" | "image/svg+xml"
  x: number;        // EMU
  y: number;        // EMU
  width: number;    // EMU
  height: number;   // EMU
  altText?: string;
  name?: string;    // <p:cNvPr name>; defaults to "Picture {cNvPrId}"
};
```

Insert a `<p:pic>` on the target slide:

1. Validate `data` non-empty, `width > 0`, `height > 0`, MIME
   recognized.
2. Compute `SHA-256(data)`. Search `presentation.media` for an existing
   `MediaPart` with the same digest:
   - Hit → reuse `mediaPartPath`. Try to reuse a `Type=".../image"`
     relationship in the slide's rels graph that points to that path.
     Mint a new rel only if no such rel exists yet on this slide.
   - Miss → mint `ppt/media/image{N}.{ext}` (N = `idGen.nextMediaPartIndex`),
     register a new `MediaPart`, set `dirty.media.add(newPath)`.
     Mint a new rel `rId{M}` in the slide's rels file pointing to it.
     If `[Content_Types].xml` does not yet have a `<Default>` for the
     extension, add one and set `dirty.contentTypes`.
3. Mint `cNvPrId = max(existing on slide) + 1`.
4. Build a typed `Picture` shape with the requested position/size.
5. Append to `slide.shapes` (z-order: top).
6. Set `dirty.slides.add(slidePartPath)`,
   `dirty.relationships.add(slideRelsPath)` if a new rel was minted.

Errors:

- `unknown-target` — `slideIndex` out of range.
- `invalid-payload` — empty data, non-positive size, unsupported MIME.

Diff: emits `node-inserted` for the picture; if a new media part was
added, also emits a `part-added` entry whose `path` is `[mediaPartPath]`.

### `pptx:add-text-box`

```typescript
type AddTextBoxPayload = {
  slideIndex: number;
  text: string;
  x: number; y: number; width: number; height: number; // EMU
  name?: string;
};
```

Append a fresh `<p:sp>` text box:

1. Validate position/size finite; `width > 0`, `height > 0`.
2. Mint `cNvPrId` (`max(existing on slide) + 1`).
3. Build `TextShape` with:
   - `placeholder = undefined`
   - `nvSpPrTail = [{ "p:cNvSpPr": [{ "a:spLocks": [], ":@": { "@_noGrp": "1" } }], ":@": {} }, { "p:nvPr": [], ":@": {} }]`
   - `spPrTail = [{ "a:prstGeom": [{ "a:avLst": [] }], ":@": { "@_prst": "rect" } }, { "a:noFill": [], ":@": {} }]`
   - `txBody.bodyPrRaw = { "a:bodyPr": [], ":@": { "@_wrap": "square", "@_rtlCol": "0" } }`
   - `txBody.paragraphs = [{ runs: [{ text }] }]`
4. Append to `slide.shapes`.
5. Set `dirty.slides.add(slidePartPath)`.

Errors:

- `unknown-target` — `slideIndex` out of range.
- `invalid-payload` — non-positive size.

Diff: `{ kind: "node-inserted", path: ["slides", slideIndex, "shapes", newIdx], summary: "text-box" }`.

## Diff format per command

Each handler returns a `DocumentDiff` whose `changes` describe what
happened, e.g.:

- `set-text` → `{ kind: "node-updated", path: ["slides", N, "shapes", M, "txBody"], field: "text", summary: "+\"Hello\"" }`
- `add-slide` → `{ kind: "node-inserted", path: ["slides", N], summary: "slide" }`
- `move-slide` → `{ kind: "node-moved", from: ["slides", from], to: ["slides", to], summary: "slide" }`
- `insert-image` → up to two changes: `node-inserted` for the picture, optional `part-added` for a fresh media part.
