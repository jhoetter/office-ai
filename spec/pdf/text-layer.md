# PDF — Text Layer

> Selectable, copyable text overlaid on the canvas. Drives selection,
> search highlighting, accessibility reading order, and "copy as
> Markdown".

Cross-references: rendering pipeline in
[`rendering-pipeline.md`](./rendering-pipeline.md);
search in [`search.md`](./search.md);
accessibility integration in [`accessibility.md`](./accessibility.md).

## Source data

The engine surfaces text via `enginePage.getTextContent()`:

```typescript
export interface PdfEngineTextItem {
  readonly str: string;
  readonly transform: readonly [number, number, number, number, number, number]; // PDF user-space
  readonly width: number;
  readonly height: number;
  readonly fontName?: string;
  readonly hasEol?: boolean;
}

export interface PdfEngineTextContent {
  readonly items: ReadonlyArray<PdfEngineTextItem>;
  readonly plain: string;  // concatenated reading-order projection
}
```

`PdfEngineTextItem.transform` is the standard 6-element PDF text
matrix `[a, b, c, d, e, f]`. `(e, f)` is the text origin. Our DOM
positioner converts this to CSS using the page's user-units → CSS
pixels scale.

For PDF.js this delegates to `getTextContent()`. For PDFium it builds
the same shape from `FPDFText_*`. The viewer doesn't care which
backend produced the items.

## DOM stack

For each rendered page the viewer composes:

```
<div class="pdf-page" style="width:Wpx; height:Hpx; position:relative">
  <canvas width="Wdpr" height="Hdpr" style="width:Wpx; height:Hpx;" />
  <div class="pdf-text-layer" aria-hidden="false" style="position:absolute; inset:0;">
    <span style="left:Lpx; top:Tpx; font-size:Spx; transform:scale(Sx,1);">…</span>
    <span …>…</span>
  </div>
  <div class="pdf-annotation-layer" style="position:absolute; inset:0; pointer-events:auto;">…</div>
  <div class="pdf-struct-tree-layer" aria-hidden="true" style="position:absolute; inset:0; pointer-events:none;">…</div>
</div>
```

Each text-layer `<span>`:

- is positioned with `left` / `top` derived from the item transform;
- uses `font-family: <fontName fallback>` mapped to a system font
  family by the substitution table;
- uses `font-size: <Spx>`;
- uses `transform: scale(Sx, 1)` to match the rendered glyph width
  (PDF.js and PDFium both produce this scale factor);
- has `data-text-index="<i>"` for search highlighting and selection
  serialization.

The text layer is **above** the canvas (so selection works) but its
text is rendered in the same color as the canvas + `color:transparent`
so the user sees only the canvas glyphs. On selection, the browser's
native selection styles paint over the transparent text — the text
appears highlighted exactly where the canvas glyphs are.

## Selection

- **Single-click drag** selects characters.
- **Double-click** selects the word.
- **Triple-click** selects the paragraph (defined as items between
  `hasEol` boundaries).
- **Shift-click** extends selection.
- **Cmd/Ctrl+A** selects all text on the current page (per-page A11Y
  expectation; full-doc select is a separate shortcut).
- **Rectangular (lasso) selection** with `Alt`+drag — selects all
  items whose center point falls inside the rect; used for image and
  table extraction.

## Copy fidelity

| Operation                              | Output                                   |
| -------------------------------------- | ---------------------------------------- |
| Cmd/Ctrl+C                             | Plain text in reading order              |
| Cmd/Ctrl+Shift+C                       | Markdown (preserves headings + lists where the struct-tree is available) |
| Right-click → Copy as plain            | Plain text                               |
| Right-click → Copy as Markdown         | Markdown                                 |
| Right-click → Copy as HTML             | HTML with `<span>` per item              |
| Right-click → Copy text + URL          | `${text} [${current page URL}#page=N]`   |

The Markdown serializer uses the struct tree when present
(`/H1` → `#`, `/H2` → `##`, `/L`/`/LI` → `- `, `/P` → paragraph).
For untagged PDFs it falls back to a heuristic: large font + bold →
heading; consistent left-margin bullet glyph → list.

## Search-hit highlighting

When [`search.md`](./search.md) returns a match `(pageNumber, start,
end)` we walk the page's text items and split the affected `<span>`s
at the match boundaries. The match characters get
`class="pdf-search-hit"` (Tailwind `bg-yellow-200/60` light, `bg-yellow-700/40`
dark) and the *current* hit gets `class="pdf-search-hit current"`
(brighter background + auto-scroll-into-view).

Highlighting is non-destructive: cancelling search restores the
original `<span>` structure.

## OCR fallback path

For pages with `hasTextLayer === false` (typical scans), the
text-layer is empty by default. The viewer shows a banner: **"This
page has no text layer. Run OCR (German + English) to make it
selectable and searchable."** Clicking triggers
`packages/pdf-ocr.addTextLayer(buffer, [pageIndex], "deu+eng")`.

The OCR pass:

1. Rasterizes the page at 300 DPI via the engine.
2. Pipes the bitmap to `tesseract.js` (lazy-loaded; ~1 MB worker).
3. Receives word-level bounding boxes + recognized text.
4. Produces a content-stream patch: invisible text (`Tr 3` text
   rendering mode) at the bbox positions with the recognized text.
5. Writes the patch back via incremental save.

After the patch lands, re-parsing the page produces a non-empty text
layer — selection, search, and the agent's `read --format text` all
work transparently.

## RTL and bidi

The engine returns items in **visual order** (left-to-right on the
page). For RTL languages (Arabic, Hebrew) the visual order is the
reverse of the logical order. We:

1. Detect RTL items by Unicode bidi property of the first letter
   (`getBidiCategory` from `unicode-properties`).
2. For mixed-direction lines, reorder items into logical order using
   the Unicode Bidirectional Algorithm (UAX #9). The reordering is
   applied to the **plain text** projection (used for search and
   copy); the `<span>` positions stay in visual order so on-screen
   selection still tracks the user's drag.
3. Set `dir="rtl"` on the affected `<span>` so the cursor caret
   moves the user expects.

This means an Arabic word in a German document is selectable and
copyable in the correct logical order while remaining visually
positioned over the canvas glyphs.

## CJK considerations

CJK pages frequently use vertical writing mode and ligature-heavy
fonts. We:

- Detect vertical pages by checking `transform[1]` (the b-component
  of the text matrix) — vertical lines have non-zero `b`.
- Render vertical text spans with `writing-mode: vertical-rl` so
  selection drags work in the natural direction.
- For ligature-heavy fonts (Han ideographs frequently fall into a
  single CID), the engine returns the visible glyph; if the
  underlying CMap maps it to multiple Unicode codepoints, we use the
  CMap-mapped string for the text-layer so search and copy return the
  correct text.
- For pages with `hasCustomCMap` we trigger PDFium fallback (see
  [`engine-strategy.md`](./engine-strategy.md)) which has more
  consistent CID → Unicode mapping than PDF.js.

## Performance

- Text-layer build is bounded to **viewport pages ± 1**. Pages farther
  away have empty text layers; building is deferred until they enter
  the window.
- Item-to-`<span>` conversion is O(items per page) which is bounded
  by typical page density (≤ 4000 items). On a Mac M1 this completes
  in well under 16 ms.
- The text layer is destroyed when the page exits the text-layer
  window, freeing DOM. The plain-text projection is retained on the
  snapshot for search and the agent.
