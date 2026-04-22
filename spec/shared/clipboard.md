# Shared — Cross-format clipboard

> Status: F1. Specifies the `application/x-officeai-embed+json`
> envelope plus per-format payloads enabling Cmd+C / Cmd+V across
> tabs and across formats. Builds on Phase 5 of the prior
> [`NIGHT_REPORT.md`](../../NIGHT_REPORT.md) (XLSX → DOCX, XLSX →
> PPTX), promotes the embed feature flag from default-off to
> default-on, and adds PPTX shape/slide and DOCX block payloads.

## Why

Today's clipboard story is uneven:

- **XLSX**: rich intra-format clipboard (formulas, styles,
  merges, fingerprint). System clipboard TSV/HTML works well.
- **DOCX**: ProseMirror handles text/HTML well; no structured
  embed for cross-DOCX paste with style fidelity.
- **PPTX**: only `pptx:duplicate-shape` / `pptx:duplicate-slide`
  — no system clipboard at all. Cannot copy a shape from one
  deck to another.

Worse, the `application/x-officeai-embed+json` envelope is
gated behind `NEXT_PUBLIC_OAI_EMBED` so cross-format paste
doesn't ship by default.

This spec lands a single shared clipboard surface used by all
three editors, and removes the flag.

## Envelope

`apps/web/app/lib/embed/envelope.ts` defines:

```ts
export const EMBED_MIME = "application/x-officeai-embed+json";

export interface EmbedEnvelope<P extends EmbedPayload = EmbedPayload> {
  /** Schema version. Bump when payload shape changes. */
  readonly version: 2;
  readonly source: { kind: "office-ai"; appVersion: string };
  readonly created: string; // ISO8601
  readonly payload: P;
}

export type EmbedPayload =
  | XlsxRangePayload
  | XlsxChartImagePayload
  | DocxBlocksPayload // NEW
  | DocxTablePayload // NEW
  | PptxShapesPayload // NEW
  | PptxSlidesPayload; // NEW
```

Version 2 is a **superset** of version 1; old `xlsx-range`
producers stay compatible. Consumers MAY accept version 1 with
the same shape.

### `pptx-shapes`

```ts
export interface PptxShapesPayload {
  readonly kind: "pptx-shapes";
  readonly slideEMU: { cx: number; cy: number };
  /** Source slide id, used by intra-deck paste for relative positioning. */
  readonly sourceSlideId?: string;
  /** Selected shape subtree(s) in our typed model JSON. */
  readonly shapes: ReadonlyArray<unknown /* Shape | TextShape | GroupShape */>;
  /** Linked media bytes (images, video posters), keyed by relPath used inside `shapes`. */
  readonly media?: Readonly<Record<string, { contentType: string; base64: string }>>;
  /** Pre-rendered PNG of the selection for cross-format consumers (DOCX/XLSX). */
  readonly png?: { dataUrl: string; widthEMU: number; heightEMU: number };
}
```

### `pptx-slides`

```ts
export interface PptxSlidesPayload {
  readonly kind: "pptx-slides";
  readonly slides: ReadonlyArray<unknown /* Slide */>;
  readonly media?: Readonly<Record<string, { contentType: string; base64: string }>>;
  /** Source layout part paths so paste can attempt layout reuse. */
  readonly layoutHints: ReadonlyArray<string>;
}
```

### `docx-blocks`

```ts
export interface DocxBlocksPayload {
  readonly kind: "docx-blocks";
  readonly blocks: ReadonlyArray<unknown /* Block */>;
  /** Style table snippets needed by the blocks: { styleId: serialized w:style }. */
  readonly styles: Readonly<Record<string, string>>;
  /** Numbering definitions referenced by list-style blocks. */
  readonly numbering?: Readonly<Record<string, string>>;
  /** Linked media bytes, keyed by relPath used inside `blocks`. */
  readonly media?: Readonly<Record<string, { contentType: string; base64: string }>>;
}
```

### `docx-table`

```ts
export interface DocxTablePayload {
  readonly kind: "docx-table";
  /** Table block from our model. */
  readonly table: unknown;
  /** Same accompanying styles/numbering/media as docx-blocks. */
  readonly styles: Readonly<Record<string, string>>;
}
```

This is the variant called out as deferred in the prior night
report — it enables DOCX-table → XLSX-range paste.

## Producers

| Editor | Action                                                          | Writes                                                                                                |
| ------ | --------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------- |
| XLSX   | Copy/cut a range                                                | text/plain (TSV), text/html, EMBED_MIME `xlsx-range`                                                  |
| XLSX   | Copy a chart selection                                          | text/plain (chart name), text/html (`<img>` with PNG), EMBED_MIME `xlsx-chart-image`                  |
| DOCX   | Copy block selection (one or more paragraphs / tables / images) | text/plain, text/html, EMBED_MIME `docx-blocks` (or `docx-table` when selection is exactly one table) |
| PPTX   | Copy shapes selection on canvas                                 | text/plain (concatenated text content), text/html (PNG), EMBED_MIME `pptx-shapes`                     |
| PPTX   | Copy slides from slide-rail                                     | text/plain (slide titles), text/html (thumbnail PNG), EMBED_MIME `pptx-slides`                        |

## Consumers

The paste matrix:

| → Target editor | xlsx-range                                                             | docx-blocks                                       | docx-table                                        | pptx-shapes                                    | pptx-slides                                |
| --------------- | ---------------------------------------------------------------------- | ------------------------------------------------- | ------------------------------------------------- | ---------------------------------------------- | ------------------------------------------ |
| XLSX            | native paste-range                                                     | (parse table cells from blocks containing tables) | native: insert as range starting at caret         | (insert PNG image)                             | (insert PNG image)                         |
| DOCX            | insert-table from rangeFingerprint                                     | native paste-blocks                               | insert-table                                      | (insert PNG image)                             | (insert PNG image)                         |
| PPTX            | insert text box with TSV (existing) or `pptx:insert-table` (preferred) | (text-only fallback)                              | (insert PNG image, OR `pptx:insert-table` if Alt) | native paste-shapes (`pptx:insert-shape-tree`) | native paste-slides (`pptx:insert-slides`) |

When two payloads are present (envelope + html/plain), the rule
is:

1. If EMBED_MIME contains a payload of a kind the target supports
   natively, use that.
2. Otherwise fall back to the most structured non-embed
   representation (HTML > plain).
3. Hold Alt/Option while pasting to force the secondary
   representation (e.g. paste an XLSX range as image instead of
   table).

This matches Office's "Paste Special" behaviour and gives users
an explicit escape hatch.

## Commands

| Command                  | Payload                                                                                         | Effect                                                                                                                      |
| ------------------------ | ----------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------- |
| `pptx:insert-shape-tree` | `{ slideId; shapes; media; offset?: { dx, dy } }`                                               | Hydrates shapes into a slide; mints fresh ids; copies media into `ppt/media/` with SHA-256 dedup                            |
| `pptx:insert-slides`     | `{ slides; media; insertAfterSlideId?; layoutMatchStrategy: "match-by-name" \| "use-default" }` | Inserts slides; hydrates media and layout refs                                                                              |
| `docx:insert-blocks`     | `{ paragraphIndex; blocks; styles; numbering?; media? }`                                        | Inserts blocks at a position; merges referenced styles into the doc style table; copies media into `word/media/` with dedup |
| `docx:insert-table`      | (existing — augmented) `{ paragraphIndex; table; styles? }`                                     | Inserts a table block, optionally importing referenced cell styles                                                          |

All commands are undoable (single user action) and broadcast
through the realtime command bus so peers see the paste live.

## Feature flag

`NEXT_PUBLIC_OAI_EMBED` defaults to **on** after this session.
The flag becomes a kill-switch:

- `1` / unset / `true` → embed enabled.
- `0` / `false` → fall back to text/HTML only (Phase-1
  behaviour) for emergency rollback.

`isEmbedEnabled()` updates accordingly. Spec deviation logged in
[`docs/build-log/clipboard.md`](../../docs/build-log/clipboard.md).

## Security

- Clipboard payloads embed media as base64. Treat all incoming
  embeds as untrusted: validate `contentType` against an
  allowlist (PNG/JPEG/GIF/WebP for images, MP4/WebM/MP3/M4A for
  media) and **reject** anything outside before splicing into
  the OOXML container.
- Maximum payload size: **50 MB**. Larger pastes warn the user
  and abort.
- Cross-origin paste: the consumer must come from the same
  origin (we don't sign the envelope; this is a UX guardrail
  not an auth boundary).

## Round-trip invariants

1. **PPTX-shape → PPTX-shape.** Copy a shape with image fill,
   paste in another deck, save, reopen — the image is present
   and identical (SHA-256 dedup'd in target's `ppt/media/`).
2. **DOCX-blocks → DOCX.** Copy a paragraph styled with
   `MyHeading1`, paste into a doc lacking that style — the style
   table grows by one entry; the pasted paragraph references it.
3. **XLSX-range → DOCX-table.** Already shipped (Phase 5 prior
   night). Regression-tested.
4. **DOCX-table → XLSX-range.** Copy a 4×3 table from DOCX,
   paste into XLSX at A1 — sheet now has 4×3 cells with values
   matching cell text. Cell styles map (bold → bold).

## Acceptance criteria

A1. **PPTX shape clipboard.** Copy a rounded-rect on slide 1,
new tab opens deck 2, paste — the shape appears with its fill,
text, and geometry adjustments.

A2. **PPTX slide clipboard.** Copy slide 3 from deck 1, paste in
deck 2 — slide 3 appears at the end of deck 2 with its layout.

A3. **DOCX blocks.** Copy a paragraph + image, paste into
another doc — paragraph and image both render; style and media
ids are unique in the target.

A4. **Cross-format DOCX→XLSX table.** Copy a 3×3 DOCX table,
paste in XLSX A1 — A1:C3 contains the cell text.

A5. **Embed default-on.** `NEXT_PUBLIC_OAI_EMBED` unset —
clipboard payload includes EMBED_MIME entries.

A6. **Kill switch works.** Setting
`NEXT_PUBLIC_OAI_EMBED=0` returns to text/HTML-only behaviour.

## Out of scope (F1)

- Encrypting envelopes between sessions.
- Cross-app paste (i.e. paste from PowerPoint desktop). Browsers
  don't surface custom MIME types from desktop apps reliably.
- Drag-and-drop between editors (clipboard only). DnD lands when
  the multi-window layout ships.
- Paste-as-link (live reference) for cross-format.
