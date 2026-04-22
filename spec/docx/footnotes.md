# DOCX — Footnotes (Fußnoten)

> Status: F1. Promotes `word/footnotes.xml` and `word/endnotes.xml`
> from container-only opaque preservation to a first-class typed
> model with a bottom-of-page rendering lane and insert/edit
> commands. Supersedes the "footnotes — out of scope" line in
> [`feature-scope.md`](feature-scope.md) (flipped to **IN-P0**).

## Why

Today the OOXML mapping treats `word/footnotes.xml` /
`word/endnotes.xml` as opaque bytes that round-trip but cannot
be read or edited. Result: every fixture with footnotes
(thesis-style documents, contracts, legal briefs) renders the
in-body footnote reference as plain `RunFootnoteReference`
opaque content with no visible target on the page. Users cannot
add a footnote, edit one, or even see which references resolve
to which note text.

Word and LibreOffice both render footnotes as a numbered list
anchored to the bottom of the page that contains the matching
reference. We mirror that.

## OOXML mapping

| Part                                                                | Today            | After F1                                                                          |
| ------------------------------------------------------------------- | ---------------- | --------------------------------------------------------------------------------- |
| `word/footnotes.xml`                                                | opaque verbatim  | typed `FootnotesPart { footnotes: Footnote[] }`, `raw` for unmodelled tail        |
| `word/endnotes.xml`                                                 | opaque verbatim  | typed `EndnotesPart { endnotes: Endnote[] }` (mirror of footnotes — same shape)   |
| `<w:footnoteReference w:id="N"/>` inside body / header runs         | `OpaqueRunChild` | `FootnoteReferenceLeaf { footnoteId: number, customMarkFollows?: boolean }`       |
| `<w:endnoteReference w:id="N"/>`                                    | `OpaqueRunChild` | `EndnoteReferenceLeaf { endnoteId: number }`                                      |
| `<w:settings w:footnotePr>` / `<w:endnotePr>` (numbering, position) | preserved opaque | passed through unchanged for F1; surfaced in F2 if user needs to change numbering |

Relationships: each footnote part is referenced from
`word/_rels/document.xml.rels` with relationship type
`...officeDocument/2006/relationships/footnotes` (or
`endnotes`). Content-Type override in `[Content_Types].xml` uses
`application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml`.
We add to `MODELED_PREFIXES` in
`packages/docx/src/parser/parse.ts`.

### `<w:footnote>` shape

```xml
<w:footnote w:type="separator" w:id="-1">
  <w:p><w:r><w:separator/></w:r></w:p>
</w:footnote>
<w:footnote w:type="continuationSeparator" w:id="0">
  <w:p><w:r><w:continuationSeparator/></w:r></w:p>
</w:footnote>
<w:footnote w:id="1">
  <w:p>
    <w:pPr><w:pStyle w:val="FootnoteText"/></w:pPr>
    <w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr>
      <w:footnoteRef/></w:r>
    <w:r><w:t xml:space="preserve"> Body of footnote one.</w:t></w:r>
  </w:p>
</w:footnote>
```

The `type` attribute is one of `normal | separator |
continuationSeparator | continuationNotice`. The first three are
the standard separator notes Word inserts on every document.

## Typed model

```ts
export type FootnoteType = "normal" | "separator" | "continuationSeparator" | "continuationNotice";

export interface Footnote {
  /** OOXML id; -1 and 0 are conventionally separator/continuation. */
  readonly id: number;
  readonly type: FootnoteType;
  /** Body blocks; identical schema to `DocxDocument.body` (paragraphs, tables, ...). */
  readonly body: ReadonlyArray<Block>;
  /** Verbatim attributes we don't need to change (`w:customMarkFollows`, etc.). */
  readonly raw?: Readonly<Record<string, string>>;
}

export interface FootnotesPart {
  readonly footnotes: ReadonlyArray<Footnote>;
  /** Other top-level children of `<w:footnotes>` we don't model. */
  readonly tail?: ReadonlyArray<unknown>;
}

export type Endnote = Footnote;
export type EndnotesPart = FootnotesPart;

export interface FootnoteReferenceLeaf {
  readonly kind: "footnote-ref";
  readonly id: NodeId;
  readonly footnoteId: number;
}

export interface EndnoteReferenceLeaf {
  readonly kind: "endnote-ref";
  readonly id: NodeId;
  readonly endnoteId: number;
}
```

`Footnote.body` reuses the `Block` type from the body model so
the same parser, serializer, and ProseMirror schema work
unchanged. Inline run formatting (bold, italic, font, color)
inside footnote text is therefore inherently supported.

`DocxDocument` gains `footnotesPart?: FootnotesPart` and
`endnotesPart?: EndnotesPart` (optional — many documents have
neither).

## Commands

### `docx:insert-footnote`

```ts
export interface InsertFootnotePayload {
  /** Stable id of the body paragraph (or any paragraph, including header/footer parts). */
  paragraphId: NodeId;
  /** Byte offset in the paragraph's flat text. */
  offset: number;
  /** Body of the new footnote — defaults to a single empty paragraph. */
  body?: ReadonlyArray<Block>;
}
```

Behaviour:

- Allocates the next unused id (`max(existing) + 1`, starting at
  1; -1 / 0 reserved for separators).
- Mints a new `Footnote` with the supplied body or one empty
  paragraph styled `FootnoteText`.
- Splits the targeted run at `offset`, splices in a fresh
  zero-length run carrying a `FootnoteReferenceLeaf` styled with
  `FootnoteReference`.
- Dirties: `body` (or the owning H/F part) **and** the
  `footnotesPart`.

### `docx:set-footnote-body`

```ts
export interface SetFootnoteBodyPayload {
  footnoteId: number;
  body: ReadonlyArray<Block>;
}
```

Replaces a single footnote's body. Round-trips through the same
block parser/serializer the body uses.

### `docx:delete-footnote`

```ts
export interface DeleteFootnotePayload {
  footnoteId: number;
}
```

- Removes the `Footnote` from the part.
- Walks every `Run` (body + headers + footers) and strips any
  `FootnoteReferenceLeaf` whose `footnoteId` matches.
- Renumbering of subsequent ids is **not** automatic; we leave
  them to keep stable references during heavy editing. A
  follow-up `docx:renumber-footnotes` command is the right place
  for that and is documented as P1.

## Renderer

`packages/docx/src/renderer/footnote-layout.ts` (new) computes
which footnote ids land on which page chunk by walking the body
runs in chunk order. The current page chunker
(`page-chunker.ts`) already returns ordered `PageChunk`s — we
extend it to expose the run sequence per chunk so the footnote
layout pass can collect refs in the order they appear.

`pageDecorationsPlugin` in
`apps/web/app/lib/page-decorations.ts` adds a footnote lane
**above** the footer zone:

```
<div class="pm-page-cap-bottom">
  <div class="pm-page-footnote-lane" data-page="N">
    <hr class="pm-footnote-separator" />
    <ol class="pm-footnote-list">
      <li data-footnote-id="1">…body of footnote 1…</li>
      <li data-footnote-id="2">…</li>
    </ol>
  </div>
  <div class="pm-page-zone pm-page-zone-footer" …>…</div>
</div>
```

Each `<li>` mounts a non-editable preview by default; clicking
it transitions the `ActivePart` state machine into
`{ kind: "footnote", footnoteId }` and hot-swaps the preview for
a per-footnote ProseMirror instance bound to a partial document
made of that footnote's body. Esc / click-out commits and
returns to body.

A click on a `FootnoteReferenceLeaf` in the body scrolls the
matching `<li>` into view, briefly highlights it, and (if Alt is
held) immediately focuses it for editing — same pattern Word uses.

## Toolbar

The body toolbar (Insert tab in the new ribbon — see
`spec/shared/ribbon-design.md`) gains an **Insert Footnote**
button that dispatches `docx:insert-footnote` at the current
caret. The Header & Footer contextual tab gains the same button
so footnotes added from a header reference work too (Word allows
this; the footnote still renders at the bottom of the body
page).

## Round-trip invariants

1. **No-edit byte-equivalence.** A document with footnotes
   parsed and serialized without any footnote edit re-emits
   `word/footnotes.xml` byte-identical to the input.
2. **Single-footnote edit.** Editing one footnote's body re-emits
   that footnote's `<w:footnote>` only; siblings stay
   byte-identical.
3. **Reference parity.** Every `FootnoteReferenceLeaf` in the
   model corresponds to exactly one `<w:footnoteReference>` on
   serialize, with matching `w:id`.
4. **Separator preservation.** The `separator` and
   `continuationSeparator` notes (`w:id="-1"` and `"0"`)
   round-trip even when the document has zero `normal`
   footnotes.
5. **Endnote symmetry.** All of the above hold for endnotes
   identically; the renderer attaches endnote bodies to the end
   of the document, not per-page.

## Acceptance criteria

A1. **Round-trip — fixture with footnotes.**
A new fixture `fixtures/docx/09-footnotes.docx` parses, exports,
re-parses with three footnotes preserved (ids 1, 2, 3) and
`word/footnotes.xml` byte-identical.

A2. **Insert.** `docx:insert-footnote({ paragraphId, offset })`
inside an empty body paragraph creates a new footnote whose body
is one empty `FootnoteText`-styled paragraph and a reference
leaf at the caret.

A3. **Edit.** `docx:set-footnote-body({ footnoteId: 1, body: [P("hello")] })`
replaces footnote 1's body and dirties only the footnotes part
(not the body).

A4. **Delete.** `docx:delete-footnote({ footnoteId: 1 })`
removes the footnote and every reference leaf to it across body

- H/F parts.

A5. **Renderer.** Loading the fixture in the editor shows a
numbered list at the bottom of the appropriate page; clicking a
list item activates a per-footnote PM; Esc returns focus to
body.

A6. **No regressions.** All existing 208+ docx tests stay green.
`make audit-roundtrip` includes the new fixture and stays at
100%.

## Out of scope (F1)

- Footnote separator customisation (different separator per
  section). Preserve only.
- Footnote numbering format / restart per section
  (`<w:footnotePr w:numFmt>`). Preserve only.
- Continuation notice rendering ("Continued on next page…").
- Endnote position other than `docEnd` (rare).
- `docx:renumber-footnotes` (P1; ids stay stable across edits).
- Bottom-of-page footnote rendering inside header/footer parts
  (Word allows refs anywhere; we render the lane only on body
  pages because H/F parts repeat across pages and rendering the
  same footnote per page is wrong).
