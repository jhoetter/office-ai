# DOCX — Agent Commands

> Complete typed list. Twenty-two commands are implemented today
> ("**P0**" + the three comment-lifecycle commands + the two
> header/footer-text commands + the two tracked-change resolution
> commands + the four typed-table commands shipped in P1.3 / W7 + the
> typed inline-image command shipped in P1.3 / W8 + the two
> list/numbering commands and two hyperlink commands shipped in P1.4 /
> W10+W11). No commands remain stubbed at the close of P1.4.

## Common types

```typescript
import type { DocxPosition, DocxSelection } from "./document-model.md";

export interface DocxCommandBase<TType extends string, TPayload> extends Command<TType, TPayload> {
  // type & payload from Command<>; convenience-typed here per command.
}

export interface TextFormatPayload {
  /** Toggle marks; absence = leave unchanged. */
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  fontFamily?: string;
  fontSize?: number; // half-points
  color?: string; // RRGGBB no #
  highlight?: string;
}
```

## Commands (P0 — implemented this session)

### `docx:insert-text`

```typescript
type InsertTextPayload = { at: DocxPosition; text: string };
```

Insert literal text at `at`. If `at.run` is set, the text splits the
target run at `at.offset`; otherwise it appends a new run with default
properties at the start of the paragraph.

OOXML impact: modifies one or more `<w:r><w:t>...</w:t></w:r>` inside the
target paragraph. May split a run.

### `docx:delete-range`

```typescript
type DeleteRangePayload = { range: DocxSelection };
```

Delete content within the range. If the range is single-paragraph the
affected runs are spliced and trimmed. If the range spans multiple
paragraphs (`start.paragraph !== end.paragraph`) the handler:

1. trims the **start** paragraph from the start boundary to its end,
2. drops every fully-contained intermediate paragraph,
3. trims the **end** paragraph from its beginning to the end boundary,
4. **merges** the trimmed start and trimmed end paragraphs into a
   single paragraph, preserving the start paragraph's `id` and `pPr`.

Non-paragraph blocks (tables, opaque blocks) inside the span are not
crossed: callers that need to delete one of those should target it with
a dedicated command.

### `docx:format-range`

```typescript
type FormatRangePayload = { range: DocxSelection; format: TextFormatPayload };
```

Apply the format toggle/value to every run intersecting the range. Runs
are split at range boundaries when needed so the formatting is exact.
Multi-paragraph ranges (`start.paragraph !== end.paragraph`) are
supported: the handler walks the paragraph span and applies the format
to the start paragraph's tail, every intermediate paragraph, and the
end paragraph's head. Non-paragraph blocks inside the span are skipped.

### `docx:insert-paragraph`

```typescript
type InsertParagraphPayload = { at: DocxPosition; style?: string };
```

Insert a new empty paragraph immediately before the paragraph at
`at.paragraph` (or at the end if `at.paragraph === body.length`). Optional
`style` sets the new paragraph's `styleId`.

### `docx:set-paragraph-style`

```typescript
type SetParagraphStylePayload = { at: DocxPosition; style: string };
```

Set `paragraph.properties.styleId` of the paragraph at `at.paragraph`.
The style id must already exist in `word/styles.xml`; we do not create
new styles in this session.

### `docx:add-comment`

```typescript
type AddCommentPayload = {
  range: DocxSelection;
  text: string;
  author: string;
  initials?: string;
};
```

Mint a new `commentId`, insert `<w:commentRangeStart>` / `<w:commentRangeEnd>`
around the range, append `<w:commentReference>` after the end marker, and
append a new `<w:comment>` to `word/comments.xml` with body `[Paragraph
{ children: [Run { children: [text(payload.text)] }] }]`.

If `word/comments.xml` does not exist, create it (see
[`serializer.md`](serializer.md) §ensureCommentsPart).

### `docx:resolve-comment`

```typescript
type ResolveCommentPayload = {
  commentId: string;
  /** Defaults to true. Pass false to reopen a previously resolved comment. */
  resolved?: boolean;
};
```

Toggle `comment.resolved`. Drives the `w15:done` attribute in
`word/commentsExtended.xml`. Idempotent: setting the current state again is a
no-op but still bumps the revision so the bus history records the
attempt. Rejected with `unknown-comment` when the id is missing.

### `docx:reply-comment`

```typescript
type ReplyCommentPayload = {
  parentId: string;
  text: string;
  author: string;
  initials?: string;
};
```

Append a new `w:comment` whose `parentId` points at `parentId`. Replies do
**not** add new range markers in the body: by OOXML convention every reply
in a thread shares the parent's `commentRangeStart`/`commentRangeEnd`
anchors, which is what makes Word render them indented under the same
inline range. The serializer emits a `w15:parentPaIdRef` entry in
`word/commentsExtended.xml`. Rejected with `unknown-comment` for a missing
parent and `empty-reply` for blank text.

### `docx:delete-comment`

```typescript
type DeleteCommentPayload = { commentId: string };
```

Remove the comment and its inline range markers (`commentRangeStart`,
`commentRangeEnd`, `commentReference`). When the target is a thread head,
every reply whose `parentId` chains back to it is removed transitively.
This matches Word's behavior when the user deletes the head of a thread
and keeps the command idempotent (re-running on a now-empty subtree is a
no-op). The serializer also drops `word/commentsExtended.xml` (and its
relationship + content-type entries) when no comment in the document
needs extended metadata anymore.

### `docx:set-header-text`

```typescript
type SetHeaderTextPayload = {
  partId: string; // OOXML part path of the header, e.g. "word/header1.xml"
  paragraphIndex: number; // 0-based index into the header part's body
  text: string; // new plain-text content of the targeted paragraph
};
```

Replace the text content of one paragraph inside a header part. The targeted
part is discovered via the part path exposed by
`snapshot.root.headersAndFooters[i].id`. The paragraph's `pPr` and the first
run's `rPr` (italics, font family, etc.) are preserved — only the text leaves
are rewritten — so heading-style headers retain their look. Idempotent:
re-applying the same text bumps the revision but produces an equivalent
snapshot.

Errors:

- `unknown-target` — the `partId` does not match any header part, OR the
  `paragraphIndex` is out of range, OR the targeted block is not a paragraph.

Dirty flags: only `headersAndFooters` (the matching part path is added to the
set). The body and other parts stay byte-identical on serialize.

### `docx:set-footer-text`

```typescript
type SetFooterTextPayload = {
  partId: string; // OOXML part path of the footer, e.g. "word/footer1.xml"
  paragraphIndex: number;
  text: string;
};
```

Mirror of `docx:set-header-text` for footer parts. Same error contract,
same dirty-flag discipline. The handler shares its implementation with
`set-header-text` (the only difference is the `kind: "header" | "footer"`
discriminator the model uses).

### `docx:accept-change`

```typescript
type AcceptChangePayload = { revisionId: string };
```

Resolve a tracked change (`<w:ins>` or `<w:del>` wrapper) by accepting it:

- `<w:ins>` accept: fold the inserted runs into the parent paragraph; drop
  the wrapper.
- `<w:del>` accept: drop the wrapper AND its children (the deletion lands).

The handler walks both the body and every header/footer body looking for
matching wrappers. The resulting snapshot, when serialized and re-parsed,
contains no `RevisionWrapper` whose `revisionId === payload.revisionId` —
this round-trip property is asserted in `tracked-changes.test.ts`.

Errors:

- `unknown-revision` — no wrapper in the document (body or any header/footer)
  has the requested `revisionId`. Empty / missing `revisionId` also throws
  `unknown-revision`.

Dirty flags: `body` if any body wrapper matched; the corresponding entries
in `headersAndFooters` if any header/footer wrapper matched. Both can be set
in the same dispatch.

### `docx:reject-change`

```typescript
type RejectChangePayload = { revisionId: string };
```

Inverse of `docx:accept-change`:

- `<w:ins>` reject: drop the wrapper AND its children (the insertion never
  lands).
- `<w:del>` reject: drop the wrapper, keep its children (the deletion is
  undone).

Same error contract, same dirty-flag discipline as `accept-change`.

## Commands (P1 — typed tables, shipped in P1.3 / W7)

### `docx:insert-table`

```typescript
type InsertTablePayload = {
  at: DocxPosition;
  rows: number;
  cols: number;
  /** Optional explicit column widths in twips. Length MUST equal `cols`. */
  columnWidths?: number[];
  /** Optional table-level properties applied verbatim. */
  properties?: Partial<TableProperties>;
};
```

Insert a fresh `rows × cols` table at `at.paragraph` (or append when
`at.paragraph === body.length`). Every cell is initialised with one empty
paragraph. The new `Table` is returned WITHOUT a `raw` blob so the
serializer regenerates the `<w:tbl>` from its typed model — this is what
lets the byte-preservation invariant continue to hold for _other_,
untouched tables in the same document.

Errors:

- `invalid-payload` — `rows < 1`, `cols < 1`, or `columnWidths.length !== cols`.
- `invalid-position` — `at.paragraph` outside `[0, body.length]`.

### `docx:set-cell-content`

```typescript
type SetCellContentPayload = {
  tableId: NodeId;
  row: number;
  col: number;
  content: BlockNode[];
};
```

Replace the body of one cell wholesale. `tableId` resolves recursively, so
nested tables can be addressed directly. `content` may include further
nested tables, but their ids must NOT collide with the target table or any
of its ancestors (cycle protection). When `content` is empty the handler
synthesises a single empty paragraph because OOXML requires every `<w:tc>`
to contain at least one paragraph.

Errors:

- `unknown-target` — missing `tableId`, no table with that id, OOB
  `row`/`col`, or `content` contains a `Table` whose `id` matches the
  target / an ancestor table.
- `merged-cell-not-supported` — the targeted cell is a `vMerge="continue"`
  continuation cell. Reflowing merge regions is deferred; callers must
  rewrite the `vMerge` chain themselves first.

### `docx:insert-row`

```typescript
type InsertRowPayload = {
  tableId: NodeId;
  /** 0-based row index. `at === rows.length` appends. */
  at: number;
};
```

Insert a fresh row matching the table's grid. Each new cell carries one
empty paragraph and a `<w:tcW>` width inherited from the corresponding
`<w:gridCol>`. The row's `header` flag is intentionally left unset so the
originally-declared header row keeps its semantics regardless of where the
new row lands.

Errors:

- `unknown-target` — missing or unknown `tableId`.
- `invalid-position` — `at` outside `[0, rows.length]`.
- `merged-cell-not-supported` — the row at `at` (the row about to be
  pushed down) starts with one or more `vMerge="continue"` cells, which
  would orphan their `restart` ancestor.

### `docx:insert-column`

```typescript
type InsertColumnPayload = {
  tableId: NodeId;
  /** 0-based column index. `at === grid.length` appends. */
  at: number;
  /** Column width in twips; defaults to an equal split. */
  width?: number;
};
```

Insert a `<w:gridCol>` at column `at` and a fresh `<w:tc>` (single empty
paragraph) at the same horizontal index in every row. When `width` is
omitted the handler equal-splits the existing declared widths; if the grid
declares no widths at all it falls back to 1000 twips.

Errors:

- `unknown-target` — missing or unknown `tableId`.
- `invalid-position` — `at` outside `[0, grid.length]`.
- `merged-cell-not-supported` — at least one row has a cell whose
  `gridSpan` straddles the requested column index. Boundary insertions
  (`at === 0` or `at === grid.length`) are always accepted because they
  sit at the table's own edges.

### `docx:insert-image`

```typescript
type InsertImagePayload = {
  at: DocxPosition;
  data: Uint8Array | ArrayBuffer;
  mimeType: string; // "image/png", "image/jpeg", "image/gif", "image/bmp", "image/tiff", "image/webp", "image/svg+xml"
  width: number; // pixels @ 96 DPI; converted to EMUs internally
  height: number;
  altText?: string; // populates <wp:docPr descr>
  name?: string; // <wp:docPr name>; defaults to "Picture {docPrId}"
};
```

Insert an inline image at `at`:

1. Validates the payload (positive dimensions, supported MIME, non-empty
   bytes).
2. Computes `SHA-256(data)` and de-duplicates against
   `snapshot.root.media`: if a media part with the same digest exists,
   reuses it (and its existing `image` relationship in
   `word/document.xml.rels`).
3. Otherwise mints a fresh `word/media/image{N}.{ext}` part path, a
   fresh `rId{N}` relationship pointing to it, and adds the new
   `MediaPart` to the typed model. The new media part is also recorded
   in `snapshot.dirty.media` so the serializer writes its bytes.
4. Mints a unique `<wp:docPr id>` (`max(existing) + 1` across the whole
   body — including images inside tables, hyperlinks, and revision
   wrappers).
5. Converts `width` / `height` from pixels to EMUs (`9525 EMU = 1 px @
96 DPI`).
6. Builds a typed `InlineImageDrawing` (no `raw`, so the serializer
   regenerates the subtree from the typed model) and splices it into
   the targeted paragraph as a fresh run. With `at.run` set, the
   targeted run is split at `at.offset` and the image run is spliced
   between the two halves; without it, the image becomes the first
   inline of the paragraph.
7. Sets `dirty.body`, plus `dirty.media`, `dirty.relationships`, and
   `dirty.contentTypes` only when the insertion actually added new
   package parts.

OOXML impact: emits a `<w:r><w:drawing><wp:inline>…<pic:pic>…</wp:inline></w:drawing></w:r>`
inside the targeted paragraph; adds a `word/media/imageN.ext` part, an
`image`-typed relationship in `word/_rels/document.xml.rels`, and a
`<Default Extension>` in `[Content_Types].xml` when one isn't already
registered for that extension.

Diff: emits a `node-inserted` change for the drawing leaf. When a brand
new media part is added, also emits a second `part-added` change whose
`path` is `[mediaPartPath]`. De-duplicated insertions emit only the
`node-inserted` change.

Errors:

- `invalid-payload` — empty bytes, non-positive width / height,
  unsupported MIME type.
- `invalid-position` — `at.paragraph` outside `[0, body.length)`.
- `not-paragraph` — `at.paragraph` resolves to a non-paragraph block
  (table, section break, opaque block, …).

## Commands (P1.4 — lists & hyperlinks)

Two pairs of mutation-aware commands shipped together: list/numbering
mutation (W10) reuses the typed `NumberingDefinitions` carrier added to
`DocxDocument` and rewires `<w:numPr>` on a paragraph; hyperlink
mutation (W11) wraps/unwraps a flat-text range and mints or reaps an
external `hyperlink`-typed relationship in `word/_rels/document.xml.rels`.

### `docx:set-paragraph-list`

```typescript
type SetParagraphListPayload = {
  paragraphId: NodeId;
  /** Concrete numbering instance id (matches `<w:num w:numId>`). */
  numId: number;
  /** 0-based level within the abstract numbering definition. */
  ilvl: number;
};
```

Set or replace the numbering reference on a paragraph. The handler:

1. Validates `numId > 0` and `ilvl >= 0`.
2. Resolves the paragraph (recursively — paragraphs nested inside table
   cells are valid targets).
3. Looks up `numId` in `snapshot.root.numbering.nums`. If the doc has no
   `numbering.xml` at all, or `numId` is unknown, the handler rejects
   `unknown-target`.
4. If the resolved `AbstractNum` declares `levels` and `ilvl` exceeds
   `levels.length - 1`, the handler rejects `invalid-payload`. Empty
   `levels` is tolerated — many docs declare numbering by reference only
   and rely on the renderer to default-format unknown levels.
5. Replaces `paragraph.properties.numbering` with `{ numId, ilvl }` and
   strips any opaque `<w:numPr>` carrier from `paragraph.properties.opaqueProps`
   so the serializer emits exactly one `<w:numPr>` from the typed model.
6. Sets `dirty.body` only — `numbering.xml` itself is untouched, so
   `dirty.numbering` stays false and the part round-trips byte-identically
   from the parts cache.

OOXML impact: rewrites `<w:numPr>` inside the paragraph's `<w:pPr>`
(creating `<w:pPr>` if missing). No other parts change.

Diff: `node-updated` for the paragraph with `field: "numbering"` and a
summary like `set list numId=1 ilvl=0`.

Errors:

- `invalid-payload` — `numId <= 0`, `ilvl < 0`, non-integer values, or
  `ilvl` outside the abstract num's declared levels.
- `unknown-target` — paragraph not found, doc has no `numbering.xml`, or
  `numId` not present in `snapshot.root.numbering.nums`.

### `docx:remove-paragraph-list`

```typescript
type RemoveParagraphListPayload = { paragraphId: NodeId };
```

Clear numbering from a paragraph: drops `paragraph.properties.numbering`
and strips any opaque `<w:numPr>` carrier. Sets `dirty.body`.

The handler is **strict**, not idempotent: a paragraph that has no
numbering today rejects `not-applicable`. This forces callers to confirm
that the target was a list item (commands are typed mutations, not
queries) and prevents silent no-ops from masking buggy callers.

Diff: `node-updated` for the paragraph with `field: "numbering"` and
summary `remove list`.

Errors:

- `invalid-payload` — missing `paragraphId`.
- `unknown-target` — paragraph not found.
- `not-applicable` — paragraph is not currently a list item.

### `docx:insert-hyperlink`

```typescript
type InsertHyperlinkPayload = {
  paragraphId: NodeId;
  /** Flat-text range within the paragraph. `start < end`, both in [0, paragraphLength]. */
  range: { start: number; end: number };
  /** External target. Mutually exclusive with `anchor`. */
  url?: string;
  /** Internal bookmark name. Mutually exclusive with `url`. */
  anchor?: string;
};
```

Wrap a contiguous flat-text range in a typed `Hyperlink`. Either `url`
(external) or `anchor` (internal/bookmark) must be set, not both. The
handler:

1. Validates the payload (XOR on `url` / `anchor`, `new URL(url)` parses
   when `url` is set, `range.start < range.end`, both inside the
   paragraph's flat-text length).
2. Resolves the paragraph; rejects `unknown-target` on miss.
3. Rejects `invalid-position` if the range straddles or overlaps an
   existing hyperlink (nested / overlapping hyperlinks are not legal in
   OOXML and Word doesn't render them either) or if the range crosses a
   non-run inline (comment markers, revision wrappers, opaque inlines).
4. Splits the runs at the range boundaries so the captured span is a
   contiguous sequence of `Run` nodes, each retaining its original
   `properties` (bold / italic / colour all survive the wrap). Word's
   `Hyperlink` character style is intentionally **not** applied — that's
   a styling concern handled by callers / UI layers.
5. When `url` is set: searches existing rels in
   `relationships.get("word/document.xml")` for an entry with
   `type === ".../hyperlink"`, `target === url`, and `targetMode === "External"`.
   Reuses the matching `id` if found; otherwise mints a fresh
   `rId{N}` and appends the new relationship. Sets
   `dirty.relationships` only when a brand-new rel was actually minted.
6. Wraps the captured runs in a `Hyperlink` node with either
   `relationshipId` (external) or `anchor` (internal) and splices it
   back into the paragraph at the original position.
7. Sets `dirty.body`. URL reachability is **not** validated; the
   handler only checks well-formed-ness via `new URL()`.

OOXML impact: emits `<w:hyperlink r:id="rIdN">…</w:hyperlink>` (or
`<w:hyperlink w:anchor="…">…</w:hyperlink>` for anchor links) inside
the targeted paragraph; appends a `hyperlink`-typed relationship to
`word/_rels/document.xml.rels` only when a new external target is
introduced.

Diff: emits two changes — a `node-inserted` for the wrapper and a
`node-updated` for the host paragraph (`field: "children"`). Rel
mutation is signalled by `dirty.relationships`, not by a separate
change kind (we don't have a typed `rel-added` change today).

Errors:

- `invalid-payload` — missing `paragraphId`, both or neither of
  `url`/`anchor`, malformed URL, non-integer or non-monotonic range.
- `unknown-target` — paragraph not found.
- `invalid-position` — range outside paragraph length, range straddles
  an existing hyperlink, or range crosses a non-run inline.

### `docx:remove-hyperlink`

```typescript
type RemoveHyperlinkPayload = { hyperlinkId: NodeId };
```

Unwrap a hyperlink: replaces the `<w:hyperlink>` node with its inline
`children` (the runs spread back into the paragraph). The handler:

1. Locates the hyperlink by id, recursively scanning paragraphs in the
   body and inside table cells.
2. Rejects `unknown-target` on miss.
3. After unwrapping, if the hyperlink had a `relationshipId` AND no
   other body hyperlink references the same id, the rel is removed
   from `relationships.get("word/document.xml")` and
   `dirty.relationships` is set. Otherwise the rel is left intact so
   sibling hyperlinks keep their target. The "still referenced" scan
   covers body paragraphs only — header / footer parts carry their own
   typed rels parts (W4) and aren't included this round.
4. Sets `dirty.body`.

OOXML impact: removes the `<w:hyperlink>` wrapper element while
preserving its child `<w:r>` runs. Removes the rel entry from
`word/_rels/document.xml.rels` only when no other body hyperlink uses
it; the rels graph is otherwise untouched.

Diff: a single `node-updated` for the host paragraph (`field: "children"`)
with a summary like `−hyperlink (rel=rId5 removed)` or `−hyperlink (rel=rId5 kept)`.

Errors:

- `invalid-payload` — missing `hyperlinkId`.
- `unknown-target` — no hyperlink with that id in the body.

## Diff format per command

Each handler returns a `DocumentDiff` whose `changes` describe what
happened, e.g.:

- `insert-text` → `{ kind: "node-updated", path: ["body", N, "children", M], field: "text", summary: "+\"Hello\"" }`
- `insert-paragraph` → `{ kind: "node-inserted", path: ["body", N], summary: "paragraph" }`
- `add-comment` → two changes: one for the inserted range markers, one for the new comment in the comments part.
