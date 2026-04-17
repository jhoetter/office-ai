# DOCX — Agent Commands

> Complete typed list. Seventeen commands are implemented today
> ("**P0**" + the three comment-lifecycle commands + the two
> header/footer-text commands + the two tracked-change resolution
> commands + the four typed-table commands shipped in P1.3 / W7).
> Only `docx:insert-image` remains stubbed; it throws
> `NotImplementedError` until a follow-up session.

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

## Commands (P1 — still stubbed; throw `NotImplementedError`)

```typescript
"docx:insert-image"        { at: DocxPosition; data: ArrayBuffer; mimeType: string; width: number; height: number }
```

The remaining stub:

1. Validates payload shape (with Zod-style guards in code).
2. Throws `new NotImplementedError("docx:insert-image", { reason: "..." })`.
3. The error is caught by the bus and surfaced as a `Mutation` with
   `status: "rejected"` so callers (CLI, UI) can react gracefully.

## Diff format per command

Each handler returns a `DocumentDiff` whose `changes` describe what
happened, e.g.:

- `insert-text` → `{ kind: "node-updated", path: ["body", N, "children", M], field: "text", summary: "+\"Hello\"" }`
- `insert-paragraph` → `{ kind: "node-inserted", path: ["body", N], summary: "paragraph" }`
- `add-comment` → two changes: one for the inserted range markers, one for the new comment in the comments part.
