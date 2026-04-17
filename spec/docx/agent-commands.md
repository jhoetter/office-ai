# DOCX — Agent Commands

> Complete typed list. Six commands land this session ("**P0**"). The other
> six are stubbed and throw `NotImplementedError` until a follow-up session.

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

## Commands (P1 — stubbed; throw `NotImplementedError`)

```typescript
"docx:insert-table"        { at: DocxPosition; rows: number; cols: number }
"docx:set-cell-content"    { tableId: NodeId; row: number; col: number; content: BlockNode[] }
"docx:insert-image"        { at: DocxPosition; data: ArrayBuffer; mimeType: string; width: number; height: number }
"docx:resolve-comment"     { commentId: string }
"docx:accept-change"       { revisionId: string }
"docx:reject-change"       { revisionId: string }
```

Each P1 stub:

1. Validates payload shape (with Zod-style guards in code).
2. Throws `new NotImplementedError("docx:insert-table", { reason: "Tables are P1; see docs/build-log/docx.md" })`.
3. The error is caught by the bus and surfaced as a `Mutation` with
   `status: "rejected"` so callers (CLI, UI) can react gracefully.

## Diff format per command

Each handler returns a `DocumentDiff` whose `changes` describe what
happened, e.g.:

- `insert-text` → `{ kind: "node-updated", path: ["body", N, "children", M], field: "text", summary: "+\"Hello\"" }`
- `insert-paragraph` → `{ kind: "node-inserted", path: ["body", N], summary: "paragraph" }`
- `add-comment` → two changes: one for the inserted range markers, one for the new comment in the comments part.
