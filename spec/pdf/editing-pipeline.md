# PDF — Editing Pipeline

> Page-level operations flowing through the `CommandBus` and
> serialized via `pdf-lib` incremental save. Annotations and form
> fills also flow through commands; this doc focuses on the page
> operations that mutate the page tree.

Cross-references: command surface in
[`agent-commands.md`](./agent-commands.md);
agent invocation in [`cli.md`](./cli.md);
serializer constraint in [`document-model.md`](./document-model.md);
edge cases (encryption, signatures, broken xref) in
[`edge-cases.md`](./edge-cases.md).

## Architecture

Page operations live in `packages/pdf-edit`. Each function is **pure**:
takes a `Uint8Array` (the original PDF) and returns a `Uint8Array`
(the mutated PDF). No DOM, no React, no global state. Backing library
is `pdf-lib` (MIT, pure-TS).

```
              CommandBus.dispatch(cmd)
                       │
                       ▼
            ┌──────────────────────┐
            │ handler in packages/ │
            │ pdf/src/commands/    │
            └─────────┬────────────┘
                      │ produces new PdfSnapshot + queues a serializer task
                      ▼
            ┌──────────────────────┐
            │ packages/pdf-edit op │  (pdf-lib)
            │ (rotate / reorder /  │
            │  delete / insert / …)│
            └─────────┬────────────┘
                      ▼
              new Uint8Array
```

The handler updates the typed model **synchronously** (e.g. a rotate
mutates `PdfPage.rotation`). The actual byte-level pdf-lib transform
runs **lazily** on `agent.exportFile()` — we batch all pending
operations into a single incremental-save pass to minimize xref
churn.

## Operations

### `pdf:rotate-pages`

Payload:

```typescript
{
  pages: ReadonlyArray<number>;          // 1-indexed
  delta: 90 | 180 | 270 | -90 | -180 | -270;
}
```

- Typed model: each `PdfPage.rotation = (existing + delta) mod 360`.
- pdf-lib op: `page.setRotation(degrees(newRotation))` for each page.
- Incremental save: each affected `/Page` dict is rewritten with a
  new `/Rotate` value. Content stream untouched.

### `pdf:set-page-rotation`

Payload:

```typescript
{ pageNumber: number; rotation: 0 | 90 | 180 | 270 }
```

Absolute set; otherwise identical to rotate-pages.

### `pdf:reorder-pages`

Payload:

```typescript
{ order: ReadonlyArray<number> }  // permutation of 1..N
```

- Typed model: `pages` rebuilt in the new order; `pageNumber` field
  is reassigned 1..N.
- pdf-lib op: walk the page tree, rebuild it in the requested order.
- Incremental save: `/Pages/Kids` is rewritten with the new order;
  individual `/Page` dicts unchanged.

### `pdf:delete-pages`

Payload:

```typescript
{ pages: ReadonlyArray<number> }
```

- Typed model: pages dropped from the array; `pageNumber` reassigned.
- pdf-lib op: `pdf.removePage(i)` for each (in descending order to
  keep indexes stable during the loop).
- Incremental save: dropped pages remain in the byte stream as
  unreferenced objects; `/Pages/Kids` and `/Pages/Count` are updated.
  A subsequent `office-agent pdf optimize` (full re-serialize) would
  garbage-collect the orphans.

### `pdf:insert-pages` (CLI only this session)

Payload:

```typescript
{ at: number; sourceBuffer: Uint8Array; sourcePages?: ReadonlyArray<number> }
```

- pdf-lib op: open the source, copy the requested pages
  (`pdf.copyPages(source, indices)`), insert at `at`. Embedded
  resources from the source are deduplicated by content hash.
- Note: not exposed as a `pdf:*` typed command in this night session
  (no UI affordance yet) but available via
  `office-agent pdf insert-pages`.

### `pdf:extract-pages` (CLI)

`office-agent pdf extract-pages --pages 1-10 --out chapter1.pdf`

- pdf-lib op: create a new PDF, `copyPages` the range, save fresh.
- This is a full serialize, not an incremental update — the output
  is a brand-new file.

### `pdf:merge` (CLI)

`office-agent pdf merge --files a.pdf b.pdf c.pdf --out merged.pdf`

- pdf-lib op: create a new PDF, `copyPages` from each source in
  order, save fresh.

### `pdf:split` (CLI)

`office-agent pdf split --file f.pdf --by range|bookmark|size --out ./parts/`

- `range`: split into N files of M pages each (or by an explicit
  range list `--ranges 1-10,11-20,21-end`).
- `bookmark`: split at every top-level `/Outlines` entry, naming
  files after the bookmark title (slug).
- `size`: greedily fill output files until each approaches a target
  size in bytes (`--max-bytes 5000000`).

### `pdf:crop`

Payload:

```typescript
{
  pages: ReadonlyArray<number> | "all";
  margin: readonly [number, number, number, number];  // [left, top, right, bottom] in user-units
}
```

- pdf-lib op: set `/CropBox` on each affected page to the inset of
  `/MediaBox` by the margins.
- Incremental save: each affected `/Page` dict gets a new
  `/CropBox`. `/MediaBox` untouched (so the crop is reversible).

### `pdf:watermark`

Payload:

```typescript
{
  pages: ReadonlyArray<number> | "all";
  text?: string;
  imageBytes?: Uint8Array;
  opacity: number;       // 0..1
  position?: "center" | "top-center" | "bottom-center" | "top-left" | "top-right" | "bottom-left" | "bottom-right";
  rotation?: number;     // degrees, default 0 (45 for "DRAFT" style)
  color?: { r: number; g: number; b: number };
}
```

- pdf-lib op: add a Form XObject containing the watermark, reference
  it from each affected page's content stream **as the last drawn
  element** (so it appears on top).
- Incremental save: one new XObject + one updated content stream per
  affected page.

### `pdf:add-page-numbers`

Payload:

```typescript
{
  pages: ReadonlyArray<number> | "all";
  position: "bottom-center" | "bottom-left" | "bottom-right" | "top-center" | "top-left" | "top-right";
  format?: string;       // "{n}" or "{n} of {total}" or "Page {n}"
  fontSize?: number;     // points, default 10
  startAt?: number;      // default 1
  fontFamily?: "Helvetica" | "Times-Roman" | "Courier";
}
```

- pdf-lib op: append `BT … ET` text-show operators to each affected
  page's content stream.

### `pdf:set-metadata`

Payload:

```typescript
{
  title?: string; author?: string; subject?: string;
  keywords?: string; creator?: string; producer?: string;
}
```

- pdf-lib op: update `/Info` dict + (if present) the `/Metadata`
  XMP stream's `<dc:title>` / `<dc:creator>` / etc.
- Incremental save: `/Info` is rewritten; XMP stream is rewritten
  if present.

### `pdf:add-bookmark`

Payload:

```typescript
{ title: string; pageNumber: number; parentId?: NodeId }
```

- Typed model: insert into `outline` tree under `parentId` (or root
  if absent).
- pdf-lib op: walk to the parent outline node, append a new
  `/Outlines` entry pointing to the destination.
- Incremental save: affected `/Outlines` chain is rewritten; the
  `/Outlines` root in the catalog gets a new `/First`/`/Last` if
  this is a top-level insertion.

## Comment commands

Comments are not native PDF annotations; they live in the
`@officeai/comments` model and sync over Y.js. Their handlers update
the typed `PdfDocument.comments` and emit `useCommandBroadcast`
events. Native mirroring (a comment ↔ sticky-note annotation) is
**off by default** and opt-in per document via
`agent.setCommentsMirroring(true)`.

The four comment commands are:

- `pdf:add-comment`
- `pdf:reply-comment`
- `pdf:edit-comment`
- `pdf:resolve-comment`
- `pdf:delete-comment`

Detail: [`agent-commands.md`](./agent-commands.md).

## Incremental save discipline

All operations above are designed for **append-only incremental
update**:

1. The original buffer is untouched.
2. New objects (rewritten `/Page` dicts, new `/Rotate` values,
   new XObjects, new content streams) are appended.
3. A new xref section linking to the new objects is appended.
4. A new trailer with `/Prev` pointing at the previous xref and
   `/Size` updated is appended.
5. The result is `originalBytes + delta`.

The serializer is implemented in
[`packages/pdf/src/serializer/serialize.ts`](../../packages/pdf/src/serializer/serialize.ts).
The single non-negotiable invariant is that every original byte
position 0..originalLength-1 is preserved verbatim in the output.

### When we full-re-serialize instead

Opt-in only, via `agent.exportFile({ rewrite: true })` or
`office-agent pdf optimize`:

- Garbage-collects orphaned objects (deleted pages' content streams).
- Renumbers the xref to be sequential.
- Re-compresses streams.
- **Breaks digital signatures** (the byte range covered by `/ByteRange`
  no longer exists). The Mutation carries `warning:
  "signature-broken-on-rewrite"`.
- The viewer surfaces a confirmation dialog before producing
  rewrite output.

## Never break a signature accidentally

Before applying any edit on a signed PDF, the handler:

1. Walks the form fields for `/Sig` widgets with `/V/ByteRange`.
2. If found, asserts the edit is **append-only** (incremental save).
3. If the user opted into a full rewrite, surfaces the warning.
4. The Mutation carries `warning: "signed-document"` so the UI can
   show the signature pane in "edits will append" mode.

## Failure modes

| Failure                                  | Handling                                                        |
| ---------------------------------------- | --------------------------------------------------------------- |
| pdf-lib throws on `loadDocument`         | Mutation rejected with `error: "parse-failed"`; original buffer untouched. |
| `pages` payload contains out-of-range N  | Rejected with `error: "invalid-page-range"`.                    |
| `order` payload not a valid permutation  | Rejected with `error: "invalid-permutation"`.                   |
| Encrypted PDF without password           | Rejected with `error: "password-required"`; UI prompts.         |
| Source PDF for `insert-pages` fails open | Rejected with `error: "source-parse-failed"`.                   |
| Watermark image bytes corrupt            | Rejected with `error: "image-decode-failed"`.                   |
| Disk full on output write (CLI)          | Reported with `error: "io-failed"`; partial output cleaned up.  |
