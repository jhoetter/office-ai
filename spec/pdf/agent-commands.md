# PDF — Agent Commands

> Complete typed `PdfCommand` union with payloads and diff shapes.
> Mirrors the as-built command set in
> [`packages/pdf/src/commands/payloads.ts`](../../packages/pdf/src/commands/payloads.ts).

Cross-references: how commands flow through `CommandBus` in
[`document-model.md`](./document-model.md);
how page commands serialize via `pdf-lib` in
[`editing-pipeline.md`](./editing-pipeline.md);
CLI surface in [`cli.md`](./cli.md);
agent shape in
[`packages/pdf/src/agent/agent.ts`](../../packages/pdf/src/agent/agent.ts).

## Common types

```typescript
import type { Command, NodeId } from "@officeai/core";
import type { PdfRect, PdfRotation } from "@officeai/pdf";

/** Discriminator union for the eleven commands shipped this session. */
export const PDF_COMMAND_TYPES = [
  "pdf:rotate-pages",
  "pdf:set-page-rotation",
  "pdf:reorder-pages",
  "pdf:delete-pages",
  "pdf:set-metadata",
  "pdf:add-bookmark",
  "pdf:add-comment",
  "pdf:reply-comment",
  "pdf:edit-comment",
  "pdf:resolve-comment",
  "pdf:delete-comment",
] as const;
export type PdfCommandType = (typeof PDF_COMMAND_TYPES)[number];
```

All payloads are JSON-serializable. All page addresses are
**1-indexed** to match the PDF spec convention; this is the same
convention `PdfPage.pageNumber` and `enginePage.getPage(n)` use.

## Diff shapes

Every command produces a `Mutation<PdfSnapshot>` whose
`diff: DocumentDiff` carries one or more `DiffChange` entries
(see [`spec/shared/document-model.md`](../shared/document-model.md)).
The shapes below describe the diff each command produces.

## Page operations

### `pdf:rotate-pages`

```typescript
interface RotatePagesPayload {
  pages: ReadonlyArray<number>; // 1-indexed
  delta: 90 | 180 | 270 | -90 | -180 | -270;
}
```

Effect:

- Each `PdfPage[i].rotation = (existing + delta) mod 360` for `i` in
  `pages`.
- Defers a pdf-lib `setRotation` for each affected page; serializer
  applies on `exportFile()`.

Diff:

```json
{
  "kind": "node-updated",
  "nodeId": "<page id>",
  "path": ["pages", "<index>"],
  "field": "rotation",
  "summary": "rotated page 3 by +90° (now 90°)"
}
```

One entry per affected page.

Validation: `pages` must contain values in `[1, numPages]`; `delta`
must be one of the literal values. Otherwise rejected with
`error: "invalid-payload"`.

---

### `pdf:set-page-rotation`

```typescript
interface SetPageRotationPayload {
  pageNumber: number;
  rotation: PdfRotation; // 0 | 90 | 180 | 270
}
```

Effect: absolute set of `PdfPage.rotation`. Otherwise identical to
`pdf:rotate-pages`.

Diff: same `node-updated` on `field: "rotation"`.

---

### `pdf:reorder-pages`

```typescript
interface ReorderPagesPayload {
  /** New page order, expressed as a permutation of 1..N.
   *  Length MUST equal the page count and be a valid permutation. */
  order: ReadonlyArray<number>;
}
```

Effect:

- `pages` is rebuilt in the new order.
- `PdfPage.pageNumber` is reassigned 1..N.
- `PdfAnnotation.pageNumber` and `PdfFormField.pageNumber` are
  remapped via the permutation.
- `PdfComment.pageNumber` is remapped.
- `PdfOutlineNode.pageNumber` is remapped.

Diff:

```json
{
  "kind": "node-updated",
  "nodeId": "root",
  "path": ["pages"],
  "field": "order",
  "summary": "reordered pages: [1,3,2,4,5,…]"
}
```

Validation: `order` must be a permutation of `1..numPages`.
Otherwise `error: "invalid-permutation"`.

---

### `pdf:delete-pages`

```typescript
interface DeletePagesPayload {
  pages: ReadonlyArray<number>; // 1-indexed
}
```

Effect:

- Affected pages dropped from `pages`.
- `pageNumber` reassigned on the survivors.
- Annotations / form fields / comments / outline entries on dropped
  pages are dropped.
- pdf-lib `removePage` is queued; serializer applies on
  `exportFile()`.

Diff: one `node-deleted` per dropped page + one `node-updated` on
`pages` summarising the new count.

Validation: `pages` non-empty, in `[1, numPages]`, and not all
pages of the document (the result must have ≥ 1 page).

---

## Metadata + outline

### `pdf:set-metadata`

```typescript
interface SetMetadataPayload {
  title?: string;
  author?: string;
  subject?: string;
  keywords?: string;
  creator?: string;
  producer?: string;
}
```

Effect: each present field is set on `PdfDocument.metadata` (and on
`/Info` + the XMP `/Metadata` stream on serialize). Absent fields
are left unchanged.

Diff: one `node-updated` per present field on
`path: ["metadata"]`.

---

### `pdf:add-bookmark`

```typescript
interface AddBookmarkPayload {
  title: string;
  pageNumber: number;
  parentId?: NodeId; // omit for root-level
}
```

Effect:

- New `PdfOutlineNode` inserted under `parentId` (or root).
- `id` is minted; `pageNumber` resolves to the destination at write
  time.

Diff:

```json
{
  "kind": "node-inserted",
  "nodeId": "<new outline id>",
  "path": ["outline", …],
  "summary": "added bookmark 'Chapter 3' on page 42"
}
```

Validation: `pageNumber` in `[1, numPages]`; `parentId` (if present)
must resolve to an existing outline node.

---

## Comments (Office-AI overlay; not native PDF annotations by default)

Comments live in the `@officeai/comments` model and sync over Y.js.
They are anchored to the PDF via `PdfRegionAnchor` (page +
normalized rect). Detail:
[`document-model.md` § Comments anchored as `pdf-region`](./document-model.md).

Mirroring to native sticky-note annotations is opt-in via
`agent.setCommentsMirroring(true)`; off by default to avoid
polluting the byte-preservation invariant.

### `pdf:add-comment`

```typescript
interface AddCommentPayload {
  id?: NodeId; // optional pre-minted id (for replay)
  author: string;
  text: string;
  pageNumber: number;
  normalizedRect: PdfRect; // 0..1 in unrotated user-space
}
```

Effect: new `PdfComment` appended to `comments`.

Diff:

```json
{
  "kind": "node-inserted",
  "nodeId": "<comment id>",
  "path": ["comments"],
  "summary": "added comment by Alice on page 5"
}
```

---

### `pdf:reply-comment`

```typescript
interface ReplyCommentPayload {
  id?: NodeId;
  parentId: NodeId; // the comment being replied to
  author: string;
  text: string;
}
```

Effect: new `PdfComment` with `parentId` set; inherits the parent's
`pageNumber` + `normalizedRect`.

Diff: `node-inserted` under `comments`.

Validation: `parentId` must resolve to an existing comment.

---

### `pdf:edit-comment`

```typescript
interface EditCommentPayload {
  commentId: NodeId;
  text: string;
}
```

Effect: `PdfComment.text` updated. Author and timestamps unchanged
(an edit log is kept on the realtime side, not in the typed model).

Diff:

```json
{
  "kind": "node-updated",
  "nodeId": "<comment id>",
  "path": ["comments", "<index>"],
  "field": "text",
  "summary": "edited comment by Alice"
}
```

Validation: `commentId` must resolve.

---

### `pdf:resolve-comment`

```typescript
interface ResolveCommentPayload {
  commentId: NodeId;
  resolved: boolean;
}
```

Effect: `PdfComment.resolved = resolved`.

Diff: `node-updated` on `field: "resolved"`.

Validation: `commentId` must resolve.

---

### `pdf:delete-comment`

```typescript
interface DeleteCommentPayload {
  commentId: NodeId;
}
```

Effect: comment + all replies (recursive) dropped.

Diff: one `node-deleted` per dropped comment.

Validation: `commentId` must resolve.

---

## Pending mutations and the human review flow

Commands dispatched with `source: "agent"` (any LLM-originated
mutation) land in the **pending queue** rather than being applied
immediately. The viewer surfaces them as proposals (dashed outline
on annotations, "preview" badge on page operations). The human
approves / rejects per the same tri-state semantics DOCX/XLSX/PPTX
already use.

```typescript
agent.applyCommand({ type: "pdf:add-bookmark", source: "agent", payload: { … } });
// → returns Mutation with status "pending"
agent.getPendingMutations(); // → [Mutation, …]
agent.approveMutation(id);   // → moves to approved
agent.rejectMutation(id);    // → drops with no effect on snapshot
```

Approval flips the snapshot to the "with this mutation applied" state
and bumps the revision. Rejection discards the mutation. Both produce
events on `agent.subscribe()`.

## Rollback

```typescript
agent.rollback(toRevision: number): void;
```

Roll back the working snapshot to a prior revision. Used by the
"discard changes" UI affordance and by the test harness.

## Future / spec-only commands (not P0)

Documented for completeness; not implemented this session. These will
expand the `PDF_COMMAND_TYPES` union when shipped.

| Command                   | Payload sketch                                          |                 Status                  |
| ------------------------- | ------------------------------------------------------- | :-------------------------------------: | -------------------------- | ------------ |
| `pdf:add-annotation`      | typed `PdfAnnotation` insert                            |                 **P1**                  |
| `pdf:update-annotation`   | annotation id + diff                                    |                 **P1**                  |
| `pdf:delete-annotation`   | annotation id                                           |                 **P1**                  |
| `pdf:fill-form`           | `Record<fieldName, value>` + `flatten?: boolean`        | **P0** (CLI; typed command lands W4–W5) |
| `pdf:reset-form`          | (no payload)                                            |              **P0** (CLI)               |
| `pdf:flatten-form`        | (no payload)                                            |              **P0** (CLI)               |
| `pdf:redact`              | `{ rects: PdfRect[]; replacement?: RGB }`               |              **P0** (CLI)               |
| `pdf:apply-redactions`    | (no payload)                                            |              **P0** (CLI)               |
| `pdf:add-watermark`       | text or image + opacity + pages                         |              **P0** (CLI)               |
| `pdf:add-page-numbers`    | position + format + start                               |              **P0** (CLI)               |
| `pdf:crop`                | `{ pages, margin: [l,t,r,b] }`                          |              **P0** (CLI)               |
| `pdf:insert-pages`        | `{ at, sourceBuffer, sourcePages? }`                    |              **P0** (CLI)               |
| `pdf:merge`               | `{ files: Uint8Array[] }` (pure helper, not bus-routed) |              **P0** (CLI)               |
| `pdf:split`               | `{ by: "range"                                          |               "bookmark"                | "size"; … }` (pure helper) | **P0** (CLI) |
| `pdf:add-text-overlay`    | page + rect + text + style                              |                 **P2**                  |
| `pdf:add-image-overlay`   | page + rect + image bytes                               |                 **P2**                  |
| `pdf:flatten-annotations` | (no payload)                                            |                 **P2**                  |

The CLI exposes the **P0** items in this table (W4–W5 wave) before
they have typed `pdf:*` commands routed through the bus. Once W4–W5
land, the typed commands take over and the CLI delegates to them
internally.
