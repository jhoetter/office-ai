# Shared Comments

> Format-agnostic vocabulary for threaded comments. The DOCX, XLSX,
> and PPTX editors each keep their own native comment model, but
> expose a `CommentsProvider` adapter so a single React UI
> (`@officeai/ui` `CommentsSidebar`) drives all three.

## Goals

1. **One sidebar, three products.** The same `CommentsSidebar` —
   thread cards, reply input, resolve / reopen, delete — renders in
   the DOCX, XLSX, and PPTX editors and behaves identically.
2. **Lossless round-trip.** Each adapter persists threading
   (`parentId`), resolved state, and creation timestamps through the
   format's native part(s) so a deck/spreadsheet/document opened in
   the corresponding Microsoft client preserves every comment.
3. **No format-specific UI in the shared layer.** The shared sidebar
   never touches DOCX paragraph anchors, XLSX A1 refs, or PPTX EMU
   pins; it only renders the canonical `CommentBody`. Anchor-aware
   rendering (margin balloons, free-pin overlays) lives in each
   editor's own UI module.
4. **Provider-driven dispatch.** The sidebar calls `provider.add /
   reply / resolve / delete`; adapters translate to the appropriate
   command-bus command (`docx:reply-comment`, `xlsx:add-comment`,
   `pptx:resolve-comment`, …) so all mutations still flow through the
   single `CommandBus` audit trail.

## Canonical types

Defined in `@officeai/comments` (`src/types.ts`):

```typescript
export type CommentText = string;

export type CommentAnchor =
  | { kind: "docx-range"; paragraphIndex: number; range?: unknown }
  | { kind: "xlsx-cell"; sheet: string; ref: string }
  | { kind: "pptx-pin"; slideIndex: number; xEmu: number; yEmu: number; shapeId?: string }
  | { kind: "none" };

export interface CommentBody {
  id: string;
  author: string;
  text: CommentText;
  createdAt?: string;
  resolved?: boolean;
  parentId?: string;
  anchor: CommentAnchor;
  /** Opaque back-reference to the format-specific comment object. */
  nativeRef?: unknown;
}

export interface CommentThread {
  parent: CommentBody;
  replies: ReadonlyArray<CommentBody>;
}

export interface CommentsProvider {
  threads(): ReadonlyArray<CommentThread>;
  add(input: { author: string; text: string; anchor: CommentAnchor }): Promise<string>;
  reply(input: { parentId: string; author: string; text: string }): Promise<string>;
  resolve(commentId: string, resolved: boolean): Promise<void>;
  delete(commentId: string): Promise<void>;
  edit?(commentId: string, text: string): Promise<void>;
  onScrollTo?(commentId: string): void;
}
```

The `range` field on `docx-range` is opaque so adapters can stash the
full `DocxSelection` (start + end positions) without leaking the DOCX
model into the shared layer.

## Threading helpers

`groupThreads(bodies)` partitions a flat list of `CommentBody` into
`CommentThread`s:

- Top-level comments (no `parentId`) become thread parents in their
  original order.
- Replies (with a `parentId` matching a known parent) attach to the
  parent's `replies` list, preserving document order.
- Orphan replies (parent missing from the list) are surfaced as their
  own top-level threads instead of silently disappearing.

The helper is pure and exported from `@officeai/comments` and
re-exported from `@officeai/ui` so adapters can compose threads
themselves if they don't want to hand the provider to the sidebar.

## Anchor semantics per format

| Format | Anchor variant | Native location |
|---|---|---|
| DOCX | `docx-range` | `<w:commentRangeStart/>` … `<w:commentRangeEnd/>` straddling runs |
| XLSX | `xlsx-cell` | `<comment ref="B7" authorId="…"/>` in `xl/comments{N}.xml` |
| PPTX | `pptx-pin` | `<p:cm pos="x,y"/>` in `ppt/comments/comment{N}.xml`, optionally tied to a shape via `officeai-shapeId` extension |

Each adapter is free to add a `nativeRef` so a future format-aware
view (e.g. revision-style comment cards in DOCX) can recover the
native blob without re-querying the agent.

## Threading + resolved state on the wire

| Format | Threading carrier | Resolved carrier |
|---|---|---|
| DOCX | `<w:commentReply w:id="…" w:parentId="…"/>` (W15 `commentsExtended`) | `<w15:commentEx w15:done="1"/>` |
| XLSX | `officeai-parentId` attr on `<comment>` (round-trips through Excel as opaque attr) | `officeai-resolved="1"` attr |
| PPTX | `officeai-parentId` attr on `<p:cm>` extension list | `officeai-resolved="1"` attr |

XLSX and PPTX use custom `officeai-*` attributes because OOXML's
threaded-comment parts are optional and not preserved by older
clients. Modern Excel additionally writes its own
`xl/threadedComments/*` parts; the parser tolerates either source.
The custom attributes survive a round-trip through Microsoft Office
because both formats preserve unknown attributes on known elements.

## Adapter contracts

Each editor exposes a small adapter that normalises its native
comments into the canonical shape:

| Editor | Adapter |
|---|---|
| DOCX | `apps/web/app/editor/docxCommentsProvider.ts` |
| XLSX | `apps/web/app/xlsx-editor/xlsxCommentsProvider.ts` |
| PPTX | `apps/web/app/pptx-editor/pptxCommentsProvider.ts` |

Each adapter:

- Wraps an agent (`DocxAgent` / `XlsxAgent` / `PptxAgent`).
- Implements `threads()` by reading the agent snapshot and calling
  `groupThreads`.
- Implements `add / reply / resolve / delete` by dispatching the
  matching `format:verb-comment` command via `agent.applyCommand`.
- Wires `onScrollTo` to the editor's native "click to locate"
  affordance:
  - **DOCX** scrolls the page-card to the highlight mark and adds
    `.pm-comment-flash` for ~1.4 s (yellow background fade).
  - **XLSX** moves the active selection to the anchored cell, scrolls
    the grid so the cell is visible (with a small margin off the
    headers), and paints a yellow flash overlay (`.xlsx-comment-flash`,
    keyframes in `apps/web/app/globals.css`) that fades over ~1.4 s.
  - **PPTX** sets the canvas selection to the anchored shape (when
    the comment carries `shapeId`) or falls back to a 1-inch flash box
    around the pin coordinates. The flash is a self-contained inline
    SVG `<animate>` so it doesn't rely on the host stylesheet.

  All three editors auto-open the comments rail when `onScrollTo`
  fires, so a click from the dashboard / outline still surfaces the
  conversation.
- Optionally accepts a `onToast` hook so the editor surface can
  surface success/error toasts without duplicating try/catch logic.

## Composer flow

Adding a new top-level comment is editor-specific because it needs a
fresh anchor:

| Editor | Composer surface |
|---|---|
| DOCX | Floating `CommentComposer` popover anchored to the live PM selection; submits with a `DocxSelection` range. |
| XLSX | Inline `CommentComposer` in the sidebar; uses the active cell selection as the `xlsx-cell` anchor. |
| PPTX | Click-to-pin overlay on the slide canvas; commits a `pptx-pin` anchor (optionally bound to the clicked shape). |

The composer dispatches through `provider.add(...)` so the sidebar
re-renders immediately after the agent's snapshot updates.

## What this spec does NOT cover

- Rich-text comment bodies. Adapters flatten to plain text for the
  shared UI; rich content stays on `nativeRef`.
- Margin balloons / overlay rendering. DOCX renders revision-style
  balloons (`TrackedChangesUI`) and PPTX renders free-pin overlays
  (`SlideCanvas`); both are editor-specific surfaces sitting beside
  the shared sidebar.
- Author identity / mentions / notifications. The shared layer carries
  a single `author` display name; product layers are free to map that
  to richer identity objects.

## Reading order

1. `packages/comments/src/types.ts` — the contract.
2. `packages/comments/src/threads.ts` — `groupThreads` helper.
3. `apps/web/app/editor/docxCommentsProvider.ts` — DOCX adapter.
4. `apps/web/app/xlsx-editor/xlsxCommentsProvider.ts` — XLSX adapter.
5. `apps/web/app/pptx-editor/pptxCommentsProvider.ts` — PPTX adapter.
6. `packages/ui/src/primitives/comments-sidebar.tsx` — the shared UI
   that consumes a provider.
