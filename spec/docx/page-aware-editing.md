# DOCX page-aware editing UX (P3.5)

> Status: P3.5 spec. Drives W18 (Ctrl+Enter inserts a page break),
> W19 (Goto page jump), W20 (page ruler), W21 (PageUp / PageDown
> scrolls a full page). Builds on P3.3 (page chunker, decorations
> plugin) and P3.4 (header/footer authoring + section commands).
> Followed by P3.6 (LLM / MCP surface for pages).

## Why

P3.3 made pages *visible* — the user sees a "Page N of M" divider
between chunks and a status bar. P3.4 made *section chrome* editable
— the user can author headers, footers, and section breaks. P3.5
closes the loop by making *navigation and authoring* page-aware:

1. **Ctrl/Cmd+Enter inserts a hard page break** at the caret. This
   is Word's #1 keyboard shortcut for power users and is the
   only correct way to force "this paragraph starts a new page"
   without inserting a new section.
2. **Goto page** (Ctrl/Cmd+G) jumps the caret + viewport to a
   specific page index. Used by reviewers walking through long
   documents and by the AI agent (P3.6) when it cites a page.
3. **A page ruler** sits above the page surface and shows the
   horizontal margins and tab stops. Mostly visual today, but
   the wiring is the same that R8 / P4 will use to drag tab stops
   and column boundaries.
4. **PageUp / PageDown scroll a full visible page** rather than the
   default browser-driven viewport. Aligns the next page's top
   edge to the editor surface — matches Word's behavior.

## Typed model

No new typed model in P3.5 — the page chunker (P3.3) and the typed
`PageBreakLeaf` / `SectionProperties` (P3.2) already give us
everything we need. P3.5 ships **commands** and **plugins** that
consume the existing types.

## Commands

### `docx:insert-page-break`

Inserts a typed `<w:br w:type="page"/>` at the caret. Splits the
targeting run at the offset, splices a fresh run containing only the
{@link PageBreakLeaf}, and merges the trailing half. Run properties
are inherited from the split run.

```ts
export interface InsertPageBreakPayload {
  /** Stable id of the target paragraph (body-level only). */
  paragraphId: NodeId;
  /** Byte offset inside the paragraph's flat-text. Clamped. */
  offset: number;
}
```

Behavior:

- Errors `unknown-target` if the paragraph is not in the body
  (page breaks inside header/footer paragraphs are valid OOXML
  but P3.5 surfaces only body-level breaks; header/footer page
  breaks are pointless because the part renders on every page).
- Errors `invalid-payload` for negative offsets.
- Dirties `body`.
- Re-emits to a typed `PageBreakLeaf` (already serialized as
  `<w:br w:type="page"/>` via P3.2 / W6).

The page chunker (P3.3) already recognises the new leaf and splits
the document into one more chunk on the next snapshot.

## Editor wiring

### W18 — Ctrl/Cmd+Enter keymap

Add a ProseMirror keymap entry to `mountDocxEditor` (via the
`extraPlugins` channel from P3.3) that intercepts `Mod-Enter` and
dispatches a `docx:insert-page-break` against the current paragraph
id + offset. Implemented as a small plugin
(`apps/web/app/lib/page-break-keymap.ts`) so the docx package stays
PM-agnostic.

Keymap behavior:

- Reads the active paragraph + offset from the PM selection (using
  the same helpers the toolbar uses for `currentParagraphId` /
  `pmSelectionToRange`).
- Falls through to PM's default `Enter` handling if the paragraph
  cannot be located (no selection, inside a non-editable widget,
  etc.) — fail-safe.
- Returns `true` to signal the keystroke was handled, so PM does
  not also insert a paragraph break.

### W19 — Goto page UX

Surface a small "Goto page" affordance in the page status bar
(below the editor surface, P3.3). Clicking the page number
("Page 3 of 26") opens an inline numeric input; submitting jumps
the caret to the start of that page.

Implementation:

- A `gotoPage(view, pageNumber, chunks)` helper that finds the
  first block index of the requested chunk, maps it to a PM doc
  position, and dispatches a transaction setting the selection
  there + scrolling it into view.
- Lives in the same `apps/web/app/lib/page-decorations.ts`
  module as `getPageChunks` and `pageNumberForPos`.

### W20 — Page ruler

A non-editable horizontal bar rendered above the editor surface
(via React, not via PM) that shows:

- Left and right margin guides (drawn from the typed
  `SectionProperties.pgMar`).
- A neutral 0–N ruler in user-selectable units (default inches).

P3.5 ships a read-only ruler. P4 / R8 makes the margin guides
draggable.

Implementation:

- `apps/web/app/editor/PageRuler.tsx` — pure render, no PM coupling.
- Resolves the active section's geometry by walking
  `snapshot.root.body` for the next `section-break` after the caret
  paragraph (or the trailing implicit section).
- Renders inches by default, falls back to centimeters when
  navigator language indicates metric (`!= en-US`).

### W21 — PageUp / PageDown

PM's default keymap binds these to `gapCursor` movement. We
override them to scroll the editor host by one viewport-page
height and snap the caret to the first / last visible block.

- For **PageDown**, we move the caret to the start of the *next*
  page chunk (per `getPageChunks`); for **PageUp**, the previous
  one.
- When already on the first / last page we fall through to PM's
  default so end-of-doc gestures still work.

Implementation lives next to W18's keymap plugin so the editor
gets a single "page-aware keymap" registration.

## Acceptance criteria

A1. **Ctrl+Enter inserts a page break.** Pressing `Mod-Enter` in a
body paragraph calls `docx:insert-page-break`; the snapshot
contains a new `PageBreakLeaf`; the page-decorations plugin
reflects the new chunk count on the next render; the export
contains `<w:br w:type="page"/>`.

A2. **Insert-page-break round-trip.** A fresh page break round-trips
through serialize → parse → identity-equal typed `PageBreakLeaf`.

A3. **Goto page jumps the caret.** Calling `gotoPage(view, 3,
chunks)` moves the selection to the first block of page 3 and
scrolls it into view. The page status bar updates to show
`Page 3 of N`.

A4. **Ruler reads geometry.** With a doc whose section margins are
1 in (1440 twips) all sides, the ruler shows margin guides at
1 in / 6.5 in (US-letter, 8.5 in wide).

A5. **PageDown advances by one page chunk.** Caret in page 1 →
press PageDown → caret lands at the first block of page 2.
Caret in last page → PageDown falls through to default behavior.

A6. **No regressions.** All 231 docx tests + 51 integration tests
+ 47 agent tests stay green. New tests cover the page-break
command + helpers.

## Out of scope (P3.5)

- Drag-to-resize margin guides on the ruler (P4 / R8).
- Tab stop ruler glyphs (P4).
- Multi-column ruler markers (P4).
- Goto-page popup with a search index (P4 — uses a different UX).
- Smooth scrolling / inertia for PageUp/Down (browser default is
  fine for now).
