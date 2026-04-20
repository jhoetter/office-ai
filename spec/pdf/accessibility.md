# PDF — Accessibility

> Tagged-PDF reading order, viewer-chrome WCAG 2.2 AA conformance,
> reflow mode for low-vision users, and a complete keyboard map.

Cross-references: text-layer for screen-reader integration in
[`text-layer.md`](./text-layer.md);
keyboard catalogue in [`keyboard-shortcuts.md`](./keyboard-shortcuts.md);
dark mode in [`dark-mode.md`](./dark-mode.md);
acceptance criteria in [`acceptance-criteria.md`](./acceptance-criteria.md).

## Goals

1. **A blind reviewer can do everything a sighted reviewer can.**
   Outline navigation, jump to page, read structured content, fill
   forms, hear / write annotations.
2. **The viewer chrome passes WCAG 2.2 AA** measured by axe-core +
   Lighthouse on every CI run.
3. **Reflow mode** lets a low-vision user read the structured content
   in a single column at any font size, ignoring the original layout.
4. **Reduced motion** is respected (`prefers-reduced-motion`).
5. **High contrast** is respected (`prefers-contrast`).

## Tagged-PDF reading

A tagged PDF carries a `/StructTreeRoot` whose nodes describe the
document's logical structure (`/H1`, `/P`, `/Figure`, `/Table`, `/L`,
`/LI`, `/Caption`, …) with optional `/Alt` text, `/ActualText`, and
`/Lang`. The struct tree drives reading order for screen readers.

### Engine projection

`packages/pdf-engine` exposes the per-page struct tree:

```typescript
export interface PdfStructNode {
  readonly role: string;                         // "/H1", "/P", "/Figure", …
  readonly lang?: string;                        // e.g. "en", "de"
  readonly alt?: string;                         // alt text
  readonly actualText?: string;                  // for ligatures, OCR, replaced text
  readonly children: ReadonlyArray<PdfStructNode>;
  /** Indices into the page's text-layer items, when this node is a leaf. */
  readonly textRange?: { readonly start: number; readonly end: number };
}

export interface PdfStructTree {
  readonly perPage: ReadonlyArray<{ pageNumber: number; root: PdfStructNode }>;
  readonly documentLang?: string;
}
```

PDF.js gives us this for free via `getStructTree()`. PDFium has the
same primitives via `FPDF_StructTree_*` — the wrapper produces the
same shape.

### DOM mapping

For each tagged page, the viewer renders a hidden ARIA-labelled DOM
tree aligned with the canvas:

- `/H1` … `/H6` → `<h1>` … `<h6>` (with `aria-level` for `/H7+`).
- `/P` → `<p>`.
- `/L` → `<ul>` (or `<ol>` if `/ListNumbering` is set).
- `/LI` → `<li>`.
- `/Figure` → `<figure>` with `<img alt="…">` placeholder; `aria-label`
  from `/Alt` or `/ActualText`.
- `/Table` / `/TR` / `/TH` / `/TD` → `<table>` / `<tr>` / `<th>` /
  `<td>`.
- `/Caption` → `<caption>` or `<figcaption>` depending on parent.
- `/Link` → `<a>` with `href` from the link annotation.
- `/Code` → `<code>`.
- `/BlockQuote` → `<blockquote>`.

The hidden DOM is positioned with `clip-path: inset(100%)` so it
remains in the accessibility tree but invisible. Screen readers
(VoiceOver, NVDA, TalkBack) navigate it linearly in document order.

### Untagged PDFs

If `/StructTreeRoot` is absent, the viewer falls back to **text-layer
reading order** (the engine's best-effort linearization of text
items). Headings are heuristic: large + bold runs starting at the
left margin become `<h2>`. Lists are heuristic: lines starting with
`•`, `-`, `*`, or `1.` become `<ul>` / `<ol>`. The viewer surfaces a
banner: **"This document is not tagged. Reading order is approximate.
Use the auto-tag tool to propose a structure (deferred)."**

## Reflow mode

Toggle via `View → Reflow` or `Ctrl+Shift+R` (overlap with rotate is
resolved by `Ctrl+Alt+R` for rotate, `Ctrl+Shift+R` for reflow — see
[`keyboard-shortcuts.md`](./keyboard-shortcuts.md)).

Reflow mode:

1. Hides the canvas.
2. Renders the struct-tree DOM at full visibility (no `clip-path`).
3. Applies a single-column layout with adjustable font size (default
   18 px, ranges 14 – 36 px in 2 px steps via `Cmd+=` / `Cmd+-`).
4. Preserves headings / lists / images (with alt text) / tables.
5. Page breaks become subtle horizontal rules with "Page N" labels.
6. Annotations are listed in a separate sidebar pane in document
   order, each with its anchor's nearest heading as context.

Reflow degrades gracefully on untagged PDFs: it renders the heuristic
struct DOM. If even the heuristic fails (pure scan with no OCR), the
viewer prompts the user to run OCR first.

## Keyboard navigation

The full keymap lives in
[`keyboard-shortcuts.md`](./keyboard-shortcuts.md). For accessibility,
the **non-negotiable** subset:

| Action                       | Default key                     | Reachable?            |
| ---------------------------- | ------------------------------- | --------------------- |
| Skip to toolbar              | `Ctrl+,`                        | Yes — `<a class="skip-link">` revealed on focus |
| Skip to sidebar              | `Ctrl+1`                        | Yes                   |
| Skip to page content         | `Ctrl+2`                        | Yes                   |
| Next / previous page         | `PageDown` / `PageUp`           | Yes                   |
| First / last page            | `Home` / `End`                  | Yes                   |
| Jump to page                 | `Ctrl+G`                        | Yes — opens GotoDialog with `aria-modal` |
| Open outline                 | `Ctrl+Shift+O`                  | Yes                   |
| Open thumbnails              | `Ctrl+Shift+T`                  | Yes                   |
| Open annotations panel       | `Ctrl+Shift+A`                  | Yes                   |
| Find                         | `Ctrl+F`                        | Yes                   |
| Find next / previous         | `Ctrl+G` / `Ctrl+Shift+G` (with find open) | Yes        |
| Zoom in / out / reset        | `Ctrl++` / `Ctrl+-` / `Ctrl+0`  | Yes                   |
| Toggle reflow                | `Ctrl+Shift+R`                  | Yes                   |
| Toggle dark mode             | `Ctrl+Alt+D`                    | Yes                   |
| Annotate (highlight)         | `Ctrl+Shift+H` (after select)   | Yes                   |
| Add comment at focused page  | `Ctrl+Alt+C`                    | Yes                   |
| Print                        | `Ctrl+P`                        | Yes                   |

Every action **must** be reachable from the keyboard. CI gate: an
e2e test sweeps every toolbar button and asserts focusable + has
`aria-label`.

## ARIA contract

- All buttons carry `aria-label` (i18n via `useTranslator()`).
- Toggle buttons (zoom mode, view mode, dark mode) carry
  `aria-pressed`.
- Sidebar tabs are `role="tablist"` / `role="tab"` / `role="tabpanel"`.
- Modal dialogs (`GotoDialog`, password prompt, signature panel) use
  `role="dialog"` + `aria-modal="true"` + focus trap.
- The page list is `role="list"` with each page as `role="listitem"`
  carrying `aria-label="Page N of M"`.
- The text layer is `aria-hidden="false"` so screen readers read it;
  the canvas is `aria-hidden="true"`.
- The struct-tree-layer carries `role="document"` at the page root.

## Live regions

- Search hit count: `<div aria-live="polite">` updates as the query
  refines.
- Page-change indicator: `aria-live="polite"` announces "Page 12 of
  340" when the user pages.
- Annotation creation: `aria-live="polite"` announces "Highlight
  added on page 5".
- Comment-thread updates from realtime: `aria-live="polite"` for
  resolved / replied; `aria-live="assertive"` is **never** used (too
  intrusive).

## `prefers-reduced-motion`

When set:

- Page-transition animations disabled (instant snap instead of
  smooth scroll on `Page Up`/`Down`).
- Sidebar slide-in disabled.
- Annotation pulse-on-create disabled.
- Toast slide-up disabled (fade-only).

## `prefers-contrast: more`

When set:

- Toolbar borders gain `2px` instead of `1px`.
- Focus rings widen to `3px outline-offset: 2px`.
- Text-layer search-hit highlight uses solid background instead of
  translucent.
- Sidebar separators darken.

## CI gates

- `axe-core` run via Playwright on every PR. Fails on any
  WCAG 2.2 AA violation in the viewer chrome.
- `Lighthouse` accessibility audit ≥ 95 on `/pdf-viewer` route.
- Reflow-mode visual snapshot on 5 fixtures (untagged + tagged + scan).
- Screen-reader smoke: VoiceOver scripted test reads page 1 of a
  tagged fixture, asserts heading + body text in correct order.

## Documentation for end users

The viewer ships an in-product **"Accessibility"** menu item that
opens a panel listing every keyboard shortcut, the current
high-contrast / dark-mode / reflow / reduced-motion state, and a
direct link to the help docs. Discoverability is itself an
accessibility feature.
