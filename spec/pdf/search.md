# PDF — Search

> Per-page text index built on parse, match highlighting in the
> text-layer, regex / whole-word / case toggles. Fuzzy and semantic
> search are P1 / P2.

Cross-references: text-layer overlay in
[`text-layer.md`](./text-layer.md);
agent surface in [`agent-commands.md`](./agent-commands.md) and
[`cli.md`](./cli.md);
keyboard shortcuts in [`keyboard-shortcuts.md`](./keyboard-shortcuts.md);
performance in [`performance.md`](./performance.md).

## Index

The per-page text index is built **once at parse time** and cached
on the snapshot. It is the same `PdfPage.text` field exposed on the
typed model:

```typescript
export interface PdfPage {
  // …
  readonly text: string;          // best-effort reading-order projection
  readonly hasTextLayer: boolean;
}
```

The text is the engine's `getTextContent().plain` projection. For
PDF.js it's reading-order linearization with `\n` between text items
that have `hasEol`. For PDFium it's the equivalent `FPDFText_*`
projection.

For pages with `hasTextLayer === false` (scans), the index is empty
until the user runs OCR
([`text-layer.md` § OCR fallback](./text-layer.md)). After OCR the
parser re-reads the page and refreshes the index.

## Query model

```typescript
export interface PdfSearchSpec {
  readonly query: string;
  readonly caseSensitive?: boolean;     // default false
  readonly regex?: boolean;             // default false; query is treated as a regex
  readonly wholeWord?: boolean;         // default false; wraps query in \b…\b
  readonly pageRange?: {
    readonly start: number;             // 1-indexed inclusive
    readonly end: number;
  };
}

export interface PdfSearchResult {
  readonly pageNumber: number;
  readonly start: number;               // char offset into PdfPage.text
  readonly end: number;
  readonly preview: string;             // ±40 chars around the match, with ellipsis
  readonly match: string;               // exact matched substring
}
```

Implementation lives on
[`PdfAgent.search()`](../../packages/pdf/src/agent/agent.ts) and is
exercised by both the in-product search panel and the
`office-agent pdf search` CLI.

## Algorithm

Search runs synchronously over the in-memory index:

1. **Build the regex.**
   - If `regex: true`, use the query as-is.
   - Else escape regex metacharacters in the query.
   - If `wholeWord: true`, wrap with `\b` … `\b`.
   - Flags: `g` always; `i` unless `caseSensitive: true`.
   - On invalid regex, throw `pdf:search invalid pattern: <message>`.
2. **Walk pages.** For each page in `pageRange` (or all pages),
   `re.exec(page.text)` in a loop, pushing each match.
3. **Build preview.** Slice `±40` chars around the match, prepend `…`
   if not at start, append `…` if not at end.
4. **Return.** No async; ≤ 200 ms for 500-page docs (verified in the
   perf budget).

## UI integration

The search panel in `PdfSidebar` shows:

- Query input with three toggle buttons: case (`Aa`), whole word (`ab`),
  regex (`.*`).
- Match count: "12 matches in 3 pages".
- A virtualized list of matches: each row shows page number, preview
  with the match bolded, and click-to-jump.
- Prev / next buttons (`F3` / `Shift+F3` keyboard) that scroll the
  page into view and highlight the *current* hit with a brighter
  background.

The header search input (Ctrl+F) is a compact version with the
toggles in a popover; results jump directly without opening the
sidebar.

## Highlighting in the text-layer

Match highlighting splits affected `<span>`s at the match boundaries
and adds `class="pdf-search-hit"` to the match characters. The
*current* hit additionally gets `class="pdf-search-hit current"` and
auto-scrolls into view. Detail:
[`text-layer.md` § Search-hit highlighting](./text-layer.md).

## Search across multiple documents (P2)

Deferred. The shape is documented now to keep the API consistent:

```typescript
PdfSearchHub.searchAll({ query, agents: PdfAgent[] }) → Promise<{ agentId: string; results: PdfSearchResult[] }[]>
```

The `PdfSearchHub` would parallelize across the open agents and
present a unified result list grouped by document.

## Fuzzy search (P1)

Deferred. The current synchronous regex path serves the common case
(typed query → exact match). For fuzzy search the design is:

- Pre-build a per-page **trigram index** at parse time (cost: ~1 ms
  per page on average text density).
- On query, score candidate matches by Levenshtein distance over
  trigram-shortlisted regions; threshold default 2 edits.
- Surface results with a confidence score; the search panel shows
  the score badge.

## Semantic search (P2)

Deferred per the in-product AI-out-of-scope decision in
[`prompt-pdf.md`](../../prompt-pdf.md). The CLI surface
(`office-agent pdf search`) returns text-search results only;
agents that want semantic search call their own embedding pipeline
on top of `office-agent pdf chunk`.

## CLI

```
office-agent pdf search --file f.pdf --query "cancellation" [--regex] [--case] [--whole-word] [--pages 1-50]
```

Output (JSON, default):

```json
{
  "format": "pdf-search-result/v1",
  "query": "cancellation",
  "regex": false,
  "caseSensitive": false,
  "wholeWord": false,
  "totalMatches": 12,
  "results": [
    {
      "pageNumber": 5,
      "start": 1234,
      "end": 1246,
      "match": "cancellation",
      "preview": "…upon written notice of cancellation no later than 30 days…"
    }
  ]
}
```

`--format text` switches to a grep-style line output for shell
piping. `--format md` emits a Markdown table.

## Performance budgets

| Metric                                          | Target           |
| ----------------------------------------------- | ---------------- |
| First hit on 500-page text-heavy document       | < 200 ms p95     |
| All hits on 500-page text-heavy document        | < 800 ms p95     |
| First hit on 1000-page document, regex query    | < 500 ms p95     |
| Memory overhead per page (in `PdfPage.text`)    | ~3-5 KB typical  |
| Parse-time cost of building the index           | already amortized into `getTextContent()` |

## Failure modes

| Case                                          | Handling                                                  |
| --------------------------------------------- | --------------------------------------------------------- |
| Invalid regex                                 | Throw `pdf:search invalid pattern: <message>`; UI surfaces an inline error. |
| Empty query                                   | Returns `[]` immediately; no UI noise.                    |
| Query in pure-scan PDF (no text layer)        | Returns `[]` for all pages with `hasTextLayer === false`; banner offers OCR. |
| Page range out of bounds                      | Clamped to `[1, numPages]`; no error.                     |
| Regex with catastrophic backtracking          | Worker timeout at 5 s; user sees "Search took too long; refine your query." |
