# DOCX LLM + MCP page surface (P3.6)

> Status: P3.6 spec. Drives W22 (`snapshotToMarkdown` page sections),
> W23 (MCP `docx_get_pages` / `docx_get_page_text`), W24
> (selection-aware page context helpers). Builds on P3.3 page chunker.
> Followed by P3.7 (real-world acceptance).

## Why

The page chunker (P3.3) gives the editor surface a per-block page
index. But the rest of the system — the LLM agent, the MCP tools, the
copilot prompt — still treats the document as one continuous body.
That breaks the most common AI conversation patterns:

- **"Summarise page 3"** has no answer because page 3 doesn't exist
  in the markdown projection.
- **"Add a heading at the top of page 4"** is impossible because the
  agent can't address paragraphs by page.
- **"What page is the figure on?"** requires every consumer to
  re-implement page resolution against the typed model.

P3.6 closes that gap: the LLM-facing surface always carries page
context, and MCP exposes page-aware reads.

## Scope

### W22 — Markdown projection with page sections

`snapshotToMarkdown(snapshot)` already projects body blocks to GFM
markdown. P3.6 extends it to optionally segment the output by page
chunk:

```ts
export interface SnapshotToMarkdownOptions {
  /**
   * When true, prepend each page's content with a `<!-- page N -->`
   * HTML comment + `## Page N` heading so downstream LLMs can cite
   * pages by number. Pages are computed via `chunkIntoPages` (P3.3)
   * with no measure function, matching the editor's no-render
   * pagination.
   *
   * Defaults to false to preserve the existing markdown shape used
   * by every CLI / MCP consumer today.
   */
  readonly withPageSections?: boolean;
}

export function snapshotToMarkdown(snapshot: DocxSnapshot, options?: SnapshotToMarkdownOptions): string;
```

Output shape with `withPageSections: true`:

```markdown
<!-- page 1 -->

## Page 1

# Title

First paragraph.

<!-- page 2 -->

## Page 2

Second-page paragraph.
```

The HTML comment is the machine-readable anchor (regex-stable across
markdown renderers); the `## Page N` heading is the human-readable
fallback. Both reset on every chunk boundary, including hard
`<w:br w:type="page"/>` breaks and section breaks of `nextPage` type.

Empty pages (e.g. a section break followed immediately by another
section break) still emit the page header — they're a real Word
construct ("blank page") and the LLM should know about them.

### W23 — MCP tools for pages

Add two new tools to `createMcpServer()`:

#### `docx_get_pages`

Lists every page chunk:

```jsonc
{
  "pages": [
    {
      "pageNumber": 1,
      "startBlockIndex": 0,
      "endBlockIndex": 4,
      "trigger": "doc-start",
      "preview": "Title\nFirst paragraph…",
    },
    {
      "pageNumber": 2,
      "startBlockIndex": 5,
      "endBlockIndex": 11,
      "trigger": "page-break",
      "preview": "Second-page paragraph…",
    },
  ],
  "total": 2,
}
```

Schema:

- `pageNumber` — 1-based.
- `startBlockIndex` / `endBlockIndex` — half-open `[start, end)` into
  `snapshot.root.body`. Lets agents target paragraphs by page (e.g.
  "insert a Heading 2 between blocks 5 and 11").
- `trigger` — what started the page. One of:
  - `"doc-start"` (page 1, always),
  - `"page-break"` (`<w:br w:type="page"/>`),
  - `"last-rendered"` (`<w:lastRenderedPageBreak/>`),
  - `"section-break"` (`<w:sectPr w:type="nextPage"/>`),
  - `"measured-overflow"` (only when `chunkIntoPages` ran with a
    measure function — never set by this tool but reserved).
- `preview` — first ≤120 characters of plain text on the page,
  whitespace-collapsed.

#### `docx_get_page_text`

Returns the markdown projection scoped to a single page:

```jsonc
{
  "pageNumber": 3,
  "startBlockIndex": 11,
  "endBlockIndex": 17,
  "format": "markdown",
  "content": "# Section heading\n\nBody…\n",
}
```

Schema:

- `handle` — agent handle.
- `page` — 1-based page number.
- `format` — `"markdown" | "text"`. Defaults to markdown.

Errors:

- `out-of-range` — when `page < 1` or `page > total`.
- `unknown handle` — same as every other docx\_\* tool.

### W24 — Selection-aware page helpers

Add three helpers to `DocxAgent`:

```ts
class DocxAgent {
  /** All page chunks for the current snapshot. */
  getPages(): ReadonlyArray<PageInfo>;

  /** Page number containing the given body paragraph index. */
  pageForParagraph(paragraphIndex: number): number | null;

  /**
   * Markdown projection of the body restricted to the given page
   * number. Returns `null` when the page is out of range.
   */
  getPageMarkdown(pageNumber: number): string | null;
}

interface PageInfo {
  readonly pageNumber: number;
  readonly startBlockIndex: number;
  readonly endBlockIndex: number;
  readonly trigger: PageTrigger;
  readonly preview: string;
}

type PageTrigger = "doc-start" | "page-break" | "last-rendered" | "section-break" | "measured-overflow";
```

These helpers wrap `chunkIntoPages` and the markdown projection;
they exist so MCP tools, the CLI, and the editor all share one
implementation of "what page is paragraph N on?".

## Acceptance criteria

A1. **Markdown page sections.**
`snapshotToMarkdown(snap, { withPageSections: true })` of a document
with one explicit page break in the middle emits `<!-- page 1 -->`
and `<!-- page 2 -->` markers in document order, each followed by a
`## Page N` heading.

A2. **`getPages` shape.**
`agent.getPages()` returns `[{ pageNumber: 1, startBlockIndex: 0,
endBlockIndex: bodyLength, trigger: "doc-start" }]` for a single-page
document, and grows by one entry per `<w:br w:type="page"/>` /
`<w:lastRenderedPageBreak/>` / `nextPage` section break.

A3. **`pageForParagraph` resolves.**
For a 3-page document with breaks at body indices 5 and 12,
`pageForParagraph(0) === 1`, `pageForParagraph(7) === 2`,
`pageForParagraph(20) === 3`. Out-of-range returns `null`.

A4. **MCP `docx_get_pages` matches.**
Calling `docx_get_pages` over the in-memory transport returns the
same data the in-process helper produces, including triggers and
previews.

A5. **MCP `docx_get_page_text` honors range checks.**
`docx_get_page_text` with `page: 999` returns an `isError: true`
content payload mentioning `out-of-range`.

A6. **No regressions.**
`snapshotToMarkdown(snap)` (no options) is byte-identical for any
snapshot to the pre-P3.6 output. Existing 237 docx tests + 51
integration tests + 47 agent tests stay green.

## Out of scope (P3.6)

- Rebuilding the `chunkIntoPages` algorithm to use measured page
  heights server-side (would require headless layout — defer to a
  follow-up).
- A `docx:set-page-margins` command (P4 / R8).
- `docx_replace_in_page` convenience (the LLM can already chain
  `docx_get_page_text` → `docx_apply_command` to do this).
