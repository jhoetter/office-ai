# DOCX phase summary (P0 → P3)

> Status: 2026-04-18. Three follow-on phases (P1, P2, P3) shipped on top
> of the original P0 session. This file is the high-level overview;
> [`docs/build-log/docx.md`](build-log/docx.md) is the
> chronological per-batch narrative with decisions, deviations and
> caveats.
>
> Sister documents (other formats): [`docs/session-summary-pptx.md`](./session-summary-pptx.md) — PowerPoint (slides) phase summary; [`docs/pptx-architecture-notes.md`](./pptx-architecture-notes.md) — where slides diverge from DOCX, and why; [`docs/build-log/{pptx,xlsx,quality-gates}.md`](./build-log/) — live per-format decision logs.

## TL;DR

The DOCX track now has:

- A typed in-memory model that round-trips real-world Word, LibreOffice,
  Google Docs and Pages output **byte-identically** for every untouched
  part. Touched parts re-emit through a typed serializer that stays
  schema-valid against the ECMA-376 XSDs (gated in CI).
- A full headless `DocxAgent` with a 20+ command surface covering text
  edits, paragraph formatting, lists, hyperlinks, tables, images,
  comments, tracked changes, headers/footers, page breaks, section
  breaks and page-number fields.
- A Word-flavoured browser editor on ProseMirror with a paginated view
  (page chunker + decoration widgets), a selection-aware toolbar that
  reflects inherited style-cascade values, drag/drop image insertion, a
  comment composer, a tracked-changes side panel, and page-aware
  navigation (`Mod-Enter`, `PageUp` / `PageDown`, click-to-goto).
- An `office-agent` CLI plus an MCP server (`docx_*` tool family)
  exposing the same agent surface to LLMs, including
  `docx_get_pages` / `docx_get_page_text` for page-aware prompting.

## Phase map

| Phase | Theme                                         | Build-log section                                                                       | Highlights                                                                                                                                                                                                                                                                              |
| ----- | --------------------------------------------- | --------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| P0    | "Walking skeleton" of the AI-native editor    | `## Decisions` → `## Validation summary (2026-04-17)`                                   | Monorepo restructure, OOXML I/O with byte-preservation, command bus with approve/reject/rollback, six P0 commands, headless `DocxAgent`, ProseMirror renderer, `office-agent` CLI, web editor surface.                                                                                  |
| P1    | Real-world coverage + tables / images / lists | `## P1.1 — W1` → `## P1.6 — typed table rendering`                                      | Real-world fixtures + LibreOffice CI smoke, multi-paragraph range edits, MCP server, typed headers/footers + tracked-change resolution, web toolbar / sidebar parity, perf + license CI gates, typed tables, image insertion, OOXML XSD validation gate, lists + hyperlinks.            |
| P2    | Polish + LLM dispatch + research synthesis    | `## P2.1 + P2.2` → `## P2.6 — Eigenpal deep-dive synthesis`                             | Selection-aware toolbar (font / size / colour / alignment / indent / list buttons reflect the caret), TOC / SDT unwrap so real-world docs render structurally, real `<img>` rendering with drag/drop/paste, comment composer + selection-aware LLM dispatch, eigenpal architecture deep-dive driving the P3 plan. |
| P3    | Word-UX parity                                | `## P3 — Word-UX parity`                                                                | Style cascade + toolbar inheritance, typed page geometry + section model, paged renderer, header/footer authoring (commands + side panel), page-aware editing UX (`Mod-Enter`, goto-page, ruler, `PageUp` / `PageDown`), LLM/MCP page surface (`docx_get_pages`, `docx_get_page_text`).  |

The full per-batch log lives in
[`docs/build-log/docx.md`](build-log/docx.md). Each section follows
the same shape: **what shipped → decisions → caveats**.

## Test counts (2026-04-18)

| Package                       | Tests   |
| ----------------------------- | ------: |
| `@officeai/core`              |      12 |
| `@officeai/docx`              |     249 |
| `@officeai/agent` (CLI + MCP) |      50 |
| `@officeai/integration-tests` |      86 |
| **Docx-relevant total**       | **397** |

(The repo also runs 617 tests in `@officeai/xlsx`; see
[`docs/build-log/xlsx.md`](build-log/xlsx.md) for that track.)

`make verify` is green: typecheck, lint, build, format, all tests, OOXML
schema validation, and the license scan.

## What's still deferred

Tracked in the build log per phase. The current shortlist (P4 candidates
fed by `spec/docx/eigenpal-synthesis.md` R8–R12 and the P3 caveats):

- **Measured pagination.** The chunker honours hard `<w:br w:type="page"/>`
  and `<w:lastRenderedPageBreak/>` hints; it does not yet run a layout
  pass that splits overflowing paragraphs.
- **In-page header/footer focus model.** P3.4 shipped a side-panel
  authoring UX wiring every typed command. The "click into a header
  preview, body greys out, toolbar retargets" visual end-state is P4
  polish.
- **Different-odd-even / restart-numbering / page-number formatting.**
  `<w:titlePg/>` is the only section-level toggle exposed today.
- **Draggable ruler.** The P3.5 `PageRuler` is read-only.
- **Auto-creation of header/footer parts.** Toggling "Different first
  page" flips the typed flag but does not synthesize the header part.
- **Editing inside SDT/TOC wrappers.** P2.3 unwraps for display only;
  edits through a carrier are intentionally blocked until the
  `subtreeDirty` plumbing is exercised by a typed command.
- **Image resize handles, alt-text editing, float/wrap.** Model carries
  the data; UX is missing.
- **Numbering auto-mint.** Bullet / numbered list buttons in P2.2 fail
  loudly via toast when no `<w:abstractNum>` exists rather than minting
  one.
- **Live LLM provider in production.** `apps/web` can call OpenAI when
  `OPENAI_API_KEY` is set; without it, P2.5's honest offline fallback
  attaches the prompt as a comment with a "no LLM was called"
  rationale.

## Reading order for a new contributor

1. [`prompt.md`](../prompt.md) — the brief.
2. [`spec/shared/`](../spec/shared) — what a "document" is in our system.
3. [`spec/docx/`](../spec/docx) — the DOCX contract. Start with
   `ooxml-mapping.md`, then `document-model.md`, then the workstream
   specs (`style-cascade.md`, `page-model.md`, `paged-renderer.md`,
   `header-footer-authoring.md`, `page-aware-editing.md`,
   `llm-page-surface.md`).
4. [`docs/build-log/docx.md`](build-log/docx.md) — what actually
   shipped, in order, with caveats.
5. [`spec/docx/eigenpal-synthesis.md`](../spec/docx/eigenpal-synthesis.md)
   — research note used as input for the P3 roadmap.
6. [`docs/roadmap-docx-p1.md`](roadmap-docx-p1.md) — historical (P1
   roadmap; superseded by what shipped, kept for context).

## How to run

```bash
pnpm install
pnpm typecheck
pnpm test                         # all 1014 tests across docx + xlsx
pnpm fixtures:docx                # regenerate synthetic DOCX fixtures
pnpm --filter @officeai/web dev   # web editor at http://localhost:3000

# CLI
pnpm --filter @officeai/agent build
node packages/agent/dist/cli.js docx read --file fixtures/docx/synthetic/01-plain-paragraphs.docx --format markdown
node packages/agent/dist/cli.js mcp     # stdio MCP server (docx_* + xlsx_* tools)
```

---

## Appendix: original P0 session notes (2026-04-17)

The notes below are kept verbatim from the first session for historical
context. Subsequent phases (P1, P2, P3) live in the build log; the
"deferred" and "harder than expected" items here have largely been
addressed or evolved — cross-reference the build log for the current
state.

### P0 scope

Full DOCX phase only (Analyze → Spec → Build → Validate). XLSX and PPTX
deferred at the time. The P0 deliverables were:

- **`@officeai/core`** — format-agnostic OOXML I/O (`OoxmlContainer`,
  XML parse/serialize, relationships, content types) and a typed
  `CommandBus` with approved/pending/working state, mutation diffs,
  and rollback.
- **`@officeai/docx`** — parser (with opaque-blob preservation),
  in-memory model, serializer (byte-preserving for untouched parts),
  six P0 command handlers, a headless `DocxAgent`, and a ProseMirror
  renderer that funnels every edit through the bus.
- **`@officeai/agent`** — an `office-agent` CLI: `read | search |
  insert-text | comment | apply` that uses the same headless agent.
- **`apps/web`** — a Notion-flavored DOCX editor surface.
- **`tests/`** — integration suite that round-trips five synthetic
  fixtures and verifies untouched parts stay byte-identical.

P0 ended at **64 tests total**. P3 ends at **397 docx-relevant tests**.

### Things that were harder than expected (P0)

1. **`fast-xml-parser` + `preserveOrder` + namespaces.** Required
   pinning `unpairedTags` (`<w:br/>`, `<w:tab/>`, `<w:cr/>`) and
   explicitly typing the option arrays as `string[]` (the `as const`
   produces `readonly string[]` which the type defs reject). Took two
   rebuilds before `parseXml(serializeXml(parseXml(x)))` was stable.
2. **Agent state plumbing in React.** The first cut of the editor read
   `agent.getSnapshot()` from a ref during render; Next.js 15 + React
   19's `react-hooks/refs` rule (correctly) errored. Refactored to
   push agent state into `useState` via the `subscribe` callback.
3. **Isomorphic SHA-256.** `node:crypto` works fine in tests and the
   CLI, but `apps/web` ships through webpack which does not handle
   the `node:` URI scheme. Switched the entire `sha256Hex` helper to
   `js-sha256` (MIT) so the same `OoxmlContainer` runs in browser
   bundles.
4. **OOXML "untouched" really means _bytewise_ untouched.** Keeping
   the raw `Uint8Array` for unchanged parts and only re-emitting
   dirty parts (driven by `DocxDirtyFlags`) made the byte-equality
   test trivial — and it's a load-bearing invariant: it's what gives
   reviewers confidence that the agent didn't silently corrupt
   anything.
5. **The `docx:add-comment` ripple.** Adding the first comment to a
   document that has no comments part means: minting `comments.xml`,
   registering its `Override` in `[Content_Types].xml`, adding a
   `comments` relationship in `word/_rels/document.xml.rels`, and
   inserting three inline markers (`commentRangeStart`,
   `commentReference`, `commentRangeEnd`) into the paragraph in the
   right order. All four had to land in one mutation or roundtrip
   would split.
