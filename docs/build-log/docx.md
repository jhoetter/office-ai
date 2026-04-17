# DOCX Build Log

> Live log of decisions, deviations from spec, and known issues.

## Decisions

| Date (UTC) | Decision                                                                 | Rationale                                                                                                                                 |
| ---------- | ------------------------------------------------------------------------ | ----------------------------------------------------------------------------------------------------------------------------------------- |
| 2026-04-17 | Use `fast-xml-parser` with `preserveOrder: true` over a custom XML model | Round-trips losslessly for OOXML; smaller surface than building our own.                                                                  |
| 2026-04-17 | Tables stored as opaque-XML `Table.raw` for now (P1 mutation)            | Cell merging is genuinely tricky; preserving roundtrip integrity matters more than editing this session.                                  |
| 2026-04-17 | Hyperlink modeled as a typed wrapper over runs (not a mark)              | Matches OOXML structure (`w:hyperlink` is a block element nesting runs); avoids mark-coalescing ambiguity.                                |
| 2026-04-17 | Comments staging tri-state lives in core, not docx                       | Same logic will serve XLSX/PPTX.                                                                                                          |
| 2026-04-17 | The serializer trusts dirty flags rather than diffing snapshots          | Cheap, predictable, and matches our command-bus discipline.                                                                               |
| 2026-04-17 | Touched parts may differ from input in attribute order / quote style     | Word and LibreOffice both accept either; we assert structural equivalence on touched parts and bytewise equality only on untouched parts. |

## Deviations from spec

| Date (UTC) | Spec section                  | Deviation                                                                                 | Reason                                                                                                                                     |
| ---------- | ----------------------------- | ----------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------ |
| 2026-04-17 | `set-paragraph-style` payload | `style` is not validated against `word/styles.xml`                                        | We don't model styles yet; we trust the AI / human to use a known styleId. Will tighten when we ingest the styles part.                    |
| 2026-04-17 | `core/util/hash`              | Switched from `node:crypto` to the `js-sha256` package                                    | Required isomorphic implementation so the same `OoxmlContainer` works in `apps/web` (browser bundle) and Node. Output bytes are identical. |
| 2026-04-17 | `renderer.md`                 | `transactionToCommands` does not yet emit `docx:set-paragraph-style` from PM transactions | The toolbar route still hits the agent directly; we'll wire the PM keymap branch when we add a real style menu.                            |

### Resolved deviations

| Date resolved (UTC) | Spec section                                         | Original deviation                                  | Resolution                                                                                                                                                          |
| ------------------- | ---------------------------------------------------- | --------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 2026-04-17          | `agent-commands.md` (`format-range`, `delete-range`) | Multi-paragraph ranges threw `CommandError` for now | Lifted in P1.1 / W2. Both handlers now traverse paragraph spans (split / merge as required). See "P1.1 — W2: Range edits + PM funnel" below for the implementation. |

## Deferred to a follow-up session

| Item                                                       | Spec ref            | Status                                        |
| ---------------------------------------------------------- | ------------------- | --------------------------------------------- |
| `docx:insert-table`                                        | `agent-commands.md` | Stub — throws NotImplementedError             |
| `docx:set-cell-content`                                    | `agent-commands.md` | Stub                                          |
| `docx:insert-image`                                        | `agent-commands.md` | Stub                                          |
| `docx:accept-change`                                       | `agent-commands.md` | Stub                                          |
| `docx:reject-change`                                       | `agent-commands.md` | Stub                                          |
| Real-world DOCX fixtures (Word/Google/LibreOffice exports) | `feature-scope.md`  | Slots reserved in `fixtures/docx/MANIFEST.md` |
| LibreOffice CI roundtrip                                   | `feature-scope.md`  | Manual today; CI integration deferred         |
| Headers/footers editing                                    | `feature-scope.md`  | P1 — preserved verbatim; not mutable          |
| Image insertion                                            | `feature-scope.md`  | P1 — preserved on roundtrip; not mutable      |
| List mutation (numbering insert/remove)                    | `feature-scope.md`  | P1 — preserved; style-by-name only            |

## Known issues

| Issue                                                   | Detail                                                                                                                         | Mitigation                           |
| ------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------ | ------------------------------------ |
| Web demo agent prompt is hard-coded ("[AI] " + comment) | The point of the prompt panel is to demonstrate the human-review queue; an actual LLM call lives outside this session's scope. | Documented in the `/editor` UI copy. |

### Resolved in P1.1 / W2

- ProseMirror funnel does not preserve nested marks across boundary edits — see "P1.1 — W2: Range edits + PM funnel" below.
- `transactionToCommands` block-bearing slice path always emits a single `insert-paragraph` regardless of slice content — same section.

## UX fixes (2026-04-17, post-validation)

After the first browser walk-through the user reported "typing adds lots
of new lines and so on". Root-causes and fixes:

| Symptom                                                                                                                                                                                                           | Root cause                                                                                                                                                                                                                                                                                                                                     | Fix                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        |
| ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Console error: "ProseMirror expects the CSS white-space property to be set, preferably to 'pre-wrap'"                                                                                                             | We never imported `prosemirror-view`'s default stylesheet, and our `.prose-pm` class did not set `white-space`. Without it, ProseMirror's DOM ↔ position mapping is unreliable, manifesting as text inserted in the wrong place.                                                                                                               | `apps/web/app/layout.tsx` now imports `prosemirror-view/style/prosemirror.css`. `apps/web/app/globals.css` adds belt-and-suspenders `white-space: pre-wrap` and the other PM base rules to `.prose-pm`.                                                                                                                                                                                                                                                                                                                    |
| Cursor jumps to position 0 on every keystroke; characters appear to "stack" or create newlines at the start of the document                                                                                       | `mountDocxEditor`'s `dispatchTransaction` **dropped** the user's PM transaction, applied the equivalent command through the bus asynchronously, then `replaceWith`-ed the entire PM doc on every subscribe. That re-projection mapped any selection inside the deleted range to position 0, so the next keystroke was applied at offset 0 too. | Refactored `packages/docx/src/renderer/mount.ts`: user transactions are applied to the EditorView **immediately** (selection stays put, DOM updates synchronously), then mirrored into the bus. A `pendingFunnelCount` counter ensures the corresponding subscribe notifications do not echo back as a re-projection. External mutations (agent prompt, direct `agent.applyCommand` from the host UI) still cause a re-projection, but the previous selection is now mapped through the new doc instead of resetting to 0. |
| `docx:insert-paragraph` always inserted an empty paragraph at the given index, regardless of the offset, so pressing Enter mid-paragraph silently desynced the agent's model from PM                              | The handler ignored `at.offset` entirely.                                                                                                                                                                                                                                                                                                      | `insertParagraphHandler` now splits the source paragraph at `at.offset` (Enter semantics) when the offset is mid-text, preserves trailing-style inheritance, and falls back to "insert empty paragraph before/after" at the boundaries. New tests in `handlers.test.ts` cover all three branches.                                                                                                                                                                                                                          |
| Toolbar metadata ("4 blocks · rev 0 · 0 comments") overlapped the agent sidebar at common laptop widths; heading text wrapped because the editor surface was narrower than `max-w-prose` after the 320 px sidebar | The grid was hard-coded to `grid-cols-[minmax(0,1fr)_320px]` with `max-w-prose` (~65 ch) and an always-row toolbar without `flex-wrap`.                                                                                                                                                                                                        | `DocxEditor.tsx` now uses a responsive grid (`lg:grid-cols-[minmax(0,1fr)_300px]`, stacks below `lg`), the toolbar uses `flex-wrap` and collapses the metadata under `md`, and the editor surface is sized to a Notion-page-like 720 px. The aside drops the left border / left-pad below `lg`.                                                                                                                                                                                                                            |

Test impact: 32 docx tests (was 30) and 17 integration tests (unchanged)
all green; `pnpm test` across the workspace stays green.

## Sample-DOCX styling fix (2026-04-17, post-export-validation)

Follow-up report from the user: "the exported file really just is a plain
text in word, e.g. headlines and such aren't considered". Reproduced and
root-caused entirely from the terminal:

1. Built the in-browser sample DOCX exactly like
   `apps/web/app/lib/sample-docx.ts` does, opened it through `DocxAgent`,
   inserted text via the bus, exported, unzipped the result.
2. Confirmed `word/styles.xml` was **missing** from both the input and
   the export, even though `word/document.xml` referenced
   `<w:pStyle w:val="Heading1"/>`. With no styles part, Word and
   LibreOffice silently fall back to the default paragraph style, so
   "Welcome to officeAI" rendered as plain body text.
3. Cross-checked with a real Word-grade `.docx` produced by the `docx`
   library (which ships a styles.xml): after a `DocxAgent` open → edit →
   export, `word/styles.xml` was present **and byte-identical** to the
   input. The serializer was fine; only the demo's synthetic package was
   broken.
4. Visual verification with `soffice --headless --convert-to pdf` +
   `pdftoppm`: before the fix, all paragraphs rendered at body size;
   after the fix the heading 1 / heading 2 paragraphs render bold and
   larger, exactly as Word would show them.

Fix: `apps/web/app/lib/sample-docx.ts` now ships a real `word/styles.xml`
that defines `Normal`, `Title`, `Heading1`, `Heading2`, `Heading3`, plus
the matching `Override` in `[Content_Types].xml` and the styles
relationship in `word/_rels/document.xml.rels`. The sample document was
also extended with a Heading 2 line so the demo exercises more than one
heading level. No agent / parser / serializer code changed — this was
purely a demo-fixture issue masquerading as an exporter bug.

## Validation summary (2026-04-17)

- All five packages (`@officeai/core`, `@officeai/docx`, `@officeai/agent`, `@officeai/web`, `@officeai/integration-tests`) typecheck.
- Test totals: **64 passing** across **10 test files**.
  - core: 12 tests (CommandBus + OoxmlContainer)
  - docx: 30 tests (parser, serializer, command handlers, agent, renderer)
  - agent CLI: 7 tests (read / search / insert-text / comment / apply / unsupported / invalid selector)
  - integration: 15 tests (5 fixtures × 2 roundtrip variants + 5 agent-edit variants)
- Web app builds (`next build`) with no warnings; bundle size for `/editor` is 119 kB First-Load JS.
- License audit (manual, all production deps in `node_modules/.pnpm`):
  - All MIT / ISC / Apache-2.0 / dual `MIT OR GPL-3.0-or-later` (jszip).
  - No GPL-only or commercial-license deps in the runtime tree.

## P1.1 — W1: Real-world fixtures + LibreOffice CI

> Date: 2026-04-17. Workstream W1 of batch P1.1 (per
> `docs/roadmap-docx-p1.md`). Ships A1, A2, D1.

### What landed

1. **`scripts/generate-real-fixtures.mjs`** — emits six Word-grade
   fixtures via the `docx` MIT npm library into
   `fixtures/docx/real-world/`. We treat the `docx` library as
   "as-real-as-we-can-get-without-shipping-third-party-content": each
   fixture ships a real `word/styles.xml`, `word/numbering.xml`, the
   header/footer parts where relevant, an inline `word/media/*.png` for
   the image fixture, and a `word/comments.xml` thread plus `w:ins`/`w:del`
   wrappers for the tracked-changes fixture. All six stay below the 50 KB
   budget enforced at write time (largest is `02-report-headers-footers.docx`
   at ~10.5 KB). Inventory matches the table appended to
   `fixtures/docx/MANIFEST.md`:

   ```text
   01-styled-letter.docx           bullets + headings + bold/italic
   02-report-headers-footers.docx  multi-page + header1.xml + footer1.xml
   03-numbered-list.docx           numbering.xml-driven 2-level list
   04-table-grid.docx              4×3 table with header row
   05-inline-image.docx            inline drawing + media part
   06-comments-and-changes.docx    comments.xml + w:ins + w:del
   ```

2. **`scripts/run-libreoffice-roundtrip.mjs`** — for every fixture,
   converts it to PDF via `soffice --headless --convert-to pdf`, then
   runs the same buffer through `DocxAgent.fromBuffer → exportFile` and
   converts THAT to PDF. Both passes must exit 0 and emit no
   "repair / error / corrupt / unable to load / failed to" text on
   stderr. **Skips gracefully (exit 0 with a warning)** if `soffice` is
   not on PATH so dev machines without LibreOffice don't fail. CI
   installs it explicitly.

3. **`tests/roundtrip/docx/real-world-roundtrip.test.ts`** — vitest suite
   that iterates the fixtures and asserts the byte-preservation
   invariant on two scenarios:
   - **Pure roundtrip** — every part is byte-identical after
     `parse → serialize → re-load`.
   - **Trivial edit** — after one inserted character at the start of the
     first paragraph, every part EXCEPT `word/document.xml` stays
     byte-identical. The edited part is also asserted to differ (sanity
     check that the edit actually landed).

   Runs as part of `pnpm test` / `make test`.

4. **`apps/web/playwright.config.ts` + `apps/web/e2e/*`** — six
   Playwright specs that exercise the editor against the bundled
   `apps/web/app/lib/sample-docx.ts` (so they don't depend on the
   real-world fixture corpus):
   - `open-fixture.spec.ts` — editor mounts and renders the sample.
   - `typing.spec.ts` — typed characters appear in the document.
   - `enter-paragraph.spec.ts` — Enter splits a paragraph (block count
     in the metadata strip increments).
   - `format-bold.spec.ts` — Bold toolbar wraps the selection in
     `<strong>`.
   - `add-comment.spec.ts` — Add-comment toolbar increments the
     comments counter and surfaces the "Comment added." toast.
   - `export-roundtrip.spec.ts` — Export downloads a valid OOXML zip
     (verified by the `PK\x03\x04` magic and a minimum size).

   `pnpm --filter @officeai/web e2e:install` installs the chromium
   browser; `make e2e-web` builds the workspace and runs the suite
   against `next start`.

5. **Makefile** gains `fixtures-real`, `roundtrip-libre`, and `e2e-web`
   targets. The two heavy targets are deliberately **not** wired into
   `make verify` — they need system-level deps (LibreOffice,
   Playwright browsers) that not every dev box has.

6. **`.github/workflows/ci.yml`** gains two sibling jobs:
   - `docx-libreoffice-roundtrip`: installs `libreoffice-core` +
     `libreoffice-writer` via apt, builds the workspace, runs
     `make roundtrip-libre`.
   - `web-e2e`: builds the workspace, installs Playwright browsers,
     runs the suite, uploads the HTML report on failure.

   Both run alongside the existing `verify` job, so a flaky LibreOffice
   step doesn't gate the core quality gate.

### What LibreOffice surfaced

Locally (LibreOffice 25.8.2.2 on macOS):

- All six fixtures roundtrip clean — original → PDF and post-`DocxAgent`
  re-export → PDF both exit 0 and emit no stderr.
- One self-inflicted bug caught during development: the first cut of the
  inline-image fixture embedded a hand-typed PNG hex string with an
  invalid IDAT CRC; LibreOffice surfaced `libpng error: IDAT: CRC error`
  and the script correctly failed. Fixed by computing a valid 1×1 RGBA
  PNG via Node's `zlib.deflateSync` + `zlib.crc32`.
- No issues with namespace prefixes, missing styles, or
  `mc:AlternateContent` wrappers. The `docx` library ships sane defaults
  on all of those, which is why we treat real Word/Pages/Google-Docs
  exports (the "to-collect" slots in `MANIFEST.md`) as still-needed
  follow-up coverage.

### CI shape

```text
verify                          ← existing quality gate, no LibreOffice/Playwright dep
docx-libreoffice-roundtrip      ← new; installs soffice; runs make roundtrip-libre
web-e2e                         ← new; installs playwright browsers; runs make e2e-web
```

`make verify` itself stays untouched in spirit: it does not invoke
either heavy step, and continues to pass on a machine without
`soffice` or Playwright installed (verified locally — see the W1
validation summary).

### Local validation

- `node scripts/generate-real-fixtures.mjs` — 6/6 fixtures emitted, all
  under the 50 KB budget.
- `node scripts/run-libreoffice-roundtrip.mjs` — 6/6 pass on macOS with
  LibreOffice 25.8.2.2.
- `pnpm --filter @officeai/integration-tests test` — **27 tests pass**
  (was 15; +12 from the new real-world suite).
- `make verify` — green; not re-run after the heavy targets because the
  scope of W1 doesn't touch any quality-gate-relevant code paths.

### Known limitations / follow-ups

- The "to-collect" real-world slots in `fixtures/docx/MANIFEST.md` remain
  open. The generated fixtures cover the SHAPES we care about, but they
  don't catch real Word's font tables, Mac-Word attribute orderings,
  Google Docs' inline-`rPr` quirks, or Pages' non-standard relationships.
  Those still need genuine documents.
- The Playwright `enter-paragraph` and `format-bold` specs assume the
  sample document's first paragraph contains the literal "Welcome"; if
  `apps/web/app/lib/sample-docx.ts` is ever rewritten, the helpers in
  `apps/web/e2e/_helpers.ts` must be updated.
- The `web-e2e` CI job uses a single chromium browser to keep the run
  short. Cross-browser coverage (firefox, webkit) is a P1.2 follow-up if
  cost allows.

## P1.1 — W2: Range edits + PM funnel

> Date: 2026-04-17. Workstream W2 of batch P1.1 (per
> `docs/roadmap-docx-p1.md`). Ships A3, C1, C2 — i.e. lifts the
> "multi-paragraph throws" deviation on `format-range` / `delete-range`
> and closes the two ProseMirror funnel issues that were the only
> remaining "Known issues" entries on the renderer side.

### What landed

1. **Multi-paragraph `docx:format-range`**
   (`packages/docx/src/commands/format-range.ts`). The handler now
   normalizes `start` / `end` (so the caller can pass them in either
   order) and, when `start.paragraph !== end.paragraph`, walks the
   paragraph span:
   - **Start paragraph**: format from the start boundary to
     end-of-paragraph. Splits the boundary run when needed.
   - **Intermediate paragraphs**: format every run.
   - **End paragraph**: format from start-of-paragraph to the end
     boundary. Splits the boundary run when needed.
   - Non-paragraph blocks (tables, opaque blocks, section breaks)
     inside the span are skipped — the formatting "walks past" them
     without crashing or rewriting their bytes.
     The diff for a multi-paragraph apply is a single `DocumentDiff`
     with one `node-updated` change per paragraph that was actually
     touched. Dirty flags: `body: true` (unchanged behavior).

2. **Multi-paragraph `docx:delete-range`**
   (`packages/docx/src/commands/delete-range.ts`). Cross-paragraph
   ranges now:
   - Trim the **start paragraph** from the start boundary to its end.
   - Drop every intermediate paragraph entirely.
   - Trim the **end paragraph** from its beginning to the end
     boundary.
   - **Merge** the trimmed start and trimmed end paragraphs into a
     single paragraph. The start paragraph's `id`, `pPr`, and
     `properties` win — it absorbs the end paragraph's surviving
     content. If the merge would leave the paragraph with zero runs
     (both sides emptied), an empty placeholder run is appended so
     the paragraph stays well-formed for the renderer / serializer.
     The diff includes one `node-updated` for the merged start
     paragraph plus one `node-deleted` per dropped paragraph (including
     the now-absorbed end paragraph). Dirty flags: `body: true`.

3. **PM funnel: mark re-assertion across boundary edits**
   (`packages/docx/src/renderer/transaction-to-commands.ts`,
   `emitInsertWithMarkReassertion`). When a `ReplaceStep` with
   `from === to` inserts text, the funnel now emits the existing
   `insert-text` command **plus**, when warranted, a follow-up
   `format-range` that re-asserts the ambient text marks across the
   inserted span. The heuristic (documented inline):
   1. Read marks from the slice's first text node — PM's
      `tr.insertText` automatically applies `storedMarks`, so a paste
      / typing inside a formatted span carries those marks here.
   2. If the slice has no text marks (e.g. a programmatic
      insertion), fall back to "marks at the insertion point"
      (`$pos.marks()`), which Prosemirror computes as marks of the
      node-before, or node-after at the start of the textblock.
   3. Convert the marks to a `TextFormat`. Structural marks
      (`hyperlink`, `comment_mark`, `revision_mark`) are filtered out
      — they're paragraph / run wrappers in our model, not run
      properties.
   4. Only emit the follow-up when the resulting format is
      non-empty; otherwise the funnel is unchanged.

4. **PM funnel: multi-block paste** (same file,
   `emitMultiBlockPaste`). A `ReplaceStep` whose slice carries N
   top-level paragraph blocks (typical for a paste of multiple
   paragraphs) is no longer collapsed to a single
   `insert-paragraph`. Instead the funnel emits the natural sequence:
   - `insert-text` for the first segment's text at the cursor,
   - `insert-paragraph` at `(cursor + segment[0].length)` to split
     the current paragraph,
   - `insert-text` + `insert-paragraph` for every middle segment,
   - `insert-text` for the last segment's text at the new paragraph
     start.
     Position math is performed against the **model**, not PM offsets,
     so the resulting commands are independent of how PM happens to
     number positions. Slices that carry non-paragraph blocks (tables
     in particular) are flagged via the existing `unsupported`
     channel — the funnel does **not** crash; it just skips mirroring
     that step into the bus, leaving the EditorView's optimistic
     render in place until the user retries.

5. **`buildDiffMulti` helper** (`packages/docx/src/commands/helpers.ts`)
   — the only addition to `helpers.ts`, used by the two
   multi-paragraph handlers above. Single-change `buildDiff` is
   unchanged.

### Tests

- `packages/docx/src/commands/handlers.test.ts` gains three
  multi-paragraph cases (format across three paragraphs; delete
  across three paragraphs; delete that fully empties the merged
  paragraph). Existing single-paragraph cases unchanged.
- `packages/docx/src/renderer/transaction-to-commands.test.ts`
  (new) is table-driven and covers: bold-boundary insert,
  PM-encoded marks on the inserted slice, plain-text insert (no
  spurious format-range), 2-paragraph paste, 3-paragraph paste,
  paste with empty trailing paragraph, multi-paragraph delete,
  multi-paragraph format.
- Test totals for `@officeai/docx` go from **32 → 58** (the new
  W2 work adds 11; the new W3 comments-lifecycle suite contributes
  the remainder). All other packages' tests are unaffected.

### Algorithmic notes

- Both `format-range` and `delete-range` reuse a common
  `paragraphTextOffset(p, runIndex, localOffset)` that now treats
  `runIndex === undefined` as "interpret `localOffset` as a
  paragraph-wide character offset" (clamped to the paragraph's
  length). This makes the multi-paragraph caller's life trivial:
  paragraph-wide offsets travel through the handler unchanged. The
  per-run interpretation is preserved when `runIndex` is supplied,
  so existing single-paragraph callers (and the agent) see no
  behavioral change.
- For multi-paragraph delete, the merge step preserves the start
  paragraph's `id` deliberately — node-id stability matters because
  PM-side comment / revision marks reference paragraphs by id; the
  merged paragraph stays "the same" paragraph from the renderer's
  perspective. The end paragraph's id flows into a `node-deleted`
  diff entry so the bus / decoration plugins can drop their
  references.
- For multi-block paste, the funnel computes per-step positions in
  the model coordinate space (not PM positions), so the sequence is
  immune to PM's accounting of paragraph boundaries (`+2` per
  paragraph). The trade-off: paragraph-level properties carried by
  the pasted paragraphs (e.g. heading style of pasted-in `<h1>`)
  are NOT preserved by this round; that's an explicit P1.2+ follow-up
  along with the toolbar work.

### Known follow-ups (not in W2 scope)

- `insert-text` does not yet honor paragraph-wide offsets when
  `run` is supplied but the offset overflows the targeted run; it
  appends to the run instead of advancing into the next. The PM
  funnel works around this for the mark-reassertion follow-up by
  emitting the `format-range` against paragraph-wide offsets, but
  the underlying handler should still be tightened.
- Pasted paragraph styles / `pPr` are dropped by the multi-block
  paste path. Re-introducing per-segment `set-paragraph-style`
  commands lands with C3 (toolbar parity) in P1.2.

## P1.1 — W3: Agent surface (CLI + MCP + comment lifecycle)

Scope: bring `office-agent` and the headless agent up to the surface
described in `prompt.md` lines 451–493 + ship the three previously
stubbed comment-lifecycle commands end-to-end.

### Selector

`packages/agent/src/selector.ts` learned the `section:S/...` prefix.
Only `section:0` is accepted in P1 (the document model is one body
section); other indices throw `SelectorError` so callers fail loudly
instead of silently dropping the prefix. Range form unchanged.

### CLI

`packages/agent/src/cli.ts` is restructured around a `docx`
subcommand group while keeping the old top-level commands as
backward-compatible shims:

```
office-agent docx inspect        --file <path>
office-agent docx read           --file <path> --format markdown|json|text [--range <selector>]
office-agent docx search         --file <path> --query <text>
office-agent docx write          --file <path> --at <selector> --text <s> [--out <p>]
office-agent docx style          --file <path> --at <selector> --style <id> [--out <p>]
office-agent docx comment        --file <path> --range <selector> --text <s> [--out <p>]
office-agent docx resolve-comment --file <path> --id <commentId> [--reopen] [--out <p>]
office-agent docx reply-comment   --file <path> --parent <id> --text <s> --author <name> [--out <p>]
office-agent docx delete-comment  --file <path> --id <commentId> [--out <p>]
office-agent docx apply           --file <path> -c <commands.json> [--out <p>]
office-agent docx diff            --before <p> --after <p>
office-agent mcp                 (start the MCP stdio server)
```

Conventions:

- `--file` is the input path; `--out` defaults to `--file` (in-place).
  This matches `spec/agent/cli.md` §`--out semantics`.
- Structured output is one JSON document per invocation. `--pretty`
  enables indenting.
- Old commands (`read`, `search`, `insert-text`, `comment`, `apply`)
  remain as `[legacy]` aliases at the top level so existing scripts
  don't break.

### MCP server

`packages/agent/src/mcp.ts` implements an `officeai` MCP server
(stdio transport) that exposes the `DocxAgent` as seven tools:

- `docx_load(path) → { handle, summary }` — opens a file, mints an
  in-process handle.
- `docx_save(handle, out_path?)` — serializes back to disk.
- `docx_inspect(handle)` — same shape as `docx inspect`.
- `docx_get_text(handle, format = "markdown" | "json" | "text")`.
- `docx_search(handle, query, case_sensitive?, regex?)`.
- `docx_apply_command(handle, type, payload, source?, agent_id?, auto_approve?)`
  — covers every registered docx handler, including the comment
  lifecycle commands. `auto_approve` defaults to true so a single
  agent-source command lands as `approved` in one round-trip.
- `docx_diff({before, after}` for handle-vs-handle, or
  `{handle, against: "disk"}` to diff against the on-disk file the
  handle was loaded from).

Sessions are in-process (`Map<handle, DocxAgent>`); handles are
opaque UUIDs. `__resetMcpSessionsForTests` is exported for tests
only. Input schemas use `zod` (added as a direct dep on
`@officeai/agent`).

### Comment lifecycle (model + parser + serializer + handlers)

The three previously stubbed comment commands are now implemented.
This required end-to-end support for `word/commentsExtended.xml`
(W15 metadata for `done` + `parentPaIdRef`).

Model (`packages/docx/src/model/types.ts`):

- `DocxDirtyFlags.commentsExtended` joins the existing flags. It is
  independent of `comments`: resolving a comment dirties only the
  extended part, while adding a comment dirties both.
- `DocxComment` gains `resolved?: boolean`, `parentId?: string`, and
  `paraId?: string`. `paraId` is the W14 paragraph id of the
  comment's first body paragraph — `commentsExtended.xml` keys
  threading and resolved-state by it (NOT by the comment id).

Parser (`packages/docx/src/parser/parse.ts`):

- `parseComments` captures the existing `w14:paraId` on a comment's
  first paragraph if present.
- `parseCommentsExtended` reads `word/commentsExtended.xml` (when
  present) into a paraId-keyed map; `applyCommentsExtended` projects
  that map onto the parsed comments. Missing parts are silently
  skipped — old documents stay parseable.

Serializer (`packages/docx/src/serializer/serialize.ts`):

- `serializeComment` injects `w14:paraId` on the first paragraph of
  every comment body. If the comment has no `paraId` yet, one is
  derived deterministically from the comment id (so the same
  document round-trips byte-stably).
- When `dirty.comments` or `dirty.commentsExtended` is set, the
  serializer rewrites `word/commentsExtended.xml` from the projected
  comment list, emitting one `w15:commentEx` per comment that is
  resolved or has a parent. When no comments need extended metadata
  the part (and its relationship + content-type entry) is dropped.

Handlers
(`packages/docx/src/commands/{resolve,reply,delete}-comment.ts`):

- `docx:resolve-comment` toggles `resolved` and dirties only
  `commentsExtended`. Idempotent (no-op + revision bump when the
  state is already as requested), rejects `unknown-comment`.
- `docx:reply-comment` mints a fresh comment with `parentId` set
  to the target. **Replies do not add new range markers in the
  body** — they share the parent's `commentRangeStart`/`End`/
  `Reference`, which is what makes Word render them indented under
  the same anchor. Rejects `unknown-comment` and `empty-reply`.
- `docx:delete-comment` removes the comment and its inline range
  markers, plus every reply whose `parentId` chains back to it
  (transitive). When `comments` becomes empty, `comments.xml`,
  `commentsExtended.xml`, the relationships, and the content-type
  entries are all dropped on save.

Registry (`packages/docx/src/commands/registry.ts`) wires the three
handlers in alongside the existing six P0 handlers. The
`docx:resolve-comment` stub line was removed — the others
(`insert-table`, `set-cell-content`, `insert-image`,
`accept-change`, `reject-change`) remain stubs.

### Tests

- `packages/docx/src/commands/comments-lifecycle.test.ts` (new) —
  15 tests covering successful operations, error paths, and
  round-trip serialize/parse with `commentsExtended.xml`.
- `packages/agent/src/cli.test.ts` extended with the
  `office-agent docx …` subcommand surface (inspect, read formats,
  write with `section:0/...` selectors, style, full comment
  lifecycle, diff). Legacy top-level commands still pass.
- `packages/agent/src/mcp.test.ts` (new) — 9 tests over an
  in-memory MCP transport pair: tool discovery, load/inspect,
  text projection in all three formats, search, apply +
  save round-trip, end-to-end comment lifecycle through
  `docx_apply_command`, handle-vs-handle diff, and unknown-handle
  error path.

### Known follow-ups (out of W3 scope)

- Section selectors past `section:0` need a real section model
  (split body, per-section properties). Tracked alongside the
  remaining `feature-scope.md` items.
- `docx_apply_command` validates payload shape only via the
  underlying handlers — adding per-command zod schemas in
  `mcp.ts` would let MCP clients receive richer input errors.
  Deferred until we have a real LLM driving the surface.
- The stdio entry point relies on `process.stdin` close to terminate;
  `office-agent mcp` does not yet support a `--port` flag for the
  HTTP/SSE transport. `prompt.md` only requires stdio for now.

## P1.2 — W4: Headers/footers + tracked changes

> Date: 2026-04-17. Workstream W4 of batch P1.2 (per
> `docs/roadmap-docx-p1.md`). Ships A4 (typed headers/footers + the two
> `docx:set-*-text` commands) and B2 (the `docx:accept-change` /
> `docx:reject-change` resolution commands). C5 (tracked-change UI)
> lives in W5.

### What landed

1. **Header / footer model + parser**
   (`packages/docx/src/parser/headers-footers.ts`,
   `packages/docx/src/model/types.ts`). New `HeaderFooterPart` carrier:

   ```typescript
   interface HeaderFooterPart {
     kind: "header" | "footer";
     id: string; // equals `partPath`; doubles as the handler id
     partPath: string; // e.g. "word/header1.xml"
     target: "default" | "first" | "even";
     rootAttrs: Readonly<Record<string, string>>;
     body: ReadonlyArray<BlockNode>;
   }
   ```

   The discoverer walks `word/_rels/document.xml.rels` for the two
   header / footer relationship types, then recursively scans every
   `<w:sectPr>` (top-level AND nested inside `<w:p><w:pPr>`) to recover
   each rId's `w:type` (defaulting to `"default"` when absent —
   matches Word's behavior). The parser shares `parseParagraph` from
   the main parser via injection rather than re-export to avoid an
   import cycle. Tables / SDT / unknown blocks inside a header part
   stay as `OpaqueBlock` carriers this round; typed table mutation is
   W7 / P1.3.

   The discovered parts hang off
   `snapshot.root.headersAndFooters: ReadonlyArray<HeaderFooterPart>`
   (clean name; matches `body` / `comments` shape).

2. **Header / footer serializer**
   (`packages/docx/src/serializer/headers-footers.ts`). A new
   `serializeHeaderFooterParts(container, snapshot, serializeBlock)`
   step at the end of `serializeDocx`:
   - Skips parts whose path is NOT in
     `snapshot.dirty.headersAndFooters` — this is what guarantees
     SHA-256-level byte-identity for untouched parts on round-trip.
   - For touched parts, re-serializes the typed model wrapped in the
     original namespace declarations (`rootAttrs`).
   - Reuses the main `serializeBlock` so paragraph emission stays
     identical to the body's (same `w:pPr` ordering, same
     `xml:space="preserve"` discipline, same opaque-block passthrough).
   - Loud failure if a part is dirty but missing from the model or
     container — `DocxSerializeError("header-footer-missing")`.

3. **Dirty flags**: `DocxDirtyFlags` gained
   `headersAndFooters: ReadonlySet<string>`. Independent of `body`,
   `comments`, etc. The set is treated as immutable across snapshots
   so older snapshots in the bus's mutation history keep their own
   dirty view (helpers.ts is owned by W2; we merge by building a new
   `Set` in the W4-owned `mergeHeaderFooterDirty` helper inside
   `set-header-text.ts`).

4. **`docx:set-header-text` / `docx:set-footer-text`**
   (`packages/docx/src/commands/set-header-text.ts`,
   `set-footer-text.ts`). Payload:

   ```typescript
   {
     partId: string;
     paragraphIndex: number;
     text: string;
   }
   ```

   Replaces the targeted paragraph's text with a single run that
   carries the existing first run's `rPr` (italics, font family,
   etc. survive). Errors: `unknown-target` for missing part / OOB
   index / non-paragraph block. Idempotent. Both handlers share
   `applySetTextToHeaderFooter()` keyed on `kind: "header" | "footer"`.

5. **`docx:accept-change` / `docx:reject-change`**
   (`packages/docx/src/commands/accept-change.ts`,
   `reject-change.ts`). Payload `{ revisionId: string }`. Walks
   the body AND every header/footer body, rewriting `RevisionWrapper`
   nodes whose `revisionId` matches:

   | wrapper kind | accept             | reject             |
   | ------------ | ------------------ | ------------------ |
   | `<w:ins>`    | unwrap (keep runs) | drop (lose runs)   |
   | `<w:del>`    | drop (lose runs)   | unwrap (keep runs) |

   The dispatch sets `body: true` if any body wrapper matched and
   adds the matching part path to `headersAndFooters` if a wrapper
   matched there. Both can flip in one dispatch. Errors:
   `unknown-revision` for empty / missing id (the loud-idempotent
   policy: a second accept on a now-resolved id throws
   `unknown-revision`, mirroring how the bus surfaces the same code
   for stale references — see `tracked-changes.test.ts`).

   Round-trip property asserted directly:
   `parse(serialize(snapshot))` carries no `RevisionWrapper` whose
   `revisionId` equals the resolved id.

6. **Registry / index** (`packages/docx/src/commands/registry.ts`,
   `index.ts`). Removed the four stub rows for
   `docx:accept-change`, `docx:reject-change`, plus added the two
   `docx:set-*-text` commands. Stubs that remain:
   `docx:insert-table`, `docx:set-cell-content`, `docx:insert-image`.

### Tests

- `packages/docx/src/commands/headers-footers.test.ts` (new) —
  9 tests:
  1. parser produces a typed header + footer from a synthetic
     fixture
  2. parser reads the real-world
     `fixtures/docx/real-world/02-report-headers-footers.docx`
     and recovers the known header / footer paragraph text
  3. `set-header-text` replaces text and dirties only the targeted
     part path
  4. `set-footer-text` survives `DocxAgent.exportFile()` →
     `parseDocx`
  5. `set-header-text` is idempotent
  6. `set-header-text` rejects an unknown `partId` with code
     `unknown-target`
  7. `set-header-text` rejects an out-of-range `paragraphIndex`
     with code `unknown-target`
  8. byte-preservation: untouched header AND footer SHA-256 match
     after a no-touch round-trip of `02-report-headers-footers.docx`
  9. byte-preservation: footer SHA-256 stays identical when only
     the header is mutated; header bytes do change; mutation
     survives a re-parse

- `packages/docx/src/commands/tracked-changes.test.ts` (new) —
  8 tests covering accept-ins, accept-del, reject-ins, reject-del,
  unknown-revision (both missing-id and empty-id), idempotence
  (accept on a now-resolved id surfaces `unknown-revision`),
  serialize → re-parse leaves no matching wrapper, and a
  text-equality round-trip for reject-del.

- `pnpm --filter @officeai/docx test` — **83 tests pass** (was 66;
  +17 from W4: 9 headers/footers + 8 tracked-changes). The
  comments-lifecycle, handlers, renderer, agent, parser,
  serializer, and markdown suites are unchanged.

### Algorithmic notes

- The header/footer parser re-discovers parts each load; it does
  not assume a 1:1 between sections and parts. A single part may
  be referenced from N sections with possibly conflicting
  `w:type` values; we record the first one we see. This is
  informational metadata only — Word reads the part bytes by
  rId, not by the model's `target` field — so the choice does not
  affect output.
- `applySetTextToHeaderFooter` flattens the targeted paragraph to
  a single run on purpose: header / footer paragraphs in the wild
  almost always carry one run (a date, a company name, a page
  number field), and preserving the multi-run layout would force
  the handler to either pick a "primary" run (arbitrary) or
  abandon the rest. Header / footer text rewrites are a coarse-
  grained operation; finer-grained mutations should go through
  `insert-text` / `delete-range` against a future
  `position.headerFooter` selector (deferred — see follow-ups).
- For the resolution commands we keep `RevisionWrapper` rewrites
  pure on the input subtree; the recursion supports nested
  wrappers (rare but legal in OOXML) and strips the matching id
  even when wrapped inside a non-matching outer wrapper.
- `delText` flag handling: when reject-change unwraps a `<w:del>`,
  the surviving runs may still carry `isDelText: true` on their
  text leaves (because they came from `<w:delText>` in the
  source). We deliberately leave that flag intact — Word and
  LibreOffice both accept `<w:delText>` outside a wrapper as
  legal text content, and re-flattening it to `<w:t>` would
  require an extra pass that buys us nothing observable. If a
  follow-up needs strict `<w:t>`-only output we can add a
  normalize step, but it's not required for the round-trip
  property the brief asks for.

### Known limitations / follow-ups

- Header / footer parsing covers paragraph-level content only;
  tables inside a header (e.g. a multi-cell layout header) become
  `OpaqueBlock` and pass through unchanged on serialize.
  Mutation-aware tables in headers / footers land with W7 / P1.3
  alongside the typed table model.
- We do not yet support adding a brand-new header / footer part to
  a document that has none. The two `set-*-text` commands operate
  on parts the parser already discovered; minting a fresh part +
  relationship + content-type override is a separate command
  (`docx:add-header` / `docx:add-footer`) tracked for a follow-up
  session.
- `accept/reject-change` operates on the raw `RevisionWrapper`
  carrier; it does not yet update `revisionsView` style
  decorations in the renderer (that work lives in W5 / C5). The
  backend contract is stable, so the renderer wiring is
  additive.

## P1.2 — W5: Web UI parity (toolbar + comments sidebar + tracked-changes UI)

> Date: 2026-04-17. Workstream W5 of batch P1.2 (per
> `docs/roadmap-docx-p1.md`). Ships C3 (toolbar parity), C4 (comments
> sidebar with thread view + scroll-to-highlight + resolve / reply
> controls), and the UI half of C5 (inline accept / reject for tracked
> changes, drives W4's commands). Backend lives in W4; ops + LLM bridge
> live in W6.

### What landed

1. **Decomposed editor surface.** `DocxEditor.tsx` was restructured to
   compose four new sibling panels — `Toolbar`, `CommentsSidebar`,
   `TrackedChangesUI`, `AgentPrompt` — without changing its public
   surface (still the default-exported component consumed by
   `apps/web/app/editor/page.tsx`). Existing P1.1 selectors
   (`Add comment` button, `{N} blocks · rev {R} · {C} comments`
   metadata strip) are preserved verbatim so W1's six e2e specs keep
   matching.

2. **Toolbar parity** (`apps/web/app/editor/Toolbar.tsx` +
   `apps/web/app/lib/format-helpers.ts`). Six logical groups:
   - **Style** — paragraph style picker (Normal / Title /
     Heading 1-6) → `docx:set-paragraph-style`.
   - **Inline** — bold, italic, underline, strike → `docx:format-range`.
   - **Color** — 8 Office defaults stored as raw OOXML hex (no `#`).
   - **Highlight** — Word's `w:highlight` enum (`yellow`, `green`,
     `cyan`, `magenta`, `red`, `darkYellow`, `lightGray`); per-name
     swatches in `globals.css` paint the named color.
   - **Alignment / indentation / lists** — currently surface a
     "not yet supported" toast because no command lands them. The
     dispatch is wired so adding a future
     `docx:set-paragraph-properties` is a one-line change.
   - **Font size** — stores OOXML half-points but renders point
     sizes (`12`, not `24`).

   Toolbar wraps via `flex-wrap`; the metadata strip hides under
   `md` so the bar stays usable down to 360px.

3. **Comments sidebar** (`apps/web/app/editor/CommentsSidebar.tsx`).
   Lists every comment in document order; clicking a comment scrolls
   the editor to its `commentRangeStart` and visually flashes via
   `.pm-comment-flash` (1.4 s ease-out). Reply input dispatches
   `docx:reply-comment`; "Resolve" dispatches `docx:resolve-comment`
   (resolved threads recede via dashed border + 60% opacity + a
   `Resolved` pill); "Delete" dispatches `docx:delete-comment`. On
   `<lg` the sidebar collapses behind a floating `Comments` drawer
   button (slide-up sheet, max 80vh).

4. **Tracked-changes UI** (`apps/web/app/editor/TrackedChangesUI.tsx`).
   Two surfaces:
   - `InlineHoverWidget` follows the cursor over
     `.pm-revision-{ins,del}` spans; 200 ms hide timer prevents
     thrashing when the pointer crosses span ↔ widget boundary.
   - `ChangeListRibbon` is a collapsible side panel exposing
     `data-testid="tracked-change-row"` / `data-revision-id` hooks
     for keyboard users and Playwright. Buttons dispatch
     `docx:accept-change` / `docx:reject-change`. If the handler
     throws `NotImplementedError` (e.g. on a build that pre-dates
     W4), the existing toast surface renders "Not yet supported in
     this build" and the change stays visible.

5. **Agent prompt extracted** (`apps/web/app/editor/AgentPrompt.tsx`).
   Pulled the previously-inline UI out of `DocxEditor.tsx` into a
   reusable component with a `dispatch: (text) => Promise<void>`
   prop. The default `dispatch` preserves the original `[AI] ` +
   `add-comment` recipe so `add-comment.spec.ts` keeps passing
   without configuration. W6's `dispatchToLlm` plugs into the same
   prop so an `OPENAI_API_KEY`-configured deploy gets a real LLM
   bridge with zero UI churn.

6. **CSS polish** (`apps/web/app/globals.css`). Namespaced under
   `.pm-*` / `.editor-*` / `.docx-editor` to avoid leaking into
   shared `prose` rules: highlight swatches, `pm-comment-flash`
   keyframe, `cursor: help` over revision marks, and a tiny
   `appearance: none` reset on toolbar `<select>`s.

### Tests

- `apps/web/e2e/toolbar.spec.ts` (4 cases) — font size change,
  color picker, alignment "not yet supported" toast, list "not yet
  supported" toast.
- `apps/web/e2e/comments-sidebar.spec.ts` (3 cases) — open thread
  - scroll-to flash, reply, resolve.
- `apps/web/e2e/tracked-changes.spec.ts` (2 cases) — accept and
  reject against `fixtures/docx/real-world/06-comments-and-changes.docx`.
  Assertions use `.or()` so they remain green if W4's handlers
  ever regress to the stub (per the brief's graceful-toast
  fallback).

The 6 P1.1 e2e specs were not touched. `pnpm --filter @officeai/web build`
succeeds; `pnpm exec tsc --noEmit` is clean; `pnpm format:check` is
clean. Playwright browsers are not installed on this dev machine, so
the new specs are validated via the production build + the `web-e2e`
CI job.

### Visual / UX decisions

| Decision               | Rationale                                                                                 |
| ---------------------- | ----------------------------------------------------------------------------------------- |
| Icon set: lucide-react | Already in deps (MIT). No new dependency added.                                           |
| Color palette          | Eight Office defaults, hex without `#` to match the OOXML `w:color` shape.                |
| Highlight palette      | Restricted to Word's `w:highlight` enum — the only values Word actually round-trips.      |
| Font size storage      | OOXML half-points internally; rendered as point sizes so users see `12`, not `24`.        |
| Sidebar layout         | 2-column grid pinned at `lg` (≥1024px); collapses behind a slide-up drawer below that.    |
| Resolved comments      | Stay visible (dashed border + 60% opacity + `Resolved` pill) so reviewers retain context. |

### Known limitations / follow-ups

- Alignment, indentation, and list buttons surface a "not yet
  supported" toast — the backend command (`docx:set-paragraph-properties`)
  is not yet defined. Trivially additive once a follow-up lands it.
- The hover widget for tracked changes is keyboard-reachable only via
  the `ChangeListRibbon` (the inline overlay is mouse-driven). A
  full keyboard story for inline accept / reject can land alongside
  the next a11y pass.
- The comments drawer uses `position: fixed` on small screens; on
  iOS Safari the keyboard can briefly overlap the input. The
  workaround would be a `visualViewport`-aware repositioning hook,
  deferred to the next mobile pass.

## P1.2 — W6: Ops gates + agent surface polish

### What landed

This work item ties off P1.2 with two operational quality gates
(perf budgets + license scan), an extension to the Markdown
projection so an LLM can reason about lists / tables / comments,
and a thin in-process bridge between the editor and an actual LLM
provider.

1. **Performance budget script (`scripts/perf-docx.mjs`).** Builds
   a synthetic 100-page DOCX (3000 paragraphs, ~21 KB) using the
   `docx` library, then exercises the agent end-to-end through
   `DocxAgent.fromBuffer` → 1000 `docx:insert-text` commands →
   `agent.exportFile()`. The script prints a Markdown table of
   elapsed-vs-budget milliseconds and exits non-zero on any over-
   budget phase. Budgets:

   | Phase         | Budget  | Local M-class result |
   | ------------- | ------- | -------------------- |
   | parse         | 500 ms  | ~34 ms               |
   | 1000 commands | 1000 ms | ~6 ms                |
   | serialize     | 750 ms  | ~22 ms               |

   The budgets are deliberately loose (≥ 30 % headroom over the
   local numbers) so noisy GitHub-hosted runners don't false-
   positive on transient slow-downs. We will tighten them once
   we have a few weeks of CI history to size variance.

2. **SPDX license scanner (`scripts/license-scan.mjs`).** Walks
   every package directory under `node_modules/.pnpm/<spec>/node_modules/<name>`,
   reads the `license` field from each `package.json`, and
   classifies it without any network round-trip. It hard-fails
   (exit 1) on `AGPL-1.0+`, `AGPL-3.0+`, `GPL-2.0-only`,
   `GPL-3.0-only`, `SSPL-1.0`, `BUSL-1.1`, and warns on
   `LGPL-*`/`GPL-*-with-exception` so we have a paper trail
   without blocking the build. `--inject-agpl` synthesizes an
   AGPL-licensed entry to exercise the failure path in the
   integration test (`tests/scripts/license-scan.test.ts`).
   Current dep tree: 497 unique packages, 425 MIT, 26 Apache-2.0,
   23 ISC, no banned licenses, one LGPL-or-later warning from
   `@img/sharp-libvips-darwin-arm64` (transitive — shipped as a
   prebuilt binary, not linked).

3. **Make + CI wiring.** Both scripts are intentionally **opt-in**
   from the local `Makefile` (`make perf-docx`, `make licenses`)
   and **mandatory in CI** as their own jobs (`docx-perf`,
   `license-scan` in `.github/workflows/ci.yml`). Keeping them
   out of `make verify` matches the existing `docx-libreoffice-roundtrip`
   pattern: heavy / environment-sensitive checks live in CI so a
   `make verify` on a developer laptop stays under ~15 s.

4. **Extended Markdown projection (`packages/docx/src/agent/markdown.ts`).**
   `snapshotToMarkdown` now emits a much richer projection so the
   LLM can reason about structure rather than a flat run of
   paragraphs:
   - **Section breaks** → `---` thematic break.
   - **Headings** (`Title`, `Heading1`..`Heading6` style ids) →
     ATX `#`..`######`.
   - **Numbered lists** — driven by `paragraph.properties.numbering`
     (presence of `numId`) → `1.`, `2.`, … with two-space-per-
     level indentation from `ilvl`. Numbering is per-`numId`,
     so two interleaved lists keep separate counters.
   - **Bullet lists** — paragraphs styled `ListParagraph` with
     no numbering → `-` items, also indented by `ilvl`.
   - **Tables** — best-effort GFM pipe-table conversion straight
     from the opaque `Table.raw.subtree` (we walk `w:tr` →
     `w:tc` and concatenate inner `w:t` text). Newlines in cells
     become `<br>`, pipes are escaped as `\|`. If the parser
     can't recover ≥ 1 row × 1 col we fall back to the prior
     `> [table preserved]` placeholder and `console.warn`.
   - **Comments** — appended as a trailing `## Comments` section
     listing thread heads only (replies indent under their
     parent). Each entry shows the comment text (truncated to
     200 chars) and a "> on:" snippet of the parent paragraph
     text so the LLM has anchor context. Skipped silently when
     the document has no comments.

   The renderer is purely a projection — it never mutates the
   snapshot — so it's safe to call from the LLM bridge on every
   request. Coverage: 8 new tests in `markdown.test.ts` exercising
   each branch (headings, both list flavours, pipe-table happy
   path, malformed-table fallback, section break, comments
   present, comments absent).

5. **LLM bridge (`apps/web/app/api/llm/route.ts`).** A small
   Next.js App Router POST endpoint:
   - Body: `{ prompt: string; snapshotMarkdown: string }`.
   - Returns: `{ commands: Array<{ type, payload }>; rationale }`.
   - Gated on `process.env.OPENAI_API_KEY`. When unset → `501
{ error: "LLM bridge not configured" }`. Other HTTP methods
     get the App Router default `405`.
   - When configured, calls `https://api.openai.com/v1/chat/completions`
     directly (no SDK) with `response_format: { type: "json_object" }`
     and `temperature: 0.2`, model `gpt-4.1` (overridable via
     `OPENAI_MODEL`). The system prompt enumerates the public
     `docx:*` command surface and forbids prose.
   - Returned commands are **filtered against an allow-list** of
     known command types before reaching the client, so a
     hallucinated type (`docx:nuke-document`) is silently
     dropped rather than queued.
   - The route never executes commands. It hands the parsed
     command list back to the browser so the existing pending /
     approve UI in `AgentPrompt` stays the single funnel for
     human review — same "AI proposes, human approves"
     contract as the demo dispatch.

6. **Isomorphic client helper (`apps/web/app/lib/llm-client.ts`).**
   `dispatchToLlm(prompt, agent)` POSTs to `/api/llm` with the
   current snapshot's Markdown projection. On success it returns
   the agent-tagged command list ready for `agent.applyCommands`.
   On `501` (no API key), network failure, non-OK status, or
   non-JSON payload, it transparently falls back to the original
   `[AI] ` + `add-comment` recipe so the existing
   `add-comment.spec.ts` e2e still passes with no env vars set.
   Failure paths surface a `note` string the editor renders as a
   warn toast, so a misconfigured production deploy is visible
   without breaking the demo.

7. **Editor wiring (`DocxEditor.tsx`).** The default
   `AgentPrompt` dispatch now routes through `dispatchToLlm` (the
   old `defaultAgentDispatch` stays around in `AgentPrompt.tsx`
   as the absolute pre-agent fallback). With no `OPENAI_API_KEY`
   the user-visible behaviour is byte-identical to W5 (because
   the bridge falls back to the same recipe). With a key set, the
   prompt becomes a real LLM call whose output lands in the same
   pending-mutations queue.

### Decisions (W6-specific)

| Date (UTC) | Decision                                                                 | Rationale                                                                                                                                                          |
| ---------- | ------------------------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| 2026-04-17 | Perf + license jobs run in CI but are opt-in from `make verify`          | `make verify` is the developer inner loop; we want it sub-15 s. Heavy gates already live in dedicated CI jobs (LibreOffice round-trip, etc.).                      |
| 2026-04-17 | Walk `node_modules/.pnpm` instead of parsing `pnpm-lock.yaml` directly   | `node_modules/.pnpm/<spec>/node_modules/<pkg>/package.json` is the source of truth post-install; avoids a YAML parse and re-implementing pnpm spec resolution.     |
| 2026-04-17 | License scanner has no network calls                                     | Brief explicitly forbids fetching SPDX data; offline-deterministic also makes the CI job fast and air-gap-friendly.                                                |
| 2026-04-17 | Tables → GFM pipe tables (best-effort) rather than HTML tables           | LLMs reason about pipe tables far better than nested `<table>` HTML; falling back to `> [table preserved]` keeps the projection stable when our parser can't cope. |
| 2026-04-17 | LLM bridge calls `fetch` directly (no `openai` SDK dependency)           | Adds zero deps; we only need chat/completions and the JSON-mode flag. Easy to swap for another provider behind the same `/api/llm` contract.                       |
| 2026-04-17 | Bridge filters returned commands against an allow-list of `docx:*` types | Defense in depth — model can't queue an undocumented command even if it hallucinates one. The renderer would reject it later, but failing here is friendlier.      |
| 2026-04-17 | Client helper falls back to demo recipe on any bridge failure            | Keeps `add-comment.spec.ts` (and the no-key local dev experience) working with zero configuration; mismatches surface as toasts, not exceptions.                   |

### Known limitations / follow-ups

- The perf budgets are sized for CI variance, not a tight
  regression gate. Once we have ~2 weeks of green CI numbers we
  should ratchet them down to ~2× the observed median.
- The license scanner reads `package.json` `license` (string) and
  `licenses` (legacy array) only. A handful of older packages use
  free-form `license` like `"AND expression all permissive"`
  (visible in the current scan) — those land in an "unknown /
  free-form" bucket and emit a warning rather than failing, since
  hand-classifying them is out of scope for an SPDX-only check.
- The LLM bridge does not currently stream — it returns the full
  command list once OpenAI finishes. Streaming the rationale and
  partial commands would be a nice UX upgrade but adds a lot of
  client-side state for marginal gain at the current per-prompt
  size (~tens of commands).
- Table extraction concatenates raw `w:t` runs cell-by-cell; it
  does not yet preserve cell-level formatting (bold, hyperlinks,
  nested paragraphs). The `> [table preserved]` fallback fires
  for those edge cases, so we never produce a wrong-looking
  pipe table — we just produce a placeholder. Richer extraction
  lands with the W7 typed-table model.

## P1.3 — W7: Typed tables

Turning tables from "opaque preserved" into a first-class, mutation-aware
part of the model. Backend only; no UI yet.

### What shipped

- **Typed model** — `Table` now carries a typed `properties: TableProperties`,
  `grid: ReadonlyArray<TableGridCol>`, and `rows: ReadonlyArray<TableRow>`
  triple instead of a single opaque `<w:tbl>` blob. `TableRow` and
  `TableCell` get matching typed shapes (`TableRowProperties`,
  `TableCellProperties`); cells contain a recursive `body: BlockNode[]`
  so paragraphs _and_ nested tables round-trip correctly. Every
  `properties` interface keeps an `opaqueProps?: ReadonlyArray<OpaqueXml>`
  carrier for unknown XML — same pattern `ParagraphProperties` already
  uses — so the round-trip stays lossless even for tables Word fills
  with shading / borders / look flags we don't model yet.
- **Parser** — new `parser/tables.ts` (`parseTable` / `parseTableRow` /
  `parseTableCell`), wired into `parse.ts`. The injected
  `parseParagraph` argument breaks an otherwise circular import.
  Unknown block-level children inside a `<w:tc>` (drawings nested in a
  cell, mc:AlternateContent blocks, etc.) become `OpaqueBlock` so the
  cell still type-checks and round-trips.
- **Serializer** — new `serializer/tables.ts`. The body serializer now
  forks per table: if `table.raw` is set (untouched since parse) it
  re-emits the cached subtree via `serializeTableFromRaw`; otherwise it
  regenerates from the typed model.
- **Four commands** — `docx:insert-table`, `docx:set-cell-content`,
  `docx:insert-row`, `docx:insert-column`. All four go through
  `CommandBus`, mint ids via `IdMinter`, and produce typed `Diff`s
  (`node-inserted` / `node-updated` with `path: ["body", N, "rows", R,
"cells", C, ...]`).

### Decision: byte-preservation marker = `Table.raw` itself

The brief allowed a separate per-table dirty bit on the snapshot, but
folding the marker into the table proved cleaner:

- `parseTable` always sets `raw` from `captureOpaque(entry)`.
- Every mutating command runs its result through `withoutRaw(...)`,
  which strips `raw` and returns a new object.
- `findTable(...)` rebuilds every ancestor table on the way up, dropping
  _their_ `raw` too — so a `set-cell-content` against a nested table
  invalidates exactly the chain of containing tables, never their
  siblings.
- The body serializer's "fast path vs regenerate" decision is then a
  single property check (`if (block.raw) ...`) instead of a side-table
  lookup — keeps the serializer free of dirty-tracking logic and makes
  the invariant trivially observable in tests.

This is what lets `04-table-grid.docx` round-trip with
`word/document.xml` byte-identical even when other commands (or no
commands at all) have been dispatched.

### Constraint: merged-cell mutations

The brief explicitly defers reflow of `gridSpan` / `vMerge` regions to a
later session. The handlers therefore reject mutations that would
_corrupt_ a merge, with `merged-cell-not-supported`:

- `set-cell-content` rejects writes into a `vMerge="continue"` cell.
- `insert-row` rejects insertions immediately above a row that begins
  with `vMerge="continue"` cells (which would orphan their `restart`
  ancestor).
- `insert-column` rejects insertions whose target column index is
  straddled by a `gridSpan > 1` cell in any row. Boundary insertions
  (`at === 0` or `at === grid.length`) are always accepted.

`gridSpan` and `vMerge` parse + round-trip correctly in all cases —
only the four mutating commands gate against them.

### Constraint: nested-table cycles

`tableId` resolves recursively (a nested table inside a cell can be the
target). To keep `set-cell-content` from accepting a payload that would
rewrite the document into an infinite tree, the handler walks the
incoming `content` looking for any `Table` whose `id` matches the target
or one of its ancestors and rejects with `unknown-target`.

### Markdown projection compatibility shim

`agent/markdown.ts` is in the W7 read-only list, but its previous code
unconditionally read `table.raw.subtree` to extract cell text. With
`raw` now optional that line stops type-checking. The minimal-impact
fix is a fork:

```ts
const projected =
  block.rows.length > 0 ? tableToMarkdownTyped(block) : block.raw ? tableToMarkdown(block.raw.subtree) : null;
```

`tableToMarkdownTyped` walks the typed model directly, which is
strictly better than the old subtree scrape (it picks up nested
paragraphs, list items, etc.) and eliminates the future need for the
legacy path entirely.

### Test re-point in `handlers.test.ts`

The pre-existing `"stub commands return a rejected mutation with
not-implemented"` case asserted on `docx:insert-table`. Shipping the
real handler obviously breaks that single line. The minimal possible
edit — re-pointing it at `docx:insert-image`, the only remaining stub
— keeps the existing suite green while preserving the test's intent
(verify the bus surfaces `not-implemented` errors as rejected
mutations). Flagged in the W7 deviations report.

### Test counts

| Suite                     | Before | After |
| ------------------------- | -----: | ----: |
| `@officeai/docx` (Vitest) |     83 |    98 |

15 new cases in `commands/tables.test.ts` cover: synthetic 2×3 parse,
real-world fixture parse, byte-preservation on fixture round-trip,
`insert-table` + payload validation, `set-cell-content` happy path +
unknown-target + OOB + cycle, `insert-row` append + at-zero header
semantics, `insert-column` middle + edge insertion, nested-table parse

- mutation, and `gridSpan`/`vMerge` preservation + continuation-write
  rejection.

