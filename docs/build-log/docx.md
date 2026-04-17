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

