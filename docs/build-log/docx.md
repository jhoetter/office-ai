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

| Date (UTC) | Spec section                                         | Deviation                                                                                 | Reason                                                                                                                                     |
| ---------- | ---------------------------------------------------- | ----------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------ |
| 2026-04-17 | `agent-commands.md` (`format-range`, `delete-range`) | Multi-paragraph ranges throw `CommandError` for now                                       | First cut; single-paragraph handling already exercises the run-splitting machinery. P1 will lift the restriction.                          |
| 2026-04-17 | `set-paragraph-style` payload                        | `style` is not validated against `word/styles.xml`                                        | We don't model styles yet; we trust the AI / human to use a known styleId. Will tighten when we ingest the styles part.                    |
| 2026-04-17 | `core/util/hash`                                     | Switched from `node:crypto` to the `js-sha256` package                                    | Required isomorphic implementation so the same `OoxmlContainer` works in `apps/web` (browser bundle) and Node. Output bytes are identical. |
| 2026-04-17 | `renderer.md`                                        | `transactionToCommands` does not yet emit `docx:set-paragraph-style` from PM transactions | The toolbar route still hits the agent directly; we'll wire the PM keymap branch when we add a real style menu.                            |

## Deferred to a follow-up session

| Item                                                       | Spec ref            | Status                                        |
| ---------------------------------------------------------- | ------------------- | --------------------------------------------- |
| `docx:insert-table`                                        | `agent-commands.md` | Stub — throws NotImplementedError             |
| `docx:set-cell-content`                                    | `agent-commands.md` | Stub                                          |
| `docx:insert-image`                                        | `agent-commands.md` | Stub                                          |
| `docx:resolve-comment`                                     | `agent-commands.md` | Stub                                          |
| `docx:accept-change`                                       | `agent-commands.md` | Stub                                          |
| `docx:reject-change`                                       | `agent-commands.md` | Stub                                          |
| Real-world DOCX fixtures (Word/Google/LibreOffice exports) | `feature-scope.md`  | Slots reserved in `fixtures/docx/MANIFEST.md` |
| LibreOffice CI roundtrip                                   | `feature-scope.md`  | Manual today; CI integration deferred         |
| Headers/footers editing                                    | `feature-scope.md`  | P1 — preserved verbatim; not mutable          |
| Image insertion                                            | `feature-scope.md`  | P1 — preserved on roundtrip; not mutable      |
| List mutation (numbering insert/remove)                    | `feature-scope.md`  | P1 — preserved; style-by-name only            |

## Known issues

| Issue                                                                                                                 | Detail                                                                                                                                                                             | Mitigation                                                                                       |
| --------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------ |
| ProseMirror funnel does not preserve nested marks across boundary edits                                               | When a ReplaceStep crosses a mark boundary (e.g. typing inside a partially-bold span), we currently translate it as plain `insert-text` without re-asserting the surrounding mark. | Acceptable for P0. P1: read marks at insertion point and follow with a `format-range` if needed. |
| `transactionToCommands` block-bearing slice path always emits a single `insert-paragraph` regardless of slice content | Good enough for "press Enter at end of paragraph" but not for paste-of-multiple-paragraphs.                                                                                        | Treat as opaque "Action deferred — see build log." Toast surfaces this in the web UI.            |
| Web demo agent prompt is hard-coded ("[AI] " + comment)                                                               | The point of the prompt panel is to demonstrate the human-review queue; an actual LLM call lives outside this session's scope.                                                     | Documented in the `/editor` UI copy.                                                             |

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
