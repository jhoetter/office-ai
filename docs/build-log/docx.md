# DOCX Build Log

> Live log of decisions, deviations from spec, and known issues.

## Decisions

| Date (UTC) | Decision | Rationale |
| --- | --- | --- |
| 2026-04-17 | Use `fast-xml-parser` with `preserveOrder: true` over a custom XML model | Round-trips losslessly for OOXML; smaller surface than building our own. |
| 2026-04-17 | Tables stored as opaque-XML `Table.raw` for now (P1 mutation) | Cell merging is genuinely tricky; preserving roundtrip integrity matters more than editing this session. |
| 2026-04-17 | Hyperlink modeled as a typed wrapper over runs (not a mark) | Matches OOXML structure (`w:hyperlink` is a block element nesting runs); avoids mark-coalescing ambiguity. |
| 2026-04-17 | Comments staging tri-state lives in core, not docx | Same logic will serve XLSX/PPTX. |
| 2026-04-17 | The serializer trusts dirty flags rather than diffing snapshots | Cheap, predictable, and matches our command-bus discipline. |
| 2026-04-17 | Touched parts may differ from input in attribute order / quote style | Word and LibreOffice both accept either; we assert structural equivalence on touched parts and bytewise equality only on untouched parts. |

## Deviations from spec

| Date (UTC) | Spec section | Deviation | Reason |
| --- | --- | --- | --- |
| 2026-04-17 | `agent-commands.md` (`format-range`, `delete-range`) | Multi-paragraph ranges throw `CommandError` for now | First cut; single-paragraph handling already exercises the run-splitting machinery. P1 will lift the restriction. |
| 2026-04-17 | `set-paragraph-style` payload | `style` is not validated against `word/styles.xml` | We don't model styles yet; we trust the AI / human to use a known styleId. Will tighten when we ingest the styles part. |
| 2026-04-17 | `core/util/hash` | Switched from `node:crypto` to the `js-sha256` package | Required isomorphic implementation so the same `OoxmlContainer` works in `apps/web` (browser bundle) and Node. Output bytes are identical. |
| 2026-04-17 | `renderer.md` | `transactionToCommands` does not yet emit `docx:set-paragraph-style` from PM transactions | The toolbar route still hits the agent directly; we'll wire the PM keymap branch when we add a real style menu. |

## Deferred to a follow-up session

| Item | Spec ref | Status |
| --- | --- | --- |
| `docx:insert-table` | `agent-commands.md` | Stub — throws NotImplementedError |
| `docx:set-cell-content` | `agent-commands.md` | Stub |
| `docx:insert-image` | `agent-commands.md` | Stub |
| `docx:resolve-comment` | `agent-commands.md` | Stub |
| `docx:accept-change` | `agent-commands.md` | Stub |
| `docx:reject-change` | `agent-commands.md` | Stub |
| Real-world DOCX fixtures (Word/Google/LibreOffice exports) | `feature-scope.md` | Slots reserved in `fixtures/docx/MANIFEST.md` |
| LibreOffice CI roundtrip | `feature-scope.md` | Manual today; CI integration deferred |
| Headers/footers editing | `feature-scope.md` | P1 — preserved verbatim; not mutable |
| Image insertion | `feature-scope.md` | P1 — preserved on roundtrip; not mutable |
| List mutation (numbering insert/remove) | `feature-scope.md` | P1 — preserved; style-by-name only |

## Known issues

| Issue | Detail | Mitigation |
| --- | --- | --- |
| ProseMirror funnel does not preserve nested marks across boundary edits | When a ReplaceStep crosses a mark boundary (e.g. typing inside a partially-bold span), we currently translate it as plain `insert-text` without re-asserting the surrounding mark. | Acceptable for P0. P1: read marks at insertion point and follow with a `format-range` if needed. |
| `transactionToCommands` block-bearing slice path always emits `insert-paragraph at start` regardless of slice content | Good enough for "press Enter at end of paragraph" but not for paste-of-multiple-paragraphs. | Treat as opaque "Action deferred — see build log." Toast surfaces this in the web UI. |
| Web demo agent prompt is hard-coded ("[AI] " + comment) | The point of the prompt panel is to demonstrate the human-review queue; an actual LLM call lives outside this session's scope. | Documented in the `/editor` UI copy. |

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
