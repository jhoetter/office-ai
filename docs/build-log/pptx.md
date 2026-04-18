# PPTX Build Log

> Live log of decisions, deviations from spec, and known issues for the
> PPTX phase. Mirrors `docs/build-log/docx.md`.

## Decisions

| Date (UTC) | Decision                                                                              | Rationale                                                                                                                                |
| ---------- | ------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------- |
| 2026-04-18 | Reuse `@officeai/core` `OoxmlContainer` + `RelationshipGraph` + `ContentTypes`        | Same machinery already used by DOCX; keeps parser/serializer surface tiny and shares the byte-cache that drives roundtrip integrity.    |
| 2026-04-18 | Model unsupported shape kinds (charts, smart-art, custom geom) as `OpaqueShape`       | Preserves byte-for-byte roundtrip without trying to introspect every OOXML quirk. Selectable + movable later via raw-attr edits if ever. |
| 2026-04-18 | Hybrid renderer: SVG geometry + HTML overlay for editable text                        | SVG handles vector fidelity (gradients, prstGeom, blipFill) cheaply; contenteditable HTML gives us native IME, caret, selection.        |
| 2026-04-18 | All commands address shapes by `(slideIndex, NodeId)`, never by cNvPrId               | NodeIds are model-internal and survive split/merge; cNvPrId is OOXML-internal and only unique per slide.                                 |
| 2026-04-18 | `insert-image` SHA-256-dedups against `presentation.media`                            | Avoids ballooning the file when an LLM pastes the same screenshot into many slides; matches Word/PowerPoint behaviour.                   |
| 2026-04-18 | Renderer split into `renderer/layout` (pure), `renderer/svg` (pure), `renderer/react` | Keeps the headless invariant (`headless-invariant.test.ts`) honest: only `renderer/react` may import React.                              |
| 2026-04-18 | `OFFICEAI_DETERMINISTIC_IDS=1` env var swaps the UUID minter for the deterministic one | Lets the CLI tests address shape NodeIds across multiple `office-agent pptx …` invocations without baking a deterministic mode into prod.|

## Deviations from spec

| Date (UTC) | Spec section                | Deviation                                                                                          | Reason                                                                                                                                                       |
| ---------- | --------------------------- | -------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| 2026-04-18 | `agent-commands.md` (`pptx:format-text`) | Re-emits only `a:solidFill` and `a:latin` from typed properties; other captured `a:rPr` children pass through verbatim | Keeps `schemeClr`/`hlinkClick`/etc. round-tripping; we'd need a typed model for every child of `a:rPr` to fully own the element, which is out-of-scope for P1.|
| 2026-04-18 | `renderer.md` (theme colors) | _(resolved 2026-04-19, F1.2)_ Parser extracts the first theme part's `a:clrScheme` into `PptxPresentation.themeDefault`; renderer threads it through `SvgRenderCtx.theme` and resolves `<a:schemeClr>` references inside `a:solidFill` opaque children | The opaque `a:solidFill > a:schemeClr` capture stays unchanged so the serializer round-trips cleanly. |
| 2026-04-18 | `feature-scope.md` (LLM bridge) | _(resolved 2026-04-19, F1.1)_ The `/pptx-editor` agent panel now routes through `/api/llm` with `format: "pptx"`; the in-process intent parser is the offline fallback when no `OPENAI_API_KEY` is set | Brings PPTX to parity with DOCX. |

### Resolved deviations

| Date (UTC) | Item                                  | Resolution                                                                                                  |
| ---------- | ------------------------------------- | ----------------------------------------------------------------------------------------------------------- |
| 2026-04-19 | LLM bridge wiring for `/pptx-editor` | F1.1 — extended `/api/llm` to switch system prompt + allow-list on `format`, added `lib/llm-client-pptx.ts`. |
| 2026-04-19 | Theme color resolution                | F1.2 — `parser/theme.ts` extracts `a:clrScheme` into `PptxPresentation.themeDefault`; renderer resolves `<a:schemeClr>` from run opaque children. |

## Deferred to a follow-up session

| Item                                                       | Spec ref            | Status                                                |
| ---------------------------------------------------------- | ------------------- | ----------------------------------------------------- |
| Tables (`a:tbl`)                                           | `agent-commands.md` | Stays opaque this session — typed model deferred.     |
| Charts (`c:chart`)                                          | `agent-commands.md` | Stays opaque (handled by `OpaqueShape`).              |
| SmartArt (`dgm:*`)                                          | `agent-commands.md` | Stays opaque.                                         |
| Animations / transitions                                   | `feature-scope.md`  | Untouched-bytes only; spec defers typed model.        |
| ~~Theme color resolution (`a:schemeClr` → `theme1.xml`)~~   | ~~`renderer.md`~~   | **Resolved (F1.2)** — first theme part is parsed into `themeDefault`; renderer resolves `a:schemeClr` references for run fills. |
| ~~LLM bridge wiring for `/pptx-editor` agent panel~~       | ~~`feature-scope.md`~~ | **Resolved (F1.1)** — uses `/api/llm` with `format: "pptx"`. |
| Real-world PPTX fixtures (PowerPoint / Google Slides / Keynote exports) | `feature-scope.md`  | Slots reserved in `fixtures/pptx/MANIFEST.md`.        |

## Phase summary

| Phase | Scope                                                                                            | Tests added |
| ----- | ------------------------------------------------------------------------------------------------ | ----------- |
| P0    | Clean-room analysis vs. reference repos → `spec/pptx/analysis.md`                               | —           |
| P1    | Full `spec/pptx/*.md` set mirroring `spec/docx/`                                                 | —           |
| P2    | `scripts/generate-pptx-fixtures.mjs` + 8 synthetic fixtures + manifest                            | —           |
| P3    | Parser + serializer with dirty-flag-driven, byte-preserving roundtrip                            | parser + serializer suites |
| P4    | `PptxAgent` skeleton + headless-invariant test                                                    | agent suite |
| P5    | P0 command handlers (`add/delete/duplicate/move-slide`, `set-text/-position/-size`)               | `p0-commands.test.ts` (13 tests) |
| P6    | P1 command handlers (`format-text`, `insert-image` with SHA-256 dedup, `add-text-box`)            | `p1-commands.test.ts` (10 tests) |
| P7    | Hybrid renderer (`renderer/layout`, `renderer/svg`, `renderer/react`)                              | layout + svg unit tests |
| P8    | `/pptx-editor` route in `apps/web` (toolbar, slides sidebar, canvas, agent panel, sample deck)    | manual browser smoke |
| P9    | `office-agent pptx …` CLI subcommand group + MCP tools (`pptx_load`/`_save`/`_inspect`/`_get_text`/`_search`/`_apply_command`/`_diff`) | `pptx-cli.test.ts` (13 tests) + 6 PPTX MCP tests |
| P10   | Validate (full repo tests, browser smoke), this build log, session summary                       | —           |

Per-package test totals at the end of P9:

- `@officeai/pptx`: **45 tests** in 10 files
- `@officeai/agent`: **43 tests** in 3 files (15 docx CLI + 15 MCP + 13 pptx CLI)

## Session summary (2026-04-18, "night shift")

End-to-end PPTX phase landed in 11 commits across 3 days:

| Commit  | Phase | Summary                                                                              |
| ------- | ----- | ------------------------------------------------------------------------------------ |
| c62fe64 | P0    | `spec/pptx/analysis.md` — clean-room study vs. reference repos                       |
| 96befa7 | P1    | Full `spec/pptx/*.md` mirror of `spec/docx/`                                         |
| 92afd98 | P2    | Synthetic fixtures + generator + manifest                                            |
| 645eba2 | P3    | Parser + serializer (byte-preserving, dirty-flag-driven)                             |
| d28a3a7 | P4    | `PptxAgent` skeleton + headless invariant test                                       |
| 1a6c2ec | P5    | P0 commands (`add/delete/duplicate/move-slide`, `set-text/-position/-size`)          |
| d4e8880 | P6    | P1 commands (`format-text`, `insert-image` with SHA-256 dedup, `add-text-box`)       |
| 6f10c3f | P7    | Hybrid renderer (`renderer/{layout,svg,react}`)                                       |
| 4b3ad13 | P8    | `/pptx-editor` Next.js route — toolbar, sidebar, canvas, agent panel, sample deck    |
| 18140e6 | P9    | `office-agent pptx …` CLI subcommand group + 7 PPTX MCP tools                        |
| _next_  | P10   | Validate + browser smoke + this build log + final commit                             |

Final test/build totals at the close of P10:

- `@officeai/pptx`: 45 tests (parser, serializer, agent, P0/P1 commands, renderer layout/SVG, headless invariant)
- `@officeai/agent`: 43 tests (15 docx CLI + 15 MCP + 13 pptx CLI)
- `@officeai/web`: `pnpm build` succeeds, `/pptx-editor` listed as a dynamic route (740 B page bundle, 119 kB First Load JS).
- Architecture check: green (`scripts/check-architecture.mjs`).
- Browser smoke: home page links to `/pptx-editor`; the route mounts the toolbar, slides sidebar (with thumbnail), interactive canvas, and Agent panel; "Add slide" and "Text box" buttons round-trip through the command bus and re-render the sidebar thumbnail.

## Known issues

- ~~The `/pptx-editor` agent panel uses an in-process intent parser only.~~
  Resolved in F1.1: routes through `/api/llm` with `format: "pptx"`. The
  in-process intent parser is the offline fallback when no API key is set.
- `dpi` is hard-coded to 96 in the renderer's pixel conversions; an
  override hook is exported (`emuToPx(emu, dpi)`) but unused. Will
  matter once we expose a zoom slider.
- Browser smoke tests for `/pptx-editor` are manual via the
  `cursor-ide-browser` MCP; an automated Playwright pass mirroring
  `tests/e2e/docx-editor.spec.ts` is queued.
