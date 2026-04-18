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
| 2026-04-19 | Playwright e2e for `/pptx-editor`     | F1.3 — `apps/web/e2e/pptx-editor.spec.ts` covers route mount, sample-deck thumbnail, Add slide / Delete enable, Text box, Bold toggle, agent panel "add a slide" via the LLM bridge offline fallback. |
| 2026-04-19 | Real-world PPTX fixtures + roundtrip integration test | F1.4 — `scripts/generate-real-pptx-fixtures.mjs` writes 3 third-party-emitter (`pptxgenjs`) decks to `fixtures/pptx/real/` (styled deck w/ hyperlinks + notes, mixed-media w/ table+image, 25-slide deck). `tests/roundtrip/pptx/real-world-roundtrip.test.ts` asserts ≥95 % byte-identity on pure roundtrip and edit-isolation on a single text-shape edit. |
| 2026-04-19 | Renderer DPI/zoom hook + toolbar zoom slider | F1.5 — `SlideCanvas` accepts `zoom` (clamped to [0.25, 3]) and `dpi` (overlay font sizing); `MIN_ZOOM` / `MAX_ZOOM` / `clampZoom` exported from `@officeai/pptx/renderer`. `PptxToolbar` gained zoom-out / range slider / zoom-in / reset (% display + click-to-100 %). New Playwright case in `apps/web/e2e/pptx-editor.spec.ts` drives the slider and asserts the canvas's `data-zoom` attribute. |

## Deferred to a follow-up session

| Item                                                       | Spec ref            | Status                                                |
| ---------------------------------------------------------- | ------------------- | ----------------------------------------------------- |
| Tables (`a:tbl`)                                           | `agent-commands.md` | Stays opaque this session — typed model deferred.     |
| Charts (`c:chart`)                                          | `agent-commands.md` | Stays opaque (handled by `OpaqueShape`).              |
| SmartArt (`dgm:*`)                                          | `agent-commands.md` | Stays opaque.                                         |
| Animations / transitions                                   | `feature-scope.md`  | Untouched-bytes only; spec defers typed model.        |
| ~~Theme color resolution (`a:schemeClr` → `theme1.xml`)~~   | ~~`renderer.md`~~   | **Resolved (F1.2)** — first theme part is parsed into `themeDefault`; renderer resolves `a:schemeClr` references for run fills. |
| ~~LLM bridge wiring for `/pptx-editor` agent panel~~       | ~~`feature-scope.md`~~ | **Resolved (F1.1)** — uses `/api/llm` with `format: "pptx"`. |
| ~~Real-world PPTX fixtures (third-party emitter)~~                       | ~~`feature-scope.md`~~ | **Resolved (F1.4)** — `fixtures/pptx/real/` shipped + integration test. PowerPoint / Google Slides / Keynote *direct* exports remain as `./real-world/` slots until we have license-clean originals. |

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
- ~~`dpi` is hard-coded to 96 in the renderer's pixel conversions.~~
  Resolved in F1.5: `SlideCanvas` accepts `zoom` + `dpi`; the toolbar
  exposes a zoom slider with a 100 % reset.
- ~~Browser smoke tests for `/pptx-editor` are manual via the
  `cursor-ide-browser` MCP; an automated Playwright pass mirroring
  `tests/e2e/docx-editor.spec.ts` is queued.~~ Resolved in F1.3:
  `apps/web/e2e/pptx-editor.spec.ts` covers mount + toolbar + agent
  panel via the existing Playwright config.
- Editing typed entrance animations on a slide drops `Slide.timingTailRaw`
  and rebuilds `<p:timing>` from the typed `Slide.animations` array. Any
  unmodeled timing children on that slide (exit animations, motion paths,
  emphasis effects, complex sequences) are lost on save. Slides whose
  animations are *not* edited continue to round-trip the original
  `<p:timing>` blob byte-for-byte.

## Session summary (2026-04-19, "follow-up shift")

After P10 closed, a follow-up shift landed five P0/P1 polish items
(F1.1–F1.5), three typed-model phases (Tables F2, Charts F3,
Animations F4 — SmartArt explicitly out of scope), and a final
validation pass:

| Commit  | Phase | Summary                                                                                  |
| ------- | ----- | ---------------------------------------------------------------------------------------- |
| 6f25a09 | F1.1  | `/pptx-editor` agent panel routes through `/api/llm` (`format: "pptx"`, allow-list)      |
| 7fa2591 | F1.2  | Theme-color resolution: parse `theme1.xml` → `themeDefault`; renderer resolves `schemeClr` |
| 773222e | F1.3  | Playwright smoke for `/pptx-editor` (mount, toolbar, agent panel via LLM offline fallback) |
| 6cf9098 | F1.3  | Ignore Playwright `test-results/` + `playwright-report/`                                  |
| 06e66cf | F1.4  | Real-world fixtures (`fixtures/pptx/real/`) + integration roundtrip test (≥95 % byte-identity) |
| 8fa9013 | F1.5  | Renderer DPI/zoom hook + toolbar zoom slider (clamped 0.25×–3×, 100 % reset)              |
| e404485 | F2.1–F2.4 | Typed Tables: `TableShape` parser + dirty serializer + 5 commands + `tableToSvg` renderer |
| c494e6b | F2.5  | Tables: CLI subcommands (`set-table-cell-text`, `add/delete-table-row/-column`), MCP visibility, projection |
| 95f278b | F3.0+F4.0 | Spec amendments for Charts + Animations (with explicit non-goals)                       |
| 795d1aa | F3.1+F3.2 | Typed `ChartShape` + `ChartPart` parser + dirty `ppt/charts/chart{N}.xml` serializer |
| 28588b1 | F3.3+F3.4 | Chart commands (`set-chart-title`, `set-chart-data`, `set-chart-type`) + native bar/line/pie/area SVG renderer |
| 48e411d | F3.5  | Charts: CLI + MCP + projection + tests + `09-with-chart.pptx` fixture                     |
| f585f0d | F4.1+F4.2 | Typed `SlideTransition` + entrance `EntranceAnimation` parser/serializer (model-driven `<p:timing>` rebuild) |
| b55f28e | F4.3+F4.4 | Animation commands (`set-slide-transition`, `add/remove/reorder-shape-animation`) + canvas badge overlay |
| 090445a | F4.5  | Animations: CLI + MCP + projection + tests + `10-with-anim.pptx` fixture                  |

Final test/build totals at the close of this follow-up shift:

- `@officeai/pptx`: **94 tests** in 14 files (parser, serializer, agent,
  P0/P1 commands, table commands, chart commands, animation commands,
  renderer layout/SVG, theme resolver, headless invariant).
- `@officeai/agent`: **58 tests** in 3 files (15 docx CLI + 18 MCP +
  25 pptx CLI).
- `@officeai/integration-tests`: **38 tests** in 6 files (license scan,
  OOXML schema validation, docx agent roundtrip, docx fixtures
  roundtrip, docx real-world roundtrip, **pptx real-world roundtrip**).
- `@officeai/web`: `pnpm build` succeeds; `/pptx-editor` ships at
  ~739 B page bundle / 119 kB First Load JS; **6/6 PPTX Playwright e2e
  tests pass** (`apps/web/e2e/pptx-editor.spec.ts`: route mount, Add
  slide, Text box, Bold toggle, zoom slider, agent panel via LLM bridge
  offline fallback).
- Architecture check: green (`scripts/check-architecture.mjs`).
- Browser smoke (cursor-ide-browser MCP): `/pptx-editor` mounts the
  toolbar (Open / Export / Add slide / Duplicate / Delete / Text box /
  Bold / Italic / Underline / Zoom out / Zoom slider / Zoom in / 100 %
  reset), the slides sidebar (sample-deck thumbnail), the canvas, and
  the Agent panel.

Resolved deviations recorded above:

- F1.1 (LLM bridge), F1.2 (theme colors), F1.3 (Playwright e2e),
  F1.4 (real-world fixtures + integration roundtrip), F1.5 (zoom/DPI).
- F2 promotes Tables out of `OpaqueShape` into a typed `TableShape`
  with five commands, a native SVG renderer, CLI subcommands, MCP
  visibility, and a byte-roundtrip test against `06-with-table.pptx`.
- F3 promotes Charts out of `OpaqueShape` into a typed `ChartShape` +
  `ChartPart`, with three commands, a minimal native bar/line/pie/area
  SVG renderer (placeholder for unsupported variants), CLI subcommands,
  MCP visibility, and a `09-with-chart.pptx` fixture.
- F4 promotes slide transitions and simple entrance animations out of
  the opaque tail into typed `Slide.transition` + `Slide.animations`,
  with four commands, a numbered-badge canvas overlay, CLI subcommands,
  MCP visibility, and a `10-with-anim.pptx` fixture. Untouched slides
  re-emit the original `<p:timing>` blob byte-for-byte; edited slides
  rebuild it from the typed model (with the loss caveat noted under
  *Known issues*).

SmartArt remains out of scope per the explicit non-goal in F4.0.
