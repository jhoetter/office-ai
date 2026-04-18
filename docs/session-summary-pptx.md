# Session Summary — PPTX Phase (Slides)

> Date: 2026-04-18 — 2026-04-19 ("night shift" + follow-up shift)
> Scope (chosen at session start): full PPTX phase (Analyze → Spec → Build → Validate), then a follow-up shift covering five P0/P1 polish items, three typed-model phases (Tables, Charts, Animations), and final validation.
> Sister documents: [`docs/session-summary.md`](./session-summary.md) (DOCX), [`docs/build-log/pptx.md`](./build-log/pptx.md) (live decision log for slides), [`docs/pptx-architecture-notes.md`](./pptx-architecture-notes.md) (where slides diverge from DOCX, and why), [`spec/pptx/`](../spec/pptx/) (authoritative specs).

## TL;DR

A working, AI-native **PowerPoint (.pptx) editor** exists end-to-end and ships on branch `feat/pptx-night-shift`:

- **`@officeai/pptx`** — parser, in-memory model, dirty-flag-driven byte-preserving serializer, 18 typed command handlers (10 P0/P1 + 5 tables + 3 charts + 4 animations), a hybrid renderer (pure layout + pure SVG + React canvas), a headless `PptxAgent`, and 94 unit/integration tests across 14 files.
- **`@officeai/agent`** — `office-agent pptx …` CLI subcommand group + 7 PPTX MCP tools (`pptx_load`, `pptx_save`, `pptx_inspect`, `pptx_get_text`, `pptx_search`, `pptx_apply_command`, `pptx_diff`). Animations + tables + charts are surfaced in both the JSON projection and the markdown projection.
- **`apps/web`** — `/pptx-editor` Next.js route with toolbar (Open / Export / Add slide / Duplicate / Delete / Text box / B / I / U / zoom out / range slider / zoom in / 100 % reset), slides sidebar, hybrid SVG+HTML canvas, and an Agent panel that routes through `/api/llm` (with an in-process intent-parser as the offline fallback).
- **`tests/`** — pptx real-world roundtrip suite (`tests/roundtrip/pptx/real-world-roundtrip.test.ts`) asserting ≥95 % byte identity on pure roundtrip and edit-isolation on a single text-shape edit.
- **6/6 PPTX Playwright e2e tests** pass (`apps/web/e2e/pptx-editor.spec.ts`) covering route mount, Add slide, Text box, Bold toggle, zoom slider, and the agent panel via the LLM bridge offline fallback.

## What shipped (slides-only)

| Artifact                  | Where                                                                                                                                                                | What it does                                                                                                                                                                                                                                                                                                                                                                                                                         |
| ------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Specs                     | `spec/pptx/{README,analysis,feature-scope,document-model,ooxml-mapping,parser,serializer,renderer,agent-commands,edge-cases,acceptance-criteria}.md`                 | Authoritative contract for slides — feature scope (P0/P1), document model, OOXML mapping, parser/serializer rules, hybrid renderer, agent commands, edge cases, and acceptance criteria. Mirrors `spec/docx/`.                                                                                                                                                                                                                       |
| Analysis                  | `spec/pptx/analysis.md`                                                                                                                                              | Clean-room study vs. `pptxgenjs`, `python-pptx`, `Aspose.Slides`, `RevealJS`, etc. — concepts only, no code copied.                                                                                                                                                                                                                                                                                                                  |
| PPTX model                | `packages/pptx/src/model/types.ts`                                                                                                                                   | `Slide`, `TextShape`, `PictureShape`, `TableShape`, `ChartShape`, `OpaqueShape`, `SlideTransition`, `EntranceAnimation`, `OpaqueXml`, dirty flags, `PptxSnapshot`.                                                                                                                                                                                                                                                                   |
| PPTX parser               | `packages/pptx/src/parser/`                                                                                                                                          | Reads `ppt/presentation.xml`, `ppt/slides/slideN.xml`, `ppt/charts/chartN.xml`, `ppt/theme/theme1.xml`. Captures unknown elements as `OpaqueShape` / `OpaqueXml`. Stable shape `NodeId`s across reloads.                                                                                                                                                                                                                             |
| PPTX serializer           | `packages/pptx/src/serializer/`                                                                                                                                      | Dirty-flag-driven; untouched parts re-emitted byte-for-byte from the `OoxmlContainer` cache. Rebuilds typed slide bodies, chart parts, and `<p:timing>` only when their backing model is dirty.                                                                                                                                                                                                                                      |
| 10 P0/P1 command handlers | `packages/pptx/src/commands/{add-slide,delete-slide,duplicate-slide,move-slide,set-text,set-shape-position,set-shape-size,format-text,insert-image,add-text-box}.ts` | The original P0/P1 set from `prompt.md` lines 414–425, plus `format-text` from line 420.                                                                                                                                                                                                                                                                                                                                             |
| 5 table commands          | `packages/pptx/src/commands/table-commands.ts`                                                                                                                       | `set-table-cell-text`, `add-table-row`, `delete-table-row`, `add-table-column`, `delete-table-column`. Promoted Tables out of `OpaqueShape` into typed `TableShape`.                                                                                                                                                                                                                                                                 |
| 3 chart commands          | `packages/pptx/src/commands/chart-commands.ts`                                                                                                                       | `set-chart-title`, `set-chart-data`, `set-chart-type` (bar / line / pie / area / placeholder for unsupported variants).                                                                                                                                                                                                                                                                                                              |
| 4 animation commands      | `packages/pptx/src/commands/animation-commands.ts`                                                                                                                   | `set-slide-transition`, `add-shape-animation`, `remove-shape-animation`, `reorder-shape-animations`. Untouched slides re-emit `<p:timing>` byte-for-byte; edited slides rebuild it from the typed model.                                                                                                                                                                                                                             |
| PptxAgent                 | `packages/pptx/src/agent/`                                                                                                                                           | Headless: `getSnapshot`, `toMarkdown`, `getSlide`, `search`, `applyCommand(s)`, `approve/rejectMutation`, `rollback`, `exportFile`, `subscribe`. Markdown projection lists transitions + animations per slide.                                                                                                                                                                                                                       |
| Hybrid renderer           | `packages/pptx/src/renderer/{layout,svg,react}/`                                                                                                                     | Pure layout fns (EMU → CSS), pure SVG factory (`shapes.ts`, `tableToSvg`, native bar/line/pie/area chart SVG, theme-color resolver), React `SlideCanvas` + `SlideThumbnail` with HTML overlay for `contenteditable` text. Animation badges (numbered yellow circles) overlay shapes that have entrance animations.                                                                                                                   |
| office-agent CLI          | `packages/agent/src/pptx-cli.ts`                                                                                                                                     | `office-agent pptx {inspect, read, search, set-text, set-shape-position, set-shape-size, add-slide, delete-slide, duplicate-slide, move-slide, format-text, insert-image, add-text-box, set-table-cell-text, add-table-row, delete-table-row, add-table-column, delete-table-column, set-chart-title, set-chart-data, set-chart-type, set-slide-transition, add-shape-animation, remove-shape-animation, reorder-shape-animations}`. |
| MCP tools                 | `packages/agent/src/mcp.ts`                                                                                                                                          | `pptx_load`, `pptx_save`, `pptx_inspect`, `pptx_get_text`, `pptx_search`, `pptx_apply_command`, `pptx_diff`. JSON + markdown projections include slide transitions, entrance animations, table cells, chart data.                                                                                                                                                                                                                    |
| Web app                   | `apps/web/app/pptx-editor/`                                                                                                                                          | `/pptx-editor` route — toolbar, slides sidebar, hybrid SVG+HTML canvas, Agent panel. Zoom slider clamps to [0.25 ×, 3 ×], 100 % reset button. Agent panel routes through `/api/llm` with `format: "pptx"`; in-process intent parser is the offline fallback.                                                                                                                                                                         |
| Synthetic fixtures        | `fixtures/pptx/synthetic/01-blank.pptx … 10-with-anim.pptx` + `MANIFEST.md`                                                                                          | 10 synthetic .pptx files (blank, title-only, title+content, multi-shape, with-image, with-table, multi-slide, large-deck, with-chart, with-anim) generated by `pnpm fixtures:pptx`.                                                                                                                                                                                                                                                  |
| Real-world fixtures       | `fixtures/pptx/real/01-styled-deck.pptx … 03-large-real-deck.pptx`                                                                                                   | 3 third-party-emitter (`pptxgenjs`) decks: styled deck w/ hyperlinks + notes, mixed-media w/ table + image, 25-slide deck. Direct PowerPoint / Google Slides / Keynote exports remain reserved slots until license-clean originals are collected.                                                                                                                                                                                    |
| Real-world roundtrip test | `tests/roundtrip/pptx/real-world-roundtrip.test.ts`                                                                                                                  | Asserts ≥95 % byte-identity on pure roundtrip and edit-isolation on a single text-shape edit.                                                                                                                                                                                                                                                                                                                                        |
| Playwright e2e            | `apps/web/e2e/pptx-editor.spec.ts`                                                                                                                                   | 6 specs: route mount + sample-deck thumbnail, Add slide enables Delete, Text box adds a shape, Bold toggle survives, zoom slider rescales, Agent panel "add a slide" via LLM bridge offline fallback.                                                                                                                                                                                                                                |
| Build log                 | `docs/build-log/pptx.md`                                                                                                                                             | Live record of every decision, deviation, deferral, phase summary, and known issue. Two session summaries inside (night shift + follow-up shift).                                                                                                                                                                                                                                                                                    |

### Test totals (slides-only)

| Package                        | Files  | Tests   | Notes                                                                                                          |
| ------------------------------ | ------ | ------- | -------------------------------------------------------------------------------------------------------------- |
| `@officeai/pptx`               | 14     | 94      | parser, serializer, agent, headless invariant, P0/P1 cmds, table cmds, chart cmds, animation cmds, layout, SVG |
| `@officeai/agent` (pptx slice) | 2      | 25 + 7  | 25 pptx-CLI tests + 7 PPTX MCP tests (out of 18 total MCP tests)                                               |
| `@officeai/integration-tests`  | 1      | 6       | `tests/roundtrip/pptx/real-world-roundtrip.test.ts`                                                            |
| `apps/web` (pptx slice)        | 1      | 6       | `apps/web/e2e/pptx-editor.spec.ts` (Playwright)                                                                |
| **PPTX total**                 | **18** | **138** | All passing.                                                                                                   |

`pnpm --filter @officeai/web build` succeeds; `/pptx-editor` ships at ~739 B page bundle / 119 kB First Load JS. Architecture check (`scripts/check-architecture.mjs`) is green.

## Phase / commit map

End-to-end PPTX work landed in **27 commits** on `feat/pptx-night-shift`:

| Commit  | Phase     | Summary                                                                                                      |
| ------- | --------- | ------------------------------------------------------------------------------------------------------------ |
| c62fe64 | P0        | `spec/pptx/analysis.md` — clean-room study vs. reference repos                                               |
| 96befa7 | P1        | Full `spec/pptx/*.md` mirror of `spec/docx/`                                                                 |
| 92afd98 | P2        | Synthetic fixtures + generator + manifest                                                                    |
| 645eba2 | P3        | Parser + serializer (byte-preserving, dirty-flag-driven)                                                     |
| d28a3a7 | P4        | `PptxAgent` skeleton + headless invariant test                                                               |
| 1a6c2ec | P5        | P0 commands (`add/delete/duplicate/move-slide`, `set-text/-position/-size`)                                  |
| d4e8880 | P6        | P1 commands (`format-text`, `insert-image` w/ SHA-256 dedup, `add-text-box`)                                 |
| 6f10c3f | P7        | Hybrid renderer (`renderer/{layout,svg,react}`)                                                              |
| 4b3ad13 | P8        | `/pptx-editor` Next.js route — toolbar, sidebar, canvas, agent panel, sample deck                            |
| 18140e6 | P9        | `office-agent pptx …` CLI subcommand group + 7 PPTX MCP tools                                                |
| 5cb24b4 | P10       | Build log + session summary                                                                                  |
| 6f25a09 | F1.1      | `/pptx-editor` agent panel routes through `/api/llm` (`format: "pptx"`, allow-list)                          |
| 7fa2591 | F1.2      | Theme-color resolution: parse `theme1.xml` → `themeDefault`; renderer resolves `schemeClr`                   |
| 773222e | F1.3      | Playwright smoke for `/pptx-editor`                                                                          |
| 6cf9098 | F1.3      | Ignore Playwright artifacts                                                                                  |
| 06e66cf | F1.4      | Real-world fixtures + integration roundtrip test                                                             |
| 8fa9013 | F1.5      | Renderer DPI/zoom hook + toolbar zoom slider                                                                 |
| e404485 | F2.1–F2.4 | Typed Tables: `TableShape` parser + dirty serializer + 5 commands + `tableToSvg`                             |
| c494e6b | F2.5      | Tables: CLI + MCP + projection                                                                               |
| 95f278b | F3.0+F4.0 | Spec amendments for Charts + Animations (with explicit non-goals)                                            |
| 795d1aa | F3.1+F3.2 | Typed `ChartShape` + `ChartPart` parser + dirty `ppt/charts/chart{N}.xml` serializer                         |
| 28588b1 | F3.3+F3.4 | Chart commands + native bar/line/pie/area SVG renderer                                                       |
| 48e411d | F3.5      | Charts: CLI + MCP + projection + tests + `09-with-chart.pptx` fixture                                        |
| f585f0d | F4.1+F4.2 | Typed `SlideTransition` + entrance `EntranceAnimation` parser/serializer (model-driven `<p:timing>` rebuild) |
| b55f28e | F4.3+F4.4 | Animation commands + canvas badge overlay                                                                    |
| 090445a | F4.5      | Animations: CLI + MCP + projection + tests + `10-with-anim.pptx` fixture                                     |
| b84a855 | F5        | Build-log session summary for the follow-up shift                                                            |

## What was deliberately deferred

- **SmartArt (`dgm:*`)** — explicit non-goal in F4.0. Stays as `OpaqueShape` (and re-emits byte-for-byte). Promoting it would need a typed diagram-layout model.
- **Exit / motion-path / emphasis animations** — only entrance animations are typed in F4. If a slide's typed animations are _edited_, any unmodeled `<p:timing>` children on that slide are dropped. Slides whose animations are _not_ edited continue to round-trip the original `<p:timing>` byte-for-byte.
- **Slide masters / slide layouts as typed model** — masters/layouts live as opaque parts; the renderer reads layout placeholders only via the slide's own opaque blobs.
- **Speaker notes** — read into `Slide.notes` but no dedicated UI yet.
- **Real PowerPoint / Google Slides / Keynote exports** in `fixtures/pptx/real/` — three slots reserved in the manifest, awaiting license-clean originals.
- **Live LLM provider for `/pptx-editor`** — wired (F1.1) through `/api/llm`, but only the in-process intent parser ships out-of-the-box; OpenAI is opt-in via `OPENAI_API_KEY`.

## Things that were harder than expected

1. **Stable shape `NodeId`s across edit cycles.** Adding/removing animations via CLI commands originally caused shape `NodeId`s to shift on the next reload, because `parseSlide` was parsing `<p:transition>` and `<p:timing>` _before_ shapes — and any change in those elements perturbed the deterministic ID minter for the shapes that came after. Fixed by reordering `parseSlide` so all shapes are parsed first, then the transition/timing tail. Documented in `docs/build-log/pptx.md` under the F4 entries.
2. **Animation roundtrip strategy.** Splicing changes into the raw `<p:timing>` XML while preserving unmodeled children (sequences, exit effects, emphasis) was excessively complex. Settled on a model-driven rebuild: untouched → re-emit `timingTailRaw` byte-for-byte; touched → drop `timingTailRaw`, rebuild a minimal `<p:timing>` from the typed `Slide.animations` list. Explicit known-issue in the build log.
3. **EMU + DPI + zoom.** SVG `viewBox` is in EMU, but the HTML overlay for `contenteditable` text needs CSS pixels at the current zoom × DPI. Factored into a single `emuToPx(emu, dpi)` helper used by both the SVG factory and the React overlay; the toolbar's zoom slider just multiplies a single `transform: scale(zoom)` on the canvas wrapper, while the overlay's font sizing tracks DPI separately so glyph rendering stays crisp at any zoom.
4. **Theme color resolution without breaking byte-roundtrip.** The opaque `<a:solidFill> > <a:schemeClr>` capture had to stay opaque so the serializer could re-emit it verbatim. The renderer threads `PptxPresentation.themeDefault` through `SvgRenderCtx.theme` and resolves `<a:schemeClr>` references _only at paint time_, never mutating the model.
5. **Charts as a separate OOXML part.** Unlike tables (which live inside `slideN.xml`), charts live in their own `ppt/charts/chartN.xml` referenced by a `<c:chart r:id="…">`. The parser had to load the chart part on-demand and the serializer's dirty-flag tracking had to span _both_ the slide that referenced the chart _and_ the chart part itself. Solved with a separate `ChartPart` entry in `PptxSnapshot.chartParts` keyed by the relationship id.
6. **Two parallel agents on the same monorepo.** A second agent was concurrently shipping DOCX hyperlinks/lists and XLSX page-zones/sizing on `main`. We ran the slides work in a worktree (`~/repos/office-ai-pptx-worktree`) on a dedicated branch (`feat/pptx-night-shift`), and explicitly left all unrelated modifications to the main checkout alone. One Playwright file from the parallel agent (`apps/web/e2e/tracked-changes.spec.ts`) currently fails with `require is not defined` — that's their cleanup, not ours; the PPTX subset (`pnpm --filter @officeai/web e2e -- pptx-editor`) is 6/6 green.

## How to run the slides editor

```bash
# install once (workspace root)
pnpm install

# typecheck the slides packages
pnpm --filter @officeai/pptx --filter @officeai/agent typecheck

# run all slides tests (94 + 25 + 6 + 6 = 131)
pnpm --filter @officeai/pptx test
pnpm --filter @officeai/agent test
pnpm --filter @officeai/integration-tests test
pnpm --filter @officeai/web e2e -- pptx-editor

# regenerate synthetic + real-world PPTX fixtures
pnpm fixtures:pptx
node scripts/generate-real-pptx-fixtures.mjs

# run the web editor (defaults to :3000; set PORT to use another port)
pnpm --filter @officeai/web dev
# → open http://localhost:3000/pptx-editor

# CLI
pnpm --filter @officeai/agent build
node packages/agent/dist/cli.js pptx inspect -i fixtures/pptx/synthetic/04-multi-shape.pptx
node packages/agent/dist/cli.js pptx read -i fixtures/pptx/synthetic/10-with-anim.pptx --format json --slide 1
node packages/agent/dist/cli.js pptx add-slide -i in.pptx -o out.pptx
node packages/agent/dist/cli.js pptx set-table-cell-text -i in.pptx -o out.pptx --slide 1 --shape <NodeId> --row 0 --col 0 --text "Hello"
node packages/agent/dist/cli.js pptx set-slide-transition -i in.pptx -o out.pptx --slide 2 --kind fade --speed med

# MCP server (for Cursor / Claude Desktop / etc.)
node packages/agent/dist/mcp.js
# tools: pptx_load, pptx_save, pptx_inspect, pptx_get_text, pptx_search, pptx_apply_command, pptx_diff
```

## Suggested next session

1. **PowerPoint / Google Slides / Keynote real-world fixtures** — collect three license-clean originals to fill the reserved slots in `fixtures/pptx/MANIFEST.md`, then promote the existing real-world roundtrip test to also assert byte-roundtrip on those three. This is the highest-confidence way to catch OOXML quirks neither `pptxgenjs` nor our synthetic fixtures expose.
2. **SmartArt typed model (P2)** — promote `dgm:*` out of `OpaqueShape`. Touch `spec/pptx/{document-model,ooxml-mapping,agent-commands}.md` first; SmartArt's diagram-layout-with-data-binding triple is the trickiest of the four `graphicFrame` payloads.
3. **Exit / motion-path / emphasis animations** — close the F4 known-issue. Either widen `EntranceAnimation` into a discriminated `Animation` union, or replace the model-driven `<p:timing>` rebuild with a splice strategy that preserves unmodeled siblings.
4. **Slide masters / layouts as typed model** — currently opaque, which means we can't propagate "change the title font on every slide" through a typed command yet. Touch `spec/pptx/document-model.md`, then add `MasterPart` + `LayoutPart` with their own dirty flags.
5. **Speaker notes UI** — model is already there (`Slide.notes`); the editor just needs a notes pane below the canvas.
6. **Live LLM provider polish for `/pptx-editor`** — F1.1 already routes through `/api/llm`; the next polish is per-command few-shot examples to keep the LLM honest about `(slideIndex, NodeId)` addressing.
