# XLSX phase summary (P0 → P13)

> Status: 2026-04-18. Thirteen phases shipped against
> [`prompt.md`](../prompt.md) lines 396–410 and the spec corpus
> in [`spec/xlsx/`](../spec/xlsx). This file is the high-level
> overview; [`docs/build-log/xlsx.md`](build-log/xlsx.md) is the
> chronological per-batch narrative with decisions, deviations
> and caveats.

## TL;DR

The XLSX track now has:

- A typed in-memory model that round-trips real-world Microsoft
  Excel and SheetJS-generated workbooks **byte-identically** for
  every untouched part. Touched parts re-emit through a typed
  serializer driven by a `partHashes` baseline (gated in CI).
- A full headless `XlsxAgent` mirroring `DocxAgent` with all
  **13 P0 commands** from the prompt plus 2 Phase-11 sizing
  additions (`xlsx:set-column-width`, `xlsx:set-row-height`).
  Every command carries a typed payload, a `precheck`, an OOXML
  impact statement and a property-tested inverse.
- A synchronous formula engine (lexer → AST → evaluator →
  dependency graph → recalc orchestrator) with **89 P0 functions**
  across math / logic / info / lookup / text categories, 624 tests
  pinning the spec.
- An Excel-flavoured browser editor at `/xlsx-editor` on a
  hand-rolled virtualized grid: open `.xlsx` from disk, multi-cell
  selection (anchor/focus, drag-extend, shift-click), type-to-edit,
  click-to-insert-ref, formula autocomplete (Tab to accept), rich
  styling toolbar (font / align / fill / number-format), merge /
  unmerge / insert / delete from the toolbar, drag-resize column
  and row headers, plus the Save round-trip.
- (P12) **Coloured formula references** — Excel-style colour
  tokens in the formula bar overlay with matching coloured
  borders on referenced cells in the grid, driven by a permissive
  `tokenizeForDisplay()` scanner that never throws on partial
  input. Refs are normalised (`A1`, `$A$1`, `Sheet2!A1` form
  three distinct keys with stable colours).
- (P12) **Excel-grade keyboard parity** — arrow nav (with
  Shift+Arrow extend and Ctrl/Cmd+Arrow jump-to-data-edge),
  Home / Ctrl+Home / Ctrl+End, Enter / Tab / Shift+Enter /
  Shift+Tab commit-and-move, F2 to focus the formula bar,
  Escape to cancel, row / column header click for whole-axis
  selection, and **Delete on a whole row / column drops the
  entire row / column** through one bus dispatch.
- An `office-agent xlsx` CLI subcommand family (`inspect`, `read`,
  `set-cell`, `set-formula`, etc.) plus an MCP server (`xlsx_*`
  tool family) sharing one transport with the DOCX tools, so the
  same agent surface is reachable from LLMs.

## Phase map

| Phase | Theme                                            | Build-log section             | Highlights                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  |
| ----- | ------------------------------------------------ | ----------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 0     | Scaffold + architecture wiring                   | `## Phase 0`                  | `@officeai/xlsx` package, monorepo wiring, command-bus reuse from `@officeai/core`.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| 1     | Four parallel clean-room analyses                | `## Phase 1`                  | SheetJS, Univer-core, Univer-formula and agent-pattern reads → `spec/xlsx/analysis*.md` synthesis.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          |
| 2     | Spec corpus                                      | `## Phase 2`                  | `agent-commands.md`, `formula-engine.md`, `document-model.md`, `ooxml-mapping.md`, `parser.md`, `serializer.md`, `renderer.md`, etc.                                                                                                                                                                                                                                                                                                                                                                                                                                                                        |
| 3     | Synthetic fixtures                               | `## Phase 3`                  | `fixtures/xlsx/01-09` covering basic grid, formulas, merges, styles, comments, large sheets, etc.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| 4     | Round-trip oracle                                | `## Phase 4`                  | Parser + serializer skeleton with byte-equality oracle and opaque-blob preservation. +12 round-trip tests.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  |
| 5     | Typed cell model + 5 / 13 P0 commands            | `## Phase 5`                  | First five P0 commands ride the bus.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        |
| 6     | Headless `XlsxAgent` + diff module               | `## Phase 6`                  | `DocxAgent` parity, integration suite, structured `DocumentDiff`.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           |
| 7a–e  | Formula engine                                   | `## Phase 7a` … `## Phase 7e` | Tokens / lexer / parser / evaluator / dependency-graph + 89 P0 functions across 5 parallel sub-agents. +475 tests.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          |
| 7f–j  | Remaining P0 command handlers                    | `## Phase 7f` … `## Phase 7j` | `set-cell-formula`, `set-cell-format`, `add-sheet`, `insert/delete-row/column`, `add-comment` — 13/13 P0 commands shipped.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  |
| 8     | CLI + MCP surface                                | `## Phase 8`                  | `office-agent xlsx` subcommands + 22 `xlsx_*` MCP tools sharing the DOCX transport.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         |
| 9     | Web surface                                      | `## Phase 9`                  | Virtualized grid, `/xlsx-editor` page, formula bar, sheet tabs, agent-snapshot subscription, first Playwright e2e.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          |
| 10    | Browser smoke + close-out                        | `## Phase 10`                 | Manual flow-through (open `/xlsx-editor` → edit A2 → set `=B3*2` and watch B4 cascade) + README refresh.                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                    |
| 11    | Excel-flavoured UX                               | `## Phase 11`                 | Open `.xlsx` from disk, multi-cell selection, type-to-edit, click-to-insert-ref, autocomplete, styling toolbar, structural buttons, drag-resize, +2 sizing commands.                                                                                                                                                                                                                                                                                                                                                                                                                                        |
| 12    | Excel parity polish: ref highlighting + keyboard | `## Phase 12`                 | Coloured ref tokens in the formula bar, matching coloured borders on referenced cells in the grid, full Excel-style keyboard nav (arrows + Shift extend + Ctrl jump-to-data-edge + Tab/Enter commit-and-move + F2/Escape), row/col header click selection + Delete-deletes-row/col. +13 e2e tests, +14 unit tests.                                                                                                                                                                                                                                                                                          |
| 13    | Clipboard, fill, undo, UX cleanup                | `## Phase 13`                 | Toolbar de-clutter; Excel-style right-click menus; `XlsxClipboardSnapshot` + `xlsx:paste-range` (relative-ref shift, transpose, paste-special modes); Cmd+C/X/V system-clipboard bridge with marching-ants overlay; headless external-clipboard parsers (Excel Desktop, Google Sheets, Numbers, CSV/TSV); `xlsx:text-to-columns` + delimiter wizard; smart fill handle (`xlsx:fill-range` + 6 series detectors); CommandBus-level Undo/Redo (Cmd+Z / Cmd+Shift+Z, toolbar buttons, agent + MCP tools, CLI `--undo`); 4 new e2e suites. +12 commands tests, +8 undo round-trips, +10 e2e tests, +1 MCP test. |

The full per-batch log lives in
[`docs/build-log/xlsx.md`](build-log/xlsx.md). Each section follows
the same shape used in the DOCX log: **what shipped → decisions →
caveats**.

## Test counts (2026-04-18)

| Package                               |   Tests |
| ------------------------------------- | ------: |
| `@officeai/xlsx` (model + engine)     |     740 |
| `@officeai/web` Playwright (xlsx e2e) |      36 |
| **XLSX-relevant total**               | **776** |

The same `@officeai/agent` (CLI + MCP) and `@officeai/integration-tests`
suites that gate DOCX also exercise the `xlsx_*` tool family; see
[`session-summary.md`](session-summary.md) for the cross-track
totals.

## What's still deferred

The current shortlist (P12 candidates fed by Phase 11 caveats and
the original `feature-scope.md` deferred list):

- **Cross-sheet formula rewrite on insert/delete** — single-sheet
  rewrite shipped in P7i; the multi-sheet ripple still pends.
- **OOXML `cols/@width` round-trip** — Phase 11 stores resized
  widths in CSS pixels; a foreign reader (Excel desktop) opening
  a workbook we resized will see the default column width until a
  follow-up writes back the character-based attribute.
- **Multi-rectangle selection (Ctrl-click areas)** — out of scope
  for P11; would require a different selection model than the
  current `{ anchor, focus }`.
- **Keyboard navigation in the body grid** — arrow-key Excel
  semantics (Enter-down, Tab-right, Ctrl-Arrow jumps) are queued
  for P12.
- **Save-loop e2e** — `open → edit → save → reopen → assert` is
  pending Playwright's `download` integration; unit serializer
  tests + manual smoke cover the path today.
- **Toolbar visual polish** — borders dropdown, font family / size
  pickers, increase/decrease decimals are spec'd in `renderer.md`
  but not on screen yet.
- **Threaded comments, VML drawing emission, defined names,
  hyperlinks, sheet reorder/delete/hide,
  `xlsx:edit-comment` / `delete-comment` / `reply-comment`
  /`resolve-comment`, rich-text comment runs** — all P12.

## Reading order for a new contributor

1. [`prompt.md`](../prompt.md) — the brief (lines 396–410 are the
   13 P0 commands).
2. [`spec/xlsx/README.md`](../spec/xlsx/README.md) — index of the
   XLSX contract.
3. [`spec/xlsx/analysis.md`](../spec/xlsx/analysis.md) — the
   clean-room synthesis decisions (what we keep / differ on /
   improve vs Univer + SheetJS).
4. [`spec/xlsx/document-model.md`](../spec/xlsx/document-model.md)
   then [`spec/xlsx/ooxml-mapping.md`](../spec/xlsx/ooxml-mapping.md).
5. [`spec/xlsx/agent-commands.md`](../spec/xlsx/agent-commands.md)
   — typed contracts for every `xlsx:*` command (incl. the Phase
   11 sizing additions appendix).
6. [`spec/xlsx/formula-engine.md`](../spec/xlsx/formula-engine.md)
   — lexer / AST / evaluator / dependency graph contract.
7. [`docs/build-log/xlsx.md`](build-log/xlsx.md) — what actually
   shipped, in order, with caveats.
8. [`docs/architecture-xlsx-deltas.md`](architecture-xlsx-deltas.md)
   — cross-cutting architectural choices that distinguish the XLSX
   product from the DOCX baseline (renderer, formula engine, style
   table, selection model, sizing, diff vocabulary, drag UX,
   parallel-agent build methodology, test pyramid).

## How to run

```bash
pnpm install
pnpm --filter @officeai/xlsx test         # 624 unit tests
pnpm --filter @officeai/web dev           # Excel editor at http://localhost:3000/xlsx-editor

# CLI
pnpm --filter @officeai/agent build
node packages/agent/dist/cli.js xlsx inspect --file fixtures/xlsx/01-basic-grid.xlsx
node packages/agent/dist/cli.js xlsx read --file fixtures/xlsx/01-basic-grid.xlsx \
  --sheet Sheet1 --range A1:D10 --format markdown
node packages/agent/dist/cli.js xlsx set-formula --file fixtures/xlsx/01-basic-grid.xlsx \
  --sheet Sheet1 --cell B5 --formula "=SUM(B1:B4)" --out updated.xlsx

# MCP server (xlsx_* + docx_* on one transport)
node packages/agent/dist/cli.js mcp

# Playwright xlsx smoke (requires the dev server on :3001)
cd apps/web && E2E_BASE_URL=http://localhost:3001 pnpm exec playwright test e2e/xlsx-*.spec.ts
```
