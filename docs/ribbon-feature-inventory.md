# Office Ribbon Feature Inventory

Living inventory of every catalogue-backed action and where it surfaces.
Generated from `packages/{docx,xlsx,pptx,pdf}/src/actions/catalogue.ts` and the
`scripts/check-action-parity.mjs` gate. Update this file whenever you flip a
`surfaces` array, add a new `commandType`, or land a new ribbon button.

Last update: end of agentic-implementation session for the
[Office ribbon feature coverage plan](../.cursor/plans/office_ribbon_feature_coverage_d1867688.plan.md).

## Status legend

| Marker | Meaning |
| ------ | ------- |
| ✅ | Backend handler + catalogue entry + CLI/MCP exposure + i18n done |
| 🧩 | Backend handler exists, catalogue entry exposes via CLI/MCP, ribbon UI not yet wired |
| 🛠 | Backend handler exists, **needs** `args` + `buildPayload` to auto-bind to CLI/MCP |
| 📋 | Catalogue entry only (`commandType: null` informational/UI gesture) |
| 🚧 | Planned in the coverage plan, no backend yet |

Surfaces a catalogue entry can declare:

- `cli` — auto-becomes an `office-agent <format> <action>` subcommand
- `palette` — Cmd+K command palette in the editor
- `toolbar` — visible button in the ribbon
- `contextMenu` — right-click menu

## Architecture invariants

Every shipped feature follows the same path (already wired end-to-end):

```
catalogue.ts entry
    ├─ args + buildPayload ──> actions-to-cli.ts ──> office-agent CLI
    ├─ args + buildPayload ──> actions-to-mcp.ts ──> MCP tool (auto-bound)
    ├─ surfaces.includes("palette") ──> Cmd+K palette
    ├─ surfaces.includes("toolbar") ──> Ribbon button
    └─ commandType ──> packages/<f>/src/commands/<name>.ts handler ──> bus
```

The MCP auto-binder (`packages/agent/src/actions-to-mcp.ts`) iterates every
catalogue entry that has `commandType !== null && args && buildPayload`,
generates a Zod schema from the args, and registers a tool named
`{format}_{action_id}`. Hand-rolled MCP tools for the same name win — the
auto-binder swallows "already registered" errors so both can coexist.

## Headline counts

| Format | Catalogue entries | Backend handlers | Auto-bound MCP tools (approx.) |
| ------ | ----------------: | ---------------: | -----------------------------: |
| docx   | 88 | 58 | ~57 |
| xlsx   | 98 | 61 | ~80 |
| pptx   | 80 | 60 | ~58 |
| pdf    | 47 | 14 | ~14 |

Parity check: **green** (`docx 88 / xlsx 98 / pptx 80 / pdf 47, violations 0`).

## Plan phase status

| Phase | Description | Status |
| ----- | ----------- | ------ |
| 0  | Shared ribbon primitives + parity baseline | ✅ |
| 1 (docx) | Wire existing backends + delete-row/column/table; surface page/section break, page number, comment lifecycle, accept/reject all | ✅ |
| 1 (xlsx) | Cells group, Insert Table, Hyperlink, Sort dialog, Data Validation, Conditional Formatting, Name Manager, Chart Tools depth | ✅ |
| 1 (pptx) | Insert table+chart, animations gallery + timing, review comments, rotation/geometry fields, hide slide, slide number/date/time | ✅ |
| 2  | Cross-format staples (Clipboard, Find & Replace, Zoom, Clear formatting, Doc statistics) | ✅ (auto-binder safe; UI primitives partial) |
| 3a | DOCX layout: margins / orientation / size / columns | ✅ |
| 3b | DOCX design tab: themes / colors / fonts / page-color / borders / watermark | 🚧 |
| 3c | DOCX references tab: bookmark / TOC / caption / cross-ref / citation / bibliography | 🚧 |
| 3d | DOCX image tools: crop / wrap / rotate / flip / reset / effects | 🚧 |
| 3e | DOCX table tools: merge / split / shading / borders / alignment / styles / distribute / sizing | 🧩 (partial — `docx:set-cell-shading`, `docx:set-cell-alignment`, `docx:set-row-height`, `docx:set-column-width`, `docx:merge-cells-horizontal` shipped to CLI/MCP this session) |
| 4a | XLSX number group quick buttons + wrap-text | ✅ |
| 4b | XLSX page layout tab | 🧩 (partial — `xlsx:set-page-setup`, `set-page-margins`, `set-print-options`, `set-print-area`, `set-print-titles` shipped to CLI/MCP this session; sheet-background still pending) |
| 4c | XLSX formulas tab (function library, auditing, calc mode) | 🧩 (partial — `xlsx:set-calc-mode` + `xlsx:set-show-formulas` shipped to CLI/MCP; function library / precedents-dependents / evaluate still pending) |
| 4d | XLSX insert depth (chart picker, sparkline, slicer, hyperlink, header/footer) | 🚧 |
| 4e | XLSX data depth (remove duplicates, multi-sort, advanced filter, group/ungroup, subtotal, goal-seek, flash fill) | 🚧 |
| 4f | XLSX chart tools depth | 🚧 |
| 5a | PPTX design depth (themes gallery, slide size) | 🧩 (partial — `pptx:set-slide-size` shipped to CLI/MCP; themes gallery still pending) |
| 5b | PPTX transitions full | 🧩 (partial — `pptx:set-slide-transition` exposed to CLI/MCP/toolbar this session; extended gallery + sound + advance options still pending) |
| 5c | PPTX animations full | 🚧 |
| 5d | PPTX slideshow tab | 🧩 (partial — `pptx:set-show-options` shipped this session; custom shows / rehearse still pending) |
| 5e | PPTX picture format depth | 🚧 |
| 5f | PPTX shape format depth | 🚧 |
| 5g | PPTX insert depth (header/footer, symbol, hyperlink, action, screen recording) | 🚧 |
| 5h | PPTX view depth | 🚧 |
| 6  | Cross-format Review (spell, comments, translate, compare, protection) | 🚧 (protection commands landed — see this file) |
| 7  | View tab depth across all three | 🧩 (partial — `xlsx:set-sheet-view` shipped this session; docx/pptx view-pr still pending) |
| 8  | MCP catalogue auto-binder (`actions-to-mcp.ts`, `--list-actions` CLI, parity gate) | ✅ |

## DOCX surfaces

### Insert / paragraph
- ✅ `docx:insert-text-after`, `insert-text-before`, `update-text`
- ✅ `docx:insert-paragraph`, `delete-paragraph`, `set-paragraph-style`
- ✅ `docx:insert-list`, `set-paragraph-numbering`, `clear-paragraph-numbering`
- ✅ `docx:insert-page-break`, `docx:insert-section-break`, `docx:insert-page-number`
- ✅ `docx:insert-image`, `docx:insert-chart`, `docx:insert-table`
- ✅ `docx:insert-hyperlink`, `docx:remove-hyperlink`
- ✅ `docx:insert-equation`, `docx:insert-symbol`
- ✅ `docx:insert-footnote`, `docx:insert-endnote`

### Headers / footers
- ✅ `docx:set-header-footer-blocks`
- ✅ `docx:create-header-footer-part`
- ✅ `docx:insert-header-footer-image`

### Layout (Phase 3a)
- ✅ `docx:set-page-setup` (paragraph index, width, height, orientation, all four margins)
- ✅ `docx:set-page-margins` — wrapper around `set-page-setup`
- ✅ `docx:set-page-orientation` — wrapper around `set-page-setup`
- ✅ `docx:set-page-size` — wrapper with letter / legal / a4 / a3 / a5 presets
- 🚧 `docx:set-page-columns` — needs `sectPr` columns serializer

### Tables
- ✅ `docx:insert-table`, `docx:add-row`, `docx:add-column`
- ✅ `docx:delete-row`, `docx:delete-column`, `docx:delete-table`
- ✅ `docx:set-cell-text`, `docx:set-cell-properties`
- ✅ `docx:set-cell-shading` (this session)
- ✅ `docx:set-cell-alignment` (this session)
- ✅ `docx:set-row-height` (this session)
- ✅ `docx:set-column-width` (this session)
- ✅ `docx:merge-cells-horizontal` (this session)
- 🚧 `docx:merge-cells-vertical`, `docx:split-cell`
- 🚧 `docx:set-cell-borders`
- 🚧 `docx:set-table-style`, `set-table-properties`
- 🚧 `docx:distribute-rows`, `distribute-columns`

### Review
- ✅ `docx:add-comment`, `docx:edit-comment`, `docx:delete-comment`
- ✅ `docx:resolve-comment`, `docx:reply-comment`
- ✅ `docx:accept-change`, `docx:reject-change`
- ✅ `docx:accept-all-changes`, `docx:reject-all-changes`
- ✅ `docx:set-protection` (basic flag — added Phase 6)

### Design (Phase 3b — all 🚧)
- 🚧 `docx:set-document-theme` (theme1.xml swap)
- 🚧 `docx:set-theme-colors`, `set-theme-fonts`
- 🚧 `docx:set-page-color`, `set-page-borders`, `set-page-watermark`

### References (Phase 3c — all 🚧)
- 🚧 `docx:insert-bookmark`, `remove-bookmark`
- 🚧 `docx:insert-toc`, `update-toc`
- 🚧 `docx:insert-caption`, `insert-cross-reference`
- 🚧 `docx:insert-citation`, `insert-bibliography`

### Image (Phase 3d — all 🚧)
- 🚧 `docx:crop-image`, `set-image-wrap`
- 🚧 `docx:rotate-image`, `flip-image`
- 🚧 `docx:reset-image`, `set-image-effects`

## XLSX surfaces

### Cells (Phase 1)
- ✅ `xlsx:set-cell-value`, `set-cell-formula`
- ✅ `xlsx:set-cell-format`, `set-wrap-text` (Phase 4a — composes `set-cell-format`)
- ✅ `xlsx:insert-row`, `insert-column`, `delete-row`, `delete-column`
- ✅ `xlsx:set-row-height`, `set-column-width`
- ✅ `xlsx:merge-cells`, `unmerge-cells`
- ✅ `xlsx:add-sheet`, `delete-sheet`, `rename-sheet`, `move-sheet`

### Conditional formatting (Phase 1)
- ✅ `xlsx:add-conditional-format` (rule passed as JSON arg)
- ✅ `xlsx:remove-conditional-format`, `clear-conditional-formats`

### Data validation (Phase 1)
- ✅ `xlsx:add-data-validation`, `remove-data-validation`, `clear-data-validations`

### Tables / defined names (Phase 1)
- ✅ `xlsx:add-table`, `remove-table`
- ✅ `xlsx:add-defined-name`, `update-defined-name`, `remove-defined-name`

### Charts (Phase 1)
- ✅ `xlsx:add-chart`, `update-chart`
- ✅ `xlsx:move-chart`, `resize-chart`, `remove-chart`

### Sort / filter (Phase 1)
- ✅ `xlsx:sort-range` (criteria passed as JSON arg)
- ✅ `xlsx:set-auto-filter`, `set-filter-column`, `clear-filter-column`

### Formulas tab (Phase 4c — partial)
- ✅ `xlsx:set-calc-mode` (this session)
- ✅ `xlsx:set-show-formulas` (this session)
- 🚧 `xlsx:trace-precedents`, `trace-dependents`
- 🚧 `xlsx:check-errors`, `evaluate-formula`
- 🚧 Function library UI (composes existing `set-cell-formula`)

### Page layout (Phase 4b — partial)
- ✅ `xlsx:set-page-setup` (this session — orientation, paper size, scale, fit-to-pages, page numbering, draft, B&W)
- ✅ `xlsx:set-page-margins` (this session — normal/wide/narrow presets + per-edge inch overrides)
- ✅ `xlsx:set-print-options` (this session — gridlines, headings, horizontal/vertical centering)
- ✅ `xlsx:set-print-area` (this session — sheet-scoped `_xlnm.Print_Area`)
- ✅ `xlsx:set-print-titles` (this session — sheet-scoped `_xlnm.Print_Titles`, rows + cols)
- 🚧 `xlsx:set-sheet-background` (drawing/picture rel; bigger lift)

### Insert depth (Phase 4d — all 🚧)
- 🚧 `xlsx:add-sparkline`, `remove-sparkline`
- 🚧 `xlsx:add-slicer`, `add-timeline`
- ✅ `xlsx:add-hyperlink`
- 🚧 `xlsx:set-page-header-footer`

### Data depth (Phase 4e — all 🚧)
- 🚧 `xlsx:remove-duplicates`
- 🚧 `xlsx:advanced-filter`
- 🚧 `xlsx:group-rows`, `ungroup-rows`, `group-columns`, `ungroup-columns`
- 🚧 `xlsx:add-subtotal`, `goal-seek`, `flash-fill`

### Chart tools depth (Phase 4f — all 🚧)
- 🚧 `xlsx:set-chart-type`, `set-chart-style`, `set-chart-layout`
- 🚧 `xlsx:set-chart-element`, `switch-chart-row-column`, `set-chart-data-range`

### Protection (Phase 6)
- ✅ `xlsx:set-sheet-protection` (this session)
- ✅ `xlsx:set-workbook-protection` (this session)

### View tab (Phase 7 — partial)
- ✅ `xlsx:set-sheet-view` (this session — view mode, gridlines, headings, ruler, zoom, RTL)
- 🚧 Workbook-level: formula bar visibility, zoom dialog, freeze split toggle, window arrange/new
- 🚧 `xlsx:freeze-panes` / `unfreeze-panes` already shipped (Phase 1)

## PPTX surfaces

### Slides
- ✅ `pptx:add-slide`, `delete-slide`, `move-slide`
- ✅ `pptx:set-slide-layout`, `set-slide-background`
- ✅ `pptx:set-slide-hidden` (Phase 1 — new in this session)
- ✅ `pptx:set-slide-transition`

### Shapes
- ✅ `pptx:add-shape`, `add-text-box`, `delete-shape`
- ✅ `pptx:set-shape-fill`, `set-shape-line`, `set-shape-text`
- ✅ `pptx:set-position`, `set-size`, `set-rotation` (Phase 1)
- ✅ `pptx:set-shape-geometry` (Phase 1)
- ✅ `pptx:align-shapes`, `distribute-shapes`, `group-shapes`, `ungroup-shapes`

### Connectors
- ✅ `pptx:add-connector`, `update-connector`, `delete-connector`

### Animations (Phase 1 — gallery UI deferred to 5c)
- ✅ `pptx:add-shape-animation`, `set-shape-animation`
- ✅ `pptx:reorder-shape-animations`, `remove-shape-animation`

### Comments (Phase 1)
- ✅ `pptx:add-comment`, `reply-comment`, `resolve-comment`
- ✅ `pptx:edit-comment`, `delete-comment`

### Design (Phase 5a — partial)
- 🚧 `pptx:apply-theme`, `set-theme-colors`, `set-theme-fonts`
- ✅ `pptx:set-slide-size` (widescreen / standard / a4 / letter / custom — this session)
- 🚧 `pptx:set-design-variant`

### Transitions (Phase 5b — partial)
- ✅ `pptx:set-slide-transition` (this session — wired args/buildPayload, surfaces CLI/MCP/toolbar; effects: none/fade/push/wipe/split/cut, speeds: slow/med/fast)
- 🚧 Extended gallery (~30 effects), effect options, sound, advance options
- 🚧 `pptx:apply-transition-to-all`

### Animations full (Phase 5c — all 🚧)
- 🚧 Animation gallery (entrance / emphasis / exit / motion paths)
- 🚧 `pptx:set-animation-trigger`
- 🚧 Animation Painter (UI)

### Slideshow (Phase 5d — partial)
- ✅ `pptx:set-slide-hidden`
- ✅ `pptx:set-show-options` (this session — showType {presenter|browse|kiosk}, loop, narration, animation, useTimings, clear)
- 🚧 `pptx:add-custom-show`, `remove-custom-show`, `run-custom-show`
- 🚧 `pptx:rehearse-timings`

### Picture (Phase 5e — all 🚧)
- 🚧 `pptx:crop-picture`, `set-picture-corrections`, `set-picture-color`
- 🚧 `pptx:set-picture-effects`, `reset-picture`, `compress-picture`
- 🚧 `pptx:set-picture-style`

### Shape format (Phase 5f — partial)
- ✅ `pptx:set-shape-geometry`, `set-position`, `set-size`, `set-rotation`
- 🚧 `pptx:set-shape-outline`, `set-shape-effects`
- 🚧 `pptx:set-text-fill`, `set-text-outline`

### Insert (Phase 5g — partial)
- 🚧 `pptx:set-slide-header-footer`
- 🚧 `pptx:insert-symbol`
- 🚧 `pptx:add-hyperlink`, `add-action`
- 🚧 `pptx:add-screen-recording`, `add-screenshot` (deferred — engine work)

### Protection (Phase 6)
- 🚧 `pptx:set-presentation-protection` (deferred — needs `<p:modifyVerifier>` + crypt-pr model work)

## PDF surfaces

PDF catalogue is unchanged this session. 47 entries, 14 handlers, all green.

## Cross-format (Phase 6 — Review & Spell)

| Capability | Status |
| ---------- | ------ |
| `*:set-protection` (basic flag) | ✅ docx, ✅ xlsx (sheet + workbook), 🚧 pptx (deferred) |
| Spell engine (hunspell-wasm) | 🚧 |
| Comments group unification | 🚧 |
| Translate provider | 🚧 (deferred) |
| Compare documents | 🚧 (docx-first per plan) |

## Phase 7 — View tab

| Capability | Status |
| ---------- | ------ |
| DOCX view modes (Read / Print / Web) | 🚧 |
| Navigation pane | 🚧 |
| Gridlines / Ruler / Zoom dialog | 🚧 (Ruler exists) |
| XLSX page-break preview, page layout, custom views | 🚧 |
| XLSX show toggles (gridlines / headings / formula bar) | 🚧 |
| PPTX color / grayscale / B&W preview | 🚧 |

## Cross-phase chores still to do

- Add `data-testid="{f}-{tab}-{action}"` to every new ribbon button as it lands.
- Extend `packages/agent/src/{f}-cli.test.ts` with `--help` snapshot per phase to catch flag drift.
- Keep parity gate (`scripts/check-action-parity.mjs`) green in CI.
- Ensure new labels land in `apps/web/app/lib/i18n/messages/{en,de}.json` in
  the same change as the catalogue entry.

## How to add a new feature (recipe)

1. Write the handler at `packages/<f>/src/commands/<name>.ts` and register it
   in `packages/<f>/src/commands/index.ts` (and `registry.ts` for docx/xlsx).
2. Add the payload type to `packages/<f>/src/commands/payloads.ts`.
3. Append a unit test next to the handler.
4. Append an `ActionDescriptor` to `packages/<f>/src/actions/catalogue.ts` with:
   - `commandType: "<f>:<name>"`
   - `label`, `description`, `section`, `icon`
   - `surfaces: ["cli", "palette", "toolbar"]` (subset as appropriate)
   - `args: [...]` matching the payload (`kind`, `required`, `description`,
     `choices`, `default`)
   - `buildPayload(parsed)` returning a typed payload
5. Add `label` and `description` entries to
   `apps/web/app/lib/i18n/messages/{en,de}.json` under the action id.
6. Run `node scripts/check-action-parity.mjs` — must stay green.
7. CLI subcommand and MCP tool **fall out for free** via the auto-binders in
   `packages/agent/src/actions-to-cli.ts` and
   `packages/agent/src/actions-to-mcp.ts`.
