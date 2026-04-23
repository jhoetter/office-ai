# Office Ribbon Feature Inventory

Living inventory of every catalogue-backed action and where it surfaces.
Generated from `packages/{docx,xlsx,pptx,pdf}/src/actions/catalogue.ts` and the
`scripts/check-action-parity.mjs` gate. Update this file whenever you flip a
`surfaces` array, add a new `commandType`, or land a new ribbon button.

Last update: end of Phase 9 + Phase 9e UI parity sweep
(DOCX bookmarks/TOC/captions/x-refs/Design-tab wired to the ribbon,
XLSX `xlsx:remove-duplicates` shipped end-to-end with a Data-tab
dialog, and remaining placeholder tooltips replaced with honest
"tracked for the next milestone" copy so users no longer read a
misleading "use CLI in the meantime" hint for verbs that have no
CLI yet) — see
[Phase 9 plan](../.cursor/plans/phase_9_office_finish_ca06d725.plan.md).

## Status legend

| Marker | Meaning                                                                              |
| ------ | ------------------------------------------------------------------------------------ |
| ✅     | Backend handler + catalogue entry + CLI/MCP exposure + i18n done                     |
| 🧩     | Backend handler exists, catalogue entry exposes via CLI/MCP, ribbon UI not yet wired |
| 🛠     | Backend handler exists, **needs** `args` + `buildPayload` to auto-bind to CLI/MCP    |
| 📋     | Catalogue entry only (`commandType: null` informational/UI gesture)                  |
| 🚧     | Planned in the coverage plan, no backend yet                                         |

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

| Format | Catalogue entries | Backend handlers | UI-dispatched | Auto-bound MCP tools (approx.) |
| ------ | ----------------: | ---------------: | ------------: | -----------------------------: |
| docx   |                90 |               60 |            52 |                            ~59 |
| xlsx   |               102 |               65 |            65 |                            ~84 |
| pptx   |                80 |               60 |            44 |                            ~58 |
| pdf    |                47 |               14 |            13 |                            ~14 |

Parity check: **green** (`scripts/check-action-parity.mjs` reports
`docx 90 / xlsx 102 / pptx 80 / pdf 47`, violations 0). Phase 9e
added two DOCX entries (`docx.insert-bookmark`,
`docx.delete-bookmark`) plus one XLSX entry
(`xlsx.remove-duplicates`); each ships with i18n labels and a
backend handler so the parity gate stays clean.

## Plan phase status

| Phase    | Description                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     | Status                                                                                                                                                                                                                                                                                                                         |
| -------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| 0        | Shared ribbon primitives + parity baseline                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                      | ✅                                                                                                                                                                                                                                                                                                                             |
| 1 (docx) | Wire existing backends + delete-row/column/table; surface page/section break, page number, comment lifecycle, accept/reject all                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 | ✅                                                                                                                                                                                                                                                                                                                             |
| 1 (xlsx) | Cells group, Insert Table, Hyperlink, Sort dialog, Data Validation, Conditional Formatting, Name Manager, Chart Tools depth                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     | ✅                                                                                                                                                                                                                                                                                                                             |
| 1 (pptx) | Insert table+chart, animations gallery + timing, review comments, rotation/geometry fields, hide slide, slide number/date/time                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  | ✅                                                                                                                                                                                                                                                                                                                             |
| 2        | Cross-format staples (Clipboard, Find & Replace, Zoom, Clear formatting, Doc statistics)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        | ✅ (auto-binder safe; UI primitives partial)                                                                                                                                                                                                                                                                                   |
| 3a       | DOCX layout: margins / orientation / size / columns                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             | ✅                                                                                                                                                                                                                                                                                                                             |
| 3b       | DOCX design tab: themes / colors / fonts / page-color / borders / watermark                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     | 🧩 (Phase 9e — ribbon "Entwurf" tab is live: Designs / Seitenfarbe / Seitenränder / Wasserzeichen prompt the user and apply session-scoped CSS variables. OOXML round-trip for these visual effects is still tracked for a follow-up plan.)                                                                                    |
| 3c       | DOCX references tab: bookmark / TOC / caption / cross-ref / citation / bibliography                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                             | 🧩 (Phase 9e — `docx:insert-bookmark` + `docx:delete-bookmark` shipped end-to-end with `OpaqueInline` round-trip and a References ribbon dialog; TOC, captions and cross-references are live as composed commands that synthesise paragraphs / text inserts. Citation + bibliography still tracked.)                           |
| 3d       | DOCX image tools: crop / wrap / rotate / flip / reset / effects                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 | 🚧                                                                                                                                                                                                                                                                                                                             |
| 3e       | DOCX table tools: merge / split / shading / borders / alignment / styles / distribute / sizing                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  | 🧩 (partial — `docx:set-cell-shading`, `docx:set-cell-alignment`, `docx:set-row-height`, `docx:set-column-width`, `docx:merge-cells-horizontal` shipped to CLI/MCP this session)                                                                                                                                               |
| 4a       | XLSX number group quick buttons + wrap-text                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     | ✅                                                                                                                                                                                                                                                                                                                             |
| 4b       | XLSX page layout tab                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            | 🧩 (partial — `xlsx:set-page-setup`, `set-page-margins`, `set-print-options`, `set-print-area`, `set-print-titles` shipped to CLI/MCP this session; sheet-background still pending)                                                                                                                                            |
| 4c       | XLSX formulas tab (function library, auditing, calc mode)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       | 🧩 (partial — `xlsx:set-calc-mode` + `xlsx:set-show-formulas` shipped to CLI/MCP; function library / precedents-dependents / evaluate still pending)                                                                                                                                                                           |
| 4d       | XLSX insert depth (chart picker, sparkline, slicer, hyperlink, header/footer)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                   | 🧩 (Phase 9c — Insert tab "Mehr" group with Sparkline / Header-footer / Recommended-Charts placeholders; tooltips updated in Phase 9e to honest "next milestone" copy.)                                                                                                                                                        |
| 4e       | XLSX data depth (remove duplicates, multi-sort, advanced filter, group/ungroup, subtotal, goal-seek, flash fill)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                | 🧩 (Phase 9e — `xlsx:remove-duplicates` shipped end-to-end with `RemoveDuplicatesDialog`, key-column picker, header toggle, CLI exposure, palette runner and unit tests. Group / Ungroup / Subtotal / Goal-seek / Flash-fill / Advanced-filter still ride the placeholder pattern, now with honest "next milestone" tooltips.) |
| 4f       | XLSX chart tools depth                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          | 🚧                                                                                                                                                                                                                                                                                                                             |
| 5a       | PPTX design depth (themes gallery, slide size)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  | 🧩 (partial — `pptx:set-slide-size` shipped to CLI/MCP; themes gallery still pending)                                                                                                                                                                                                                                          |
| 5b       | PPTX transitions full                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           | 🧩 (partial — `pptx:set-slide-transition` exposed to CLI/MCP/toolbar this session; extended gallery + sound + advance options still pending)                                                                                                                                                                                   |
| 5c       | PPTX animations full                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                            | ✅ (Phase 9c — full preset gallery driven by `presetsByCategory()`, trigger picker via `pptx:set-shape-animation`, Animation Painter Copy/Paint composes `pptx:add-shape-animation`)                                                                                                                                           |
| 5d       | PPTX slideshow tab                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              | 🧩 (partial — `pptx:set-show-options` shipped this session; custom shows / rehearse still pending)                                                                                                                                                                                                                             |
| 5e       | PPTX picture format depth                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                       | 🚧                                                                                                                                                                                                                                                                                                                             |
| 5f       | PPTX shape format depth                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         | 🚧                                                                                                                                                                                                                                                                                                                             |
| 5g       | PPTX insert depth (header/footer, symbol, hyperlink, action, screen recording)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                  | 🧩 (Phase 9c — Insert tab "Symbole" group: `insert-symbol` live (UI-only via `document.execCommand("insertText")` on focused contenteditable); `Hyperlink` / `Aktion` / `Kopf-/Fußzeile` Coming-soon triggers; backends deferred)                                                                                              |
| 5h       | PPTX view depth                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 | 🚧                                                                                                                                                                                                                                                                                                                             |
| 6        | Cross-format Review (spell, comments, translate, compare, protection)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                           | 🚧 (protection commands landed — see this file)                                                                                                                                                                                                                                                                                |
| 7        | View tab depth across all three                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                 | 🧩 (partial — `xlsx:set-sheet-view` shipped this session; docx/pptx view-pr still pending)                                                                                                                                                                                                                                     |
| 8        | MCP catalogue auto-binder (`actions-to-mcp.ts`, `--list-actions` CLI, parity gate)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              | ✅                                                                                                                                                                                                                                                                                                                             |
| 9a       | "No more lying buttons" — DOCX Find/Replace, titlePg state, unsupported toast; XLSX palette fix + fx + protect toggle; PPTX hidden slides, `<p:showPr>`, format-aware Outline rail, editor canvas auto-trigger; Playwright `lying-buttons.spec.ts`                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              | ✅                                                                                                                                                                                                                                                                                                                             |
| 9b       | Cheap wins from existing backends — DOCX Tab/Shift+Tab list demote/promote + level picker, in-place footnotes panel; XLSX AutoSum splitter, % / $ / comma / inc-dec-decimal quick buttons, A↑/Z↓ sort, hide/unhide row+col, sheet-tab color (new `xlsx:set-sheet-tab-color`), Insert Function (fx) wizard; PPTX animation drag-reorder; PPTX shape-outline / effects / text-fill / text-outline ribbon shape (Coming-soon — backends deferred to follow-up plan)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                | ✅                                                                                                                                                                                                                                                                                                                             |
| 9c       | Small new features — DOCX Design / References ribbon shape, XLSX Insert / Data depth ribbon shape (all Coming-soon, backends deferred to dedicated plans matching the §3b/§3c/§4d/§4e dependencies); PPTX animation gallery (already driven by `presetsByCategory()`), set-animation-trigger (already wired via `pptx:set-shape-animation`), Animation Painter (UI-only, composes `pptx:add-shape-animation`); PPTX `insert-symbol` (UI-only via `document.execCommand("insertText")` on the focused contenteditable overlay), `set-slide-header-footer` / `add-hyperlink` / `add-action` ribbon shape (Coming-soon)                                                                                                                                                                                                                                                                                                                                                                            | ✅                                                                                                                                                                                                                                                                                                                             |
| 9d       | PDF reader parity sweep — fixed `scripts/check-action-parity.mjs` PDF `uiDirs` path; wired `pdf:add-bookmark`, `pdf:set-metadata`, `pdf:reorder-pages` (drag in thumbnail rail) through `PdfToolbar` / `PdfSidebar` / `PdfMetadataDialog`; updated EN+DE i18n for bookmark / metadata strings; this inventory refreshed                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                         | ✅                                                                                                                                                                                                                                                                                                                             |
| 9e       | UI parity sweep — DOCX References tab is no longer a row of dead "Coming-soon" buttons: `Lesezeichen` opens the new `BookmarkDialog` driven by `docx:insert-bookmark` / `docx:delete-bookmark` (round-trip via `OpaqueInline`), `TOC` / `Aktualisieren` synthesise `Heading1-9` paragraphs into a "Inhalt" block, `Beschriftung` / `Querverweis` compose `docx:insert-paragraph` + `docx:insert-text`. DOCX Design tab buttons (`Designs`, `Seitenfarbe`, `Seitenränder`, `Wasserzeichen`) now apply session-scoped CSS variables with explicit toasts about persistence. XLSX Data tab "Duplikate entfernen" is fully wired to `xlsx:remove-duplicates` (handler + dialog + CLI + palette runner + 5 unit tests). Remaining placeholder buttons keep the planned ribbon shape but now display honest "tracked for the next milestone" tooltips instead of the misleading "use CLI in the meantime" copy that confused users about what was actually available. Action-parity gate still green. | ✅                                                                                                                                                                                                                                                                                                                             |

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

47 catalogue entries, 14 backend handlers; **13 of 14 dispatched from
the UI** (the 14th is `pdf:rotate-pages` which is dispatched but uses a
different argument name mapping than the script's heuristic catches —
tracked separately, not a regression).

### Phase 9d wiring landed this session

| Catalogue id        | UI location                                                       | Notes                                                                                          |
| ------------------- | ----------------------------------------------------------------- | ---------------------------------------------------------------------------------------------- |
| `pdf.add-bookmark`  | `PdfToolbar` "Add bookmark" + `PdfSidebar` Outline tab "+" button | Prompts for a title; defaults to "Page N"; auto-flips sidebar to Outline tab                   |
| `pdf.set-metadata`  | `PdfToolbar` "Document properties" → `PdfMetadataDialog`          | Editable: title / author / subject / keywords; read-only: creator / producer / dates / version |
| `pdf.reorder-pages` | `PdfSidebar` Thumbnails tab — drag-and-drop                       | Validates `order.length === totalPages`; dispatches the new permutation                        |

### Surfaces summary

- ✅ Page navigation, rotation, deletion, print
- ✅ Highlight + sticky annotations (`pdf:add-annotation`,
  `pdf:remove-annotation`, `pdf:update-annotation`)
- ✅ Comments lifecycle (`pdf:add-comment`, `edit-comment`,
  `resolve-comment`, `reply-comment`)
- ✅ Bookmarks (`pdf:add-bookmark` — Phase 9d)
- ✅ Metadata (`pdf:set-metadata` — Phase 9d)
- ✅ Page reorder via thumbnail drag (Phase 9d)
- 🚧 Form filling, redaction, OCR — out of scope for the current
  reader; tracked for the eventual PDF editor MVP plan

## Cross-format (Phase 6 — Review & Spell)

| Capability                      | Status                                                  |
| ------------------------------- | ------------------------------------------------------- |
| `*:set-protection` (basic flag) | ✅ docx, ✅ xlsx (sheet + workbook), 🚧 pptx (deferred) |
| Spell engine (hunspell-wasm)    | 🚧                                                      |
| Comments group unification      | 🚧                                                      |
| Translate provider              | 🚧 (deferred)                                           |
| Compare documents               | 🚧 (docx-first per plan)                                |

## Phase 7 — View tab

| Capability                                             | Status            |
| ------------------------------------------------------ | ----------------- |
| DOCX view modes (Read / Print / Web)                   | 🚧                |
| Navigation pane                                        | 🚧                |
| Gridlines / Ruler / Zoom dialog                        | 🚧 (Ruler exists) |
| XLSX page-break preview, page layout, custom views     | 🚧                |
| XLSX show toggles (gridlines / headings / formula bar) | 🚧                |
| PPTX color / grayscale / B&W preview                   | 🚧                |

## UI wiring landed this session

The catalogue entries below now have **ribbon buttons + Cmd+K palette
runners + dialogs** in the web app, in addition to their pre-existing
CLI/MCP exposure. Every dialog mirrors its Office counterpart (Page
Setup, Set Up Show, Protect Document, etc.) and reads current state
from the snapshot's opaque XML so toggles reflect document truth.

| Format | Catalogue id                   | Ribbon location                                  | Dialog (if any)                            |
| ------ | ------------------------------ | ------------------------------------------------ | ------------------------------------------ |
| xlsx   | `xlsx.set-page-setup`          | Seitenlayout → Seite einrichten                  | `XlsxPageSetupDialog`                      |
| xlsx   | `xlsx.set-page-margins`        | Seitenlayout → Ränder splitter                   | shares `XlsxPageSetupDialog` (Margins tab) |
| xlsx   | `xlsx.set-print-options`       | Seitenlayout → Blattoptionen toggles             | `XlsxPageSetupDialog` (Sheet tab)          |
| xlsx   | `xlsx.set-print-area`          | Seitenlayout → Drucken splitter                  | —                                          |
| xlsx   | `xlsx.set-print-titles`        | Seitenlayout → Drucken                           | `XlsxPageSetupDialog` (Sheet tab)          |
| xlsx   | `xlsx.set-calc-mode`           | Formeln → Berechnung splitter                    | —                                          |
| xlsx   | `xlsx.set-show-formulas`       | Formeln → Formelüberwachung toggle               | —                                          |
| xlsx   | `xlsx.set-sheet-protection`    | Überprüfen → Schützen                            | `ProtectSheetDialog`                       |
| xlsx   | `xlsx.set-workbook-protection` | Überprüfen → Schützen                            | `ProtectWorkbookDialog`                    |
| xlsx   | `xlsx.set-sheet-view`          | Ansicht → Anzeigen / Zoom                        | `ZoomDialog`                               |
| pptx   | `pptx.set-slide-size`          | Entwurf → Anpassen splitter                      | `SlideSizeDialog`                          |
| pptx   | `pptx.set-show-options`        | Bildschirmpräsentation → Einrichten              | `SetUpShowDialog`                          |
| pptx   | `pptx.set-slide-hidden`        | Bildschirmpräsentation → Folie ausblenden toggle | —                                          |
| pptx   | `pptx.set-slide-transition`    | Übergänge → Effekt-Galerie                       | —                                          |
| docx   | `docx.set-protection`          | Überprüfen → Schützen                            | `ProtectDocumentDialog`                    |
| docx   | `docx.set-cell-shading`        | Tabellentools → Entwurf splitter                 | inline preset menu                         |
| docx   | `docx.set-cell-alignment`      | Tabellentools → Ausrichtung trio                 | —                                          |
| docx   | `docx.set-row-height`          | Tabellentools → Größe (cm input)                 | —                                          |
| docx   | `docx.set-column-width`        | Tabellentools → Größe (cm input)                 | —                                          |
| docx   | `docx.merge-cells-horizontal`  | Tabellentools → Verbinden (to-column input)      | —                                          |

DOCX table-cell commands target the `{row, column}` pair surfaced by
the contextual tab's "Zielzelle" picker — DOCX tables render today as
ProseMirror node atoms, so cell-level caret editing is not yet
available; the explicit row/column inputs are the same fallback Word's
Table Properties dialog uses when invoked from the menu rather than a
cell selection.

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
