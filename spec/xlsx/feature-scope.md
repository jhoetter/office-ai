# XLSX — Feature Scope

> The contract for what `@officeai/xlsx` **must do** and **must not
> attempt** in this build. Every other spec doc and every agent command
> reads this file as ground truth. If something is in scope here it
> ships. If it is out of scope it is preserved as opaque blobs (never
> deleted, never edited).

Scope is taken from [`prompt.md`](../../prompt.md) lines 230–275, refined
by the analysis in [`analysis.md`](analysis.md).

## In scope

### File I/O

- Open any real-world `.xlsx` produced by Excel 2010–365 (desktop and web), LibreOffice Calc, Google Sheets export.
- Save back to `.xlsx` with **byte-preservation** for every OOXML part the editor did not touch. Modified parts are re-emitted from the in-memory model.
- Headless I/O: load from `ArrayBuffer` / `Uint8Array`, export to `ArrayBuffer` — works in Node and in the browser through the `XlsxAgent` surface.

### Cell editing

- Set / clear cell value (string, number, boolean, error, blank).
- Set / clear cell formula (formula string + cached value driven by the engine).
- Set / clear cell number format (built-in IDs and custom format strings).
- Set / clear cell font, fill, border, alignment.
- Set / clear cell hyperlink (external URL, internal sheet ref, mailto).
- Set / clear cell comment / note (author, text, body run formatting).

### Range editing

- Set range values from a 2D array (CSV-shaped paste).
- Apply format to a range (font / fill / border / alignment / number format) — handler walks the rectangle and applies per-cell.
- Merge / unmerge cells (rectangular ranges only).

### Structural editing

- Insert N rows at row R; delete N rows starting at R.
- Insert N columns at col C; delete N columns starting at C.
- Add new sheet (with optional name + insert position).
- Rename sheet.
- Set sheet tab color.
- Set freeze panes (split row, split column).
- Reorder sheets.

### Filters

- Apply / clear an auto-filter on a sheet's used range. We do not
  implement filter conditions in P0; existing filter conditions are
  preserved through round-trip.
- Sort a column ascending or descending (stable sort; preserves row associations).

### Conditional formatting

- Read existing rules of these kinds and preserve through round-trip:
  - cell-value (greater than, less than, equal, between)
  - text-contains
  - date-occurring (today, this week, last month, etc.)
  - color scales (2-color and 3-color)
  - data bars (basic — no axis customization)
- Add new rules of the cell-value kind via the agent.
- All other conditional-formatting rule types preserve through round-trip but are not authored by the agent in this scope.

### Charts

- Render existing charts (line, bar, column, pie, area) in the renderer
  via image rasterization fallback (we do **not** ship a JS chart engine
  in P0; charts come from the original XLSX as embedded images via the
  drawing parts).
- Preserve all chart parts through round-trip.
- Authoring new charts is **out of scope**.

### Defined names

- Read defined names (workbook-scope and sheet-scope) and surface them
  through the agent API.
- Preserve through round-trip.
- Adding / editing defined names is deferred (round-trip preservation only in P0).

### Formulas

- Parse and evaluate the priority list of ~150 functions per
  `prompt.md` lines 247–264 — see [`formula-engine.md`](formula-engine.md)
  for the full breakdown.
- Update `cachedValue` for affected cells when a referenced cell
  changes (forward dep graph + topological recalc).
- Detect circular references and surface them as a structured `#REF!`
  with cycle metadata in the mutation diff.
- Preserve formulas we cannot evaluate (unknown function name, unsupported
  syntax) as their original strings; emit `#NAME?` as the cached value
  and flag the cell for the renderer.

### Comments / notes

- Read existing comments and threaded comments; surface in the agent API.
- Add a new comment to a cell.
- Reply to / resolve / delete a comment (preserves the threading rels graph).

### Hyperlinks

- Read existing hyperlinks. Add / remove cell-level hyperlinks via the agent.

### Multi-sheet workbooks

- Multiple sheets, sheet rename, sheet tab color, sheet reorder, sheet visibility (show/hide).

### Renderer

- Virtualized DOM grid (only visible rows render).
- Sticky row + column headers; sheet-tab strip below; formula bar above.
- Frozen panes honored.
- Cell selection, range selection, keyboard navigation.
- Inline edit on `Enter` / typing; `Escape` cancels.
- Every interaction → command on the bus (no direct model mutation).

## Explicitly out of scope (preserve only)

These OOXML structures must round-trip byte-clean but the editor and
agent **do not edit them**.

- Pivot tables (caches, definitions, layouts) — any edit to a pivot is
  rejected with a structured error.
- Slicers and timelines.
- Sparklines (preserve; creating new ones is out of scope).
- Power Query (`xl/queryTables/`, `xl/connections.xml`, custom XML).
- Power Pivot (`xl/dataModel/`).
- VBA macros (`xl/vbaProject.bin`) — preserve, never execute.
- Goal Seek / Solver / Scenario Manager.
- External workbook references (`xl/externalLinks/`).
- Custom XML data binding (`xl/customXml/`).
- Themes other than `theme1.xml` (rare; preserve).
- Sheet protection algorithms (preserve hash; do not ourselves enforce).
- Print settings, page setup, headers/footers (sheet-level), printer drivers.

## P0 vs P1 within in-scope

**P0 (this phase, ships in `make verify` green):**

- File I/O
- Cell editing (value, formula, format, font/fill/border/alignment, comment, hyperlink)
- Range editing (set values, apply format, merge/unmerge)
- Structural editing (insert/delete row, insert/delete column, add sheet, rename sheet)
- Formulas: full lexer/parser/AST/evaluator/dep-graph + the math/stats, logic, info, lookup, text categories from the priority list
- Comments: add/reply/resolve/delete
- Hyperlinks: add/remove
- Multi-sheet basics
- Renderer with virtualization, formula bar, sheet tabs

**P1 (deferred but tracked in `docs/build-log/xlsx.md`):**

- Filters and sorting
- Conditional-format rule authoring (read/preserve in P0)
- Chart authoring (render from image fallback in P0)
- Defined-name authoring (read/preserve in P0)
- Date/time + finance + array formula categories (parsed and registered as `#NAME?` until shipped)
- Sheet tab color, freeze panes UI (model + serializer in P0; renderer in P1)

The P0/P1 split is mirrored in [`acceptance-criteria.md`](acceptance-criteria.md): only P0 items gate "XLSX is shipped".
