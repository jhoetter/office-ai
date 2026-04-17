# XLSX — Renderer

> Browser-side projection of `XlsxSnapshot` into a virtualized DOM grid.
> Every interaction — keystroke, click, drag, scroll, paste — funnels
> back through the **command bus**. The renderer is a skin; the agent
> owns the model.

This document mirrors the SHAPE of [`spec/docx/renderer.md`](../docx/renderer.md):
schema, model→DOM, DOM→commands, single-funnel discipline, pending
mutation visualization, headless guarantees.

References:

- [`spec/xlsx/feature-scope.md`](feature-scope.md) §"Renderer".
- [`spec/xlsx/edge-cases.md`](edge-cases.md) §"Renderer" (EC-V1..V4).
- [`spec/xlsx/document-model.md`](document-model.md) — `XlsxSnapshot`,
  `Sheet`, `Cell`.
- [`spec/xlsx/agent-commands.md`](agent-commands.md) — the command set
  the renderer dispatches into.
- [`spec/xlsx/analysis.md`](analysis.md) §1.2 row "Renderer" — the
  decision to use DOM (not canvas).

---

## 1. Goal

Given an `XlsxAgent` (which holds a working `XlsxSnapshot`), render an
interactive spreadsheet grid that:

1. **Renders any in-scope feature** the model supports (cells, formulas,
   merges, formats, comments, hyperlinks, conditional formatting,
   freeze panes, multi-sheet, sheet tab colors, hidden rows/cols).
2. **Funnels every edit through the command bus.** No PM-style "direct
   transaction" path exists; the renderer never mutates the snapshot.
3. **Scrolls smoothly through 50k rows** by virtualizing both axes —
   only cells in the visible viewport are mounted.
4. **Visualizes pending agent mutations** via cell decorations sourced
   from the agent's pending-mutation list.
5. **Stays headless-compatible** — the renderer is one subexport of
   `@officeai/xlsx`; `agent`, `formula`, and `commands` subexports must
   not import any DOM or React code (verified by a build-time guard;
   §11).

Non-goals:

- Toolbar UI (lives in `apps/web`).
- Undo/redo (the bus owns it).
- Persistence (the agent owns it).
- Authoring of charts, slicers, pivot tables (out of scope per
  `feature-scope.md`).

---

## 2. Architecture

```
┌──────────────────────────────────────────────────────────────┐
│ XlsxEditor                          (top-level mount)        │
│ ┌──────────────────────────────────────────────────────────┐ │
│ │ FormulaBar         A1   |  =SUM(B2:B10)            [✓]   │ │
│ ├──────────────────────────────────────────────────────────┤ │
│ │  ┌────┬───────────────────────────────────────────────┐  │ │
│ │  │    │ A    │ B    │ C    │ D    │ E    │ F    │ G   │  │ │
│ │  ├────┼──────┴──────┴──────┴──────┴──────┴──────┴─────┤  │ │
│ │  │ 1  │ Grid (virtualized)                            │  │ │
│ │  │ 2  │   - row header                                │  │ │
│ │  │ 3  │   - column header                             │  │ │
│ │  │ 4  │   - viewport                                  │  │ │
│ │  │ ⋮  │   - frozen quadrants                          │  │ │
│ │  └────┴───────────────────────────────────────────────┘  │ │
│ ├──────────────────────────────────────────────────────────┤ │
│ │ SheetTabStrip   [Sheet1] [Sheet2*] [+]                   │ │
│ └──────────────────────────────────────────────────────────┘ │
└──────────────────────────────────────────────────────────────┘
```

Layered components:

| Component        | Responsibility                                                                                                                                               |
| ---------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `XlsxEditor`     | Top-level mount; owns the agent and the working snapshot subscription. Single React tree.                                                                    |
| `FormulaBar`     | Active cell ref + formula input. Commits on Enter via `xlsx:set-cell-value` or `xlsx:set-cell-formula` (auto-detected). Esc reverts.                         |
| `SheetTabStrip`  | Bottom strip of sheet tabs with names + tab colors from model. Right-click menu: rename, color, delete, move. Add-sheet button.                              |
| `Grid`           | Virtualized; row header column + column header row + 4 viewport quadrants (for frozen panes).                                                                |
| `RowHeader`      | Sticky left column showing row numbers (1-based at the boundary).                                                                                            |
| `ColHeader`      | Sticky top row showing column letters (A, B, …, AA, …).                                                                                                      |
| `Viewport`       | The four panes (top-left, top-right, bottom-left, bottom-right) split by `frozenRows` × `frozenCols`. Each pane mounts only the cells in its visible window. |
| `Cell`           | Render-only; computes display string + style from the model.                                                                                                 |
| `CellEditor`     | Mounted only on the active cell when editing. Controlled input. Commits on Enter, Tab, or selection change.                                                  |
| `SelectionLayer` | Absolutely-positioned overlay drawing the selection rectangle, the active-cell border, and pending-mutation decorations.                                     |

Layout: pure DOM with Tailwind utility classes. **No canvas.** The
decision (per [`analysis.md`](analysis.md) §1.2) is to mirror DOCX —
DOM gives us accessibility, easier testing, no font-rasterization
work, no GPU edge cases, and parity with the rest of OfficeAI.

---

## 3. Virtualization

The single most important property of the renderer: **only mount cells
in the visible viewport.** With 50k rows and the default 20 px row
height, naively mounting everything would create a 1M-element DOM tree.
Instead we mount ~40 rows × ~20 cols per quadrant — a few hundred cell
elements at most, regardless of sheet size.

```
window for a quadrant Q:
  Q.firstRow = clamp(floor((scrollTop - Q.top) / rowHeight), 0, sheet.maxRow)
  Q.lastRow  = clamp(ceil((scrollTop - Q.top + Q.height) / rowHeight), 0, sheet.maxRow)
  Q.firstCol = clamp(floor((scrollLeft - Q.left) / colWidth), 0, sheet.maxCol)
  Q.lastCol  = clamp(ceil((scrollLeft - Q.left + Q.width) / colWidth), 0, sheet.maxCol)

  // Add a ~5-row/col overscan so fast scroll doesn't reveal blank gaps
  Q.firstRow -= 5; Q.lastRow += 5
  Q.firstCol -= 5; Q.lastCol += 5
```

Row heights and column widths come from the model (`Sheet.rows[r].height`,
`Sheet.cols[c].width`). When unset, defaults are **20 px** (rows) and
**64 px** (columns) — matching Excel's defaults at 100% zoom.

Since rows and columns may have **variable** heights/widths (custom row
heights, hidden rows, autofit columns), we maintain two prefix-sum
indexes per sheet:

```typescript
class SheetGeometry {
  // Cumulative pixel offset at the top of row r (0-based).
  rowOffset(r: number): number;
  // Inverse: the row whose top is at or before y.
  rowAt(y: number): number;
  colOffset(c: number): number;
  colAt(x: number): number;
}
```

Both are O(log n) via a Fenwick tree built lazily from the model on
sheet activation. Hidden rows/cols contribute zero to the prefix sum.

Scroll is **rAF-throttled**: the scroll handler updates a `scrollTop` /
`scrollLeft` ref and schedules a `requestAnimationFrame` to re-window
the cells. No more than one re-window per frame.

---

## 4. Frozen panes

The viewport is split into four quadrants:

```
┌─────────────────┬─────────────────────────────────┐
│ TopLeft         │ TopRight                        │
│ frozen rows ×   │ frozen rows ×                   │
│ frozen cols     │ scrollable cols                 │
├─────────────────┼─────────────────────────────────┤
│ BottomLeft      │ BottomRight                     │
│ scrollable      │ scrollable rows ×               │
│ rows ×          │ scrollable cols                 │
│ frozen cols     │                                 │
└─────────────────┴─────────────────────────────────┘
```

Sizes are driven by `sheet.frozenRows` and `sheet.frozenCols` (0 when
no panes are frozen — most sheets). Each quadrant scrolls
independently:

- TopLeft: doesn't scroll.
- TopRight: scrolls horizontally with BottomRight.
- BottomLeft: scrolls vertically with BottomRight.
- BottomRight: scrolls both axes.

We use a single scroll container around BottomRight and read its
`scrollLeft`/`scrollTop` to drive the others via `transform: translate(...)`.
This avoids the multi-scrollbar UX problem and matches Excel's behavior.

When `frozenRows === 0 && frozenCols === 0` (the common case) we collapse
to a single quadrant — three of the four `Viewport` instances render
nothing. The quadrant code paths are the same; the only difference is
the rectangle they cover.

---

## 5. Selection model

```typescript
interface Selection {
  readonly anchor: CellRef; // where the selection started
  readonly head: CellRef; // current focus
  readonly mode: "cell" | "range" | "row" | "column" | "all";
  readonly editing: boolean; // true → CellEditor mounted on `head`
}
```

The selection is **renderer-local state** — it does not flow through the
command bus. Selecting a cell doesn't mutate the model, so the bus has
nothing to record. Only edits do.

Keyboard navigation:

| Key                  | Action                                              |
| -------------------- | --------------------------------------------------- |
| ←/↑/→/↓              | Move active cell by 1                               |
| Shift + arrow        | Extend range by 1 in that direction                 |
| Ctrl + arrow         | Jump to next non-empty cell (or sheet edge)         |
| Ctrl + Shift + arrow | Extend range to next non-empty cell                 |
| Tab / Shift+Tab      | Move right/left (Tab + Enter keep an entry rect)    |
| Enter / Shift+Enter  | Move down/up (within entry rect if set)             |
| Ctrl+A               | Select all cells; second Ctrl+A selects whole sheet |
| Esc                  | Cancel edit / collapse range to anchor              |
| F2                   | Enter edit mode on active cell                      |
| Delete / Backspace   | `xlsx:set-cell-value` with `{ kind: "blank" }`      |
| Ctrl+C / Ctrl+X      | Copy / cut selected range to clipboard              |
| Ctrl+V               | Paste — dispatches `xlsx:set-range-values`          |

Mouse:

| Gesture                      | Action                                                                           |
| ---------------------------- | -------------------------------------------------------------------------------- |
| Click cell                   | Set selection to that cell                                                       |
| Drag from cell               | Range selection                                                                  |
| Click row header             | Select entire row                                                                |
| Click column header          | Select entire column                                                             |
| Click corner (header×header) | Select all                                                                       |
| Double-click cell            | Enter edit mode                                                                  |
| Type a printable character   | Enter edit mode with that character                                              |
| Right-click                  | Context menu (cut/copy/paste, insert/delete row/col, format, hyperlink, comment) |

Mouse and keyboard handlers are colocated in `useSelection()`.

---

## 6. Edit mode

```
EnterEditMode(cell):
  selection.editing = true
  CellEditor mounts on `cell`, populated with the cell's formula text
  if present (else its display string)

CommitEdit():
  raw = editor.value
  if raw.startsWith("=") then
    dispatch({ type: "xlsx:set-cell-formula", payload: { sheet, ref: cell, formula: raw } })
  elif raw === "" then
    dispatch({ type: "xlsx:set-cell-value", payload: { sheet, ref: cell, value: { kind: "blank" } } })
  else
    parsed = parseLiteral(raw)            // numeric? boolean? date? else string
    dispatch({ type: "xlsx:set-cell-value", payload: { sheet, ref: cell, value: parsed } })
  selection.editing = false

CancelEdit():
  selection.editing = false
  // No dispatch; the model is unchanged.
```

`parseLiteral` mirrors Excel's input parsing: pure-numeric strings
become numbers, `TRUE`/`FALSE` become booleans, dates that match the
locale's date format become date-typed numbers (with a date number
format applied), everything else is a string. Implementation lives in
`packages/xlsx/src/agent/literal-parser.ts` and is shared with the agent
API so programmatic and human inputs converge.

The `CellEditor` is a controlled `<input>` (single-line) or
`<textarea>` (after Alt+Enter inserts a newline) absolutely positioned
on top of the underlying `Cell`. It does not unmount on every
keystroke; it lives until the edit is committed or cancelled.

---

## 7. Cell rendering

```
function Cell({ row, col, sheet, snapshot }):
  cell = sheet.cells.get(`${row}:${col}`)
  if !cell: return <EmptyCell row col />

  const styleId = cell.styleId
  const xf = styleId ? snapshot.root.styles.xfFor(styleId) : DEFAULT_XF
  const numFmt = cell.numFmt ?? xf.numFmt ?? "General"

  const display = formatCell(cell.value, numFmt, xf, snapshot.root.styles)
  const cssStyle = xfToCss(xf, snapshot.root.styles)
  const cf = resolveConditionalFormatForCell(row, col, sheet, cell)
  const merged = sheet.merges.coverFor(row, col)

  if merged && (merged.top !== row || merged.left !== col):
    // Non-anchor cells in a merge render nothing
    return null

  return (
    <div class="xlsx-cell" style={merge(cssStyle, cf?.css)} data-row data-col>
      <span class="xlsx-cell-content">{display}</span>
    </div>
  )
```

### 7.1 Format display

Numbers are formatted via `XLSX.SSF.format(numFmt, cell.value.value)`
(the SheetJS SSF library; see analysis-sheetjs §6). Dates use the
serial-to-display path with the cell's number format. Errors render as
their textual code (`"#REF!"`, `"#NAME?"`, …) in red. Booleans render
as `TRUE`/`FALSE` in uppercase.

Per EC-V2: a date format applied to a string value renders the literal
string. We never coerce types at render time.

### 7.2 Alignment

Horizontal: from `xf.alignment.horizontal` (`"left" | "center" | "right" | "justify" | "fill"`).
Default by type: numbers/dates → right, booleans/errors → center, strings → left.

Vertical: from `xf.alignment.vertical` (`"top" | "center" | "bottom"`),
default `"bottom"` (Excel parity).

Indent (`xf.alignment.indent`) is applied as `padding-left: indent * 8px`.

Wrap (`xf.alignment.wrapText`) sets `white-space: pre-wrap`.

Rotation (`xf.alignment.textRotation`) becomes a CSS `transform: rotate(...)`
(only on values 0/-90/90 in P0; other angles round to nearest).

### 7.3 Borders

`xf.borderId` resolves to `styles.borders[id]`. Each side becomes a CSS
`border-{side}: {width} {style} {color}`. Mapping table:

| OOXML style | CSS style | Width |
| ----------- | --------- | ----- |
| `none`      | (omit)    | 0     |
| `thin`      | `solid`   | 1px   |
| `medium`    | `solid`   | 2px   |
| `thick`     | `solid`   | 3px   |
| `hair`      | `solid`   | 0.5px |
| `dotted`    | `dotted`  | 1px   |
| `dashed`    | `dashed`  | 1px   |
| `double`    | `double`  | 3px   |

Non-rectangular border types (diagonal up/down) are rendered with an
SVG overlay positioned at the cell. Edge cases (different border styles
on adjacent cells, "the higher-priority border wins") use CSS
border-collapse semantics — Excel's resolution is approximated, not
identical, and the delta is documented in the build log.

### 7.4 Fills

`xf.fillId` resolves to `styles.fills[id]`. Pattern types: `none` →
omit; `solid` → `background-color: fgColor`; `gray125` → a CSS
linear-gradient approximation. Theme colors resolve through
`styles.themeColor(idx, tint)` which looks up the theme's color scheme
and applies tint per Excel's HSL formula.

### 7.5 Fonts

`xf.fontId` → `styles.fonts[id]`: `font-family`, `font-size`,
`font-weight`, `font-style`, `text-decoration`, `color`. Theme fonts
(`scheme="major"|"minor"`) resolve through the workbook's theme.

### 7.6 Spillover (EC-V1)

When a cell's content is wider than its column **and** the cell to its
right is empty (no value, no formula, no comment, no hyperlink), the
content spills into the right neighbor. Implementation: render the
content with `position: absolute; left: 0; right: -overflowPx;
overflow: hidden` and let it visually overhang into adjacent empty
cells. Stops at the first non-empty neighbor.

For right-aligned cells, spill goes left; for centered cells, spill
goes both directions evenly.

### 7.7 Conditional formatting

```
resolveConditionalFormatForCell(row, col, sheet, cell):
  // Walk the rules in priority order (lowest priority # first per OOXML spec).
  // First match wins for "stop if true"; otherwise we layer fill+border+font.
  let cssLayers = {}
  for rule in sheet.conditionalFormats sorted by priority:
    if !rule.ranges.some(r => r.contains(row, col)): continue
    if !ruleMatches(rule, cell, row, col, sheet): continue
    const dxf = sheet.styles.dxfs[rule.dxfId]
    cssLayers = merge(cssLayers, dxfToCss(dxf))
    if rule.stopIfTrue: break
  return Object.keys(cssLayers).length > 0 ? { css: cssLayers } : undefined
```

`ruleMatches` is dispatched per rule kind:

- `cell-is`: compare `cell.value.value` to `rule.formulas[0]` (and
  `rule.formulas[1]` for `between`).
- `contains-text`: case-insensitive substring search.
- `date-occurring`: compare cell's date serial to the period's
  computed range (today, this week, last month, etc.).
- `color-scale`: linear interpolation between rule.stops based on the
  cell's value relative to the rule's min/max.
- `data-bar`: fill width proportional to value; render as an absolutely
  positioned `<div>` inside the cell with computed width.
- `opaque-cf`: not authored by us; not evaluated. We surface it as a
  small dotted indicator in dev mode but ignore it in prod.

We **cache** conditional-format resolutions per (sheet, row, col,
revision). Cache invalidates on any cell mutation in any of the rule's
sqref ranges, on any rule add/remove, and on sheet revision bump.

Per [`feature-scope.md`](feature-scope.md): `cell-is`, `contains-text`,
`date-occurring`, `color-scale`, `data-bar` are evaluated; everything
else round-trips but is not visualized.

---

## 8. Pending-mutation visualization

The agent exposes a `getPendingMutations()` API. Each pending mutation
carries a `{ sheet, ranges }` projection. The renderer reads this on
every snapshot tick and decorates the affected cells:

```
function PendingDecorations({ pendingMutations, sheet }):
  return pendingMutations
    .filter(m => m.sheet === sheet.id)
    .flatMap(m => m.ranges)
    .map(range => (
      <div class="xlsx-pending"
           style={absoluteRect(range, sheetGeometry)} />
    ))
```

CSS:

```css
.xlsx-pending {
  position: absolute;
  pointer-events: none;
  border: 2px solid var(--ai-violet); /* matches DOCX renderer's pending color */
  background: color-mix(in srgb, var(--ai-violet) 8%, transparent);
  border-radius: 2px;
  z-index: 5; /* above cells, below selection */
}
```

When the user approves the mutation, the bus removes it from
`pendingMutations`, the snapshot updates, and the decoration disappears.
Rejected mutations behave identically.

A hover on a pending-mutation decoration shows a popover with the
`mutation.diff` summary (e.g. "B5: 100 → 200" or "Set fill on
A1:C10"). Clicking the popover's Approve/Reject buttons dispatches
`bus:approve(mutationId)` / `bus:reject(mutationId)`.

---

## 9. The single-funnel discipline

```typescript
function XlsxEditor({ agent }: { agent: XlsxAgent }) {
  const snapshot = useSnapshot(agent);    // subscribes to agent.subscribe()
  const dispatch = useCallback(
    (cmd: Command) => agent.applyCommands([cmd]).catch(reportError),
    [agent]
  );

  return (
    <DispatchContext.Provider value={dispatch}>
      <FormulaBar snapshot={snapshot} />
      <Grid snapshot={snapshot} />
      <SheetTabStrip snapshot={snapshot} />
    </DispatchContext.Provider>
  );
}
```

Every component that produces an edit calls `dispatch(cmd)`. There is
no other mutation path. This is the same invariant DOCX's renderer
upholds — see [`spec/docx/renderer.md`](../docx/renderer.md) §"The
single-funnel plugin".

The agent's `subscribe` callback fires after every applied command (or
batch). The `useSnapshot` hook tears off the latest snapshot and
triggers a React re-render. Cells re-mount only if their row/col is in
the new viewport window AND their underlying cell (or applicable style)
has changed.

### 9.1 What goes through the bus

| User action                                  | Command(s) dispatched                                |
| -------------------------------------------- | ---------------------------------------------------- |
| Type into cell + Enter                       | `xlsx:set-cell-value` or `xlsx:set-cell-formula`     |
| Delete on selected range                     | `xlsx:set-range-values { values: [[blank, …]] }`     |
| Paste (Ctrl+V)                               | `xlsx:set-range-values { values }` (parsed CSV/TSV)  |
| Toolbar bold toggle                          | `xlsx:apply-range-format { format: { bold: true } }` |
| Drag fill handle                             | `xlsx:set-range-values { values: extrapolated }`     |
| Right-click → Insert row                     | `xlsx:insert-rows { at, count: 1 }`                  |
| Right-click → Delete row                     | `xlsx:delete-rows { at, count: 1 }`                  |
| Drag column edge                             | `xlsx:set-column-width { col, width }`               |
| Right-click sheet tab → Rename               | `xlsx:rename-sheet { sheetId, name }`                |
| Sheet tab + button                           | `xlsx:add-sheet { name?, atIndex? }`                 |
| Drag sheet tab                               | `xlsx:reorder-sheets { order }`                      |
| Right-click cell → Insert comment            | `xlsx:add-comment { sheet, ref, author, text }`      |
| Right-click cell → Insert hyperlink + dialog | `xlsx:add-hyperlink { sheet, ref, target, … }`       |
| Format menu → Conditional format → New rule  | `xlsx:add-conditional-format { sheet, range, rule }` |
| Set freeze panes via View menu               | `xlsx:set-freeze-panes { sheet, rows, cols }`        |
| Set sheet tab color via context menu         | `xlsx:set-sheet-tab-color { sheet, color }`          |

### 9.2 What doesn't go through the bus

- **Selection.** Selection is renderer-local state. The bus has nothing
  to record.
- **Scrolling.** Pure UI; no model change.
- **Active cell change.** Same — purely UI.
- **CellEditor input keystrokes** before commit. The buffered text
  lives in the editor's local state; only the final commit dispatches.
- **Tooltip/popover open/close**, **menu open/close**, **drag preview
  during a drag-selection** — all UI ephemera, none touch the model.

---

## 10. Performance

Budgets (per [`acceptance-criteria.md`](acceptance-criteria.md) G5):

| Workload                                       | Budget                |
| ---------------------------------------------- | --------------------- |
| First paint of a 50k-row × 10-col workbook     | < 200 ms after parse  |
| Scroll (continuous fling 0 → 50k rows)         | 60 fps; ≤ 16 ms/frame |
| Single cell edit roundtrip (commit → re-paint) | < 32 ms (2 frames)    |
| Switch sheets                                  | < 100 ms              |

The viewport-only render strategy keeps DOM size bounded:

```
~40 rows × ~20 cols × 4 quadrants = ~3,200 cells max
+ row header (~40) + col header (~20) + chrome ≈ 3,300 elements
```

This is comfortably within React's per-render budget. Re-renders
triggered by snapshot updates only re-mount cells whose underlying
cell + applicable style changed (computed via shallow comparison of
the precomputed `displayString` and `cssStyleHash` per cell).

Heavy operations:

- **Sheet activation** rebuilds `SheetGeometry` (Fenwick over rows +
  cols). For 50k rows this is ~5 ms — done once per sheet activation,
  cached for the session.
- **Conditional-format evaluation** caches per visible cell; cache
  invalidates on the rule's sqref range only.
- **Format string evaluation** caches per (cell.value, numFmt) tuple.
- **Theme color resolution** caches per (themeIdx, tint) tuple.

We use React 18 with concurrent rendering enabled. Scroll updates are
declared `startTransition` to prevent input lag.

---

## 11. Headless guarantee

The package's subexport map enforces the boundary:

```jsonc
// packages/xlsx/package.json
{
  "exports": {
    ".": { "types": "./dist/index.d.ts", "import": "./dist/index.js" },
    "./agent": { "types": "./dist/agent/index.d.ts", "import": "./dist/agent/index.js" },
    "./formula": { "types": "./dist/formula/index.d.ts", "import": "./dist/formula/index.js" },
    "./commands": { "types": "./dist/commands/index.d.ts", "import": "./dist/commands/index.js" },
    "./renderer": { "types": "./dist/renderer/index.d.ts", "import": "./dist/renderer/index.js" },
  },
}
```

A build-time guard test (`packages/xlsx/scripts/check-headless.mjs`)
asserts:

```
agent / formula / commands subexports must NOT transitively import:
  - "react" / "react-dom"
  - any module under "src/renderer/"
  - "document" or "window" globals (greppable AST check)
```

The renderer subexport is the **only** browser-coupled entry point. A
Node script can `import { XlsxAgent } from "@officeai/xlsx/agent"`,
parse a file, run agent commands, and serialize — never touching the
DOM or React.

---

## 12. Testing

The renderer is tested with `@testing-library/react` + `jsdom`. Every
test follows the same shape:

```typescript
test("typing into A1 dispatches set-cell-value", async () => {
  const agent = await XlsxAgent.fromBuffer(emptyXlsxBuf);
  const dispatched: Command[] = [];
  spyOn(agent, "applyCommands").mockImplementation(async (cmds) => {
    dispatched.push(...cmds);
    return ok();
  });

  render(<XlsxEditor agent={agent} />);
  const a1 = await screen.findByTestId("cell-0-0");
  await user.dblClick(a1);
  await user.type(screen.getByRole("textbox"), "42{Enter}");

  expect(dispatched).toEqual([
    {
      type: "xlsx:set-cell-value",
      payload: {
        sheet: "Sheet1",
        ref: { row: 0, col: 0 },
        value: { kind: "number", value: 42 },
      },
    },
  ]);
});
```

The pattern mirrors DOCX's renderer tests (which use ProseMirror's test
builders): construct UI, simulate user interaction, assert the
dispatched command sequence. No real model mutation — the agent is
spied on so we test the **funnel**, not the bus.

Test files:

| File                                  | Coverage                                      |
| ------------------------------------- | --------------------------------------------- |
| `xlsx-editor.test.tsx`                | Mount + smoke render + sheet switch           |
| `formula-bar.test.tsx`                | Formula bar input + commit + revert           |
| `grid-virtualization.test.tsx`        | Only visible cells mount; scroll re-windows   |
| `frozen-panes.test.tsx`               | Quadrant split renders correctly              |
| `selection-keyboard.test.tsx`         | All keyboard navigation rules                 |
| `selection-mouse.test.tsx`            | Click, drag, header click, double-click       |
| `cell-editor.test.tsx`                | Enter/Esc/Tab commit semantics                |
| `cell-render-format.test.tsx`         | Number formats, dates, errors, alignments     |
| `cell-render-borders.test.tsx`        | Border style mapping                          |
| `cell-render-fills.test.tsx`          | Fill colors + theme color resolution          |
| `cell-render-spillover.test.tsx`      | EC-V1                                         |
| `cell-render-merged.test.tsx`         | EC-V3 (window-edge merged cell)               |
| `conditional-format-display.test.tsx` | All 5 rule kinds                              |
| `pending-mutations.test.tsx`          | Decoration appears + popover + approve/reject |
| `funnel-discipline.test.tsx`          | Every UI gesture goes through `dispatch`      |
| `headless-guard.test.mjs`             | Renderer subexport import boundary            |

---

## 13. Accessibility

The grid is a `<div role="grid">` with row × column descendants
(`role="row"`, `role="gridcell"`). Selection is announced via
`aria-selected` and the active cell carries `aria-current="true"`.
Cell content is in the accessible tree (no canvas occlusion). Keyboard
navigation works without a mouse for everything except drag-fill (which
is a P1 enhancement).

Color-only signals (CF backgrounds, error red text) are augmented with
`aria-label`s ("error: division by zero") and screen-reader-only text.

---

## 14. What the renderer does NOT do

- **Toolbar UI** — lives in `apps/web`. The renderer exports primitive
  components (`Cell`, `Grid`, `FormulaBar`, `SheetTabStrip`,
  `XlsxEditor`); the consuming app composes them with toolbars,
  ribbons, and chrome.
- **Undo/redo** — the bus owns it. Ctrl+Z dispatches `bus:undo`;
  Ctrl+Y/Ctrl+Shift+Z dispatch `bus:redo`. The renderer just invokes
  these.
- **Persistence** — the agent owns it. Save / Save As are
  app-level concerns.
- **Authoring of out-of-scope features** — pivot tables, chart
  authoring, slicer authoring, sparkline authoring. The right-click
  menus and toolbar entries for these are absent (or disabled with a
  "deferred" tooltip).
- **Direct model mutation** — the renderer never touches the snapshot.
  All edits go through `dispatch(cmd)`.

---

## 15. File layout

```
packages/xlsx/src/renderer/
  XlsxEditor.tsx
  FormulaBar.tsx
  SheetTabStrip.tsx
  Grid.tsx
  RowHeader.tsx
  ColHeader.tsx
  Viewport.tsx
  Cell.tsx
  CellEditor.tsx
  SelectionLayer.tsx
  PendingDecorations.tsx
  hooks/
    useSnapshot.ts
    useSelection.ts
    useDispatch.ts
    useSheetGeometry.ts
    useViewportWindow.ts
    useScrollSync.ts
  format/
    formatCell.ts
    xfToCss.ts
    dxfToCss.ts
    themeColor.ts
    resolveCf.ts
  geometry/
    SheetGeometry.ts
    fenwick.ts
  styles.css                  # Tailwind layer with .xlsx-* utilities
  *.test.tsx
```

Mirrors the DOCX renderer's layout in `packages/docx/src/renderer/` so
the patterns are recognizable across formats.
