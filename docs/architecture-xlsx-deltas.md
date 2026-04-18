# XLSX architectural deltas vs DOCX

> Companion to [`session-summary-xlsx.md`](session-summary-xlsx.md)
> and [`build-log/xlsx.md`](build-log/xlsx.md). The build log is
> chronological; this doc is _spatial_ — it explains the cross-cutting
> shape of the XLSX product against the DOCX baseline so a new
> contributor (or a future LLM agent) does not have to derive these
> by reading both stacks side-by-side.
>
> The mirror in the other direction — what DOCX did differently from
> XLSX — lives in
> [`architecture-docx-deltas.md`](architecture-docx-deltas.md).

The two products share the same headless-first chassis (typed
`CommandBus` from `@officeai/core`, `OoxmlContainer` for byte-
preserving I/O, `DocumentAgent` interface, MCP transport, CLI
shell, Next.js host), but the file formats themselves push against
different architectural axes. The list below is what XLSX did
_differently_ and _why_.

---

## 1. Renderer: hand-rolled grid vs ProseMirror

|                       | DOCX                                                  | XLSX                                                                |
| --------------------- | ----------------------------------------------------- | ------------------------------------------------------------------- |
| Renderer              | `prosemirror-view` + `prosemirror-state`              | hand-rolled virtualized grid in `apps/web/app/xlsx-editor/Grid.tsx` |
| Document-model bridge | bidirectional (PM doc ↔ our model via `doc-to-pm.ts`) | unidirectional (our model → grid; mutations go through the bus)     |
| Selection             | PM `Selection` (positions in a tree)                  | `{ anchor: {r,c}; focus: {r,c} }` (single rectangle)                |
| Input handling        | PM transactions intercepted in plugins                | DOM event handlers dispatch typed commands directly                 |

**Why**: ProseMirror is purpose-built for rich text. A spreadsheet
grid is two-dimensional, sparse, has its own selection algebra,
needs viewport-bounded virtualization, and never needs collaborative
text-flow. Bringing PM along would require modeling cells as
inline atoms — fightable but pointless.

**Cost**: we own input handling (keyboard, mouse, drag, focus
management), virtualization, and selection rendering. **Benefit**:
zero impedance mismatch — every render reads `XlsxSnapshot`
directly; no second source of truth to keep in sync.

## 2. The formula engine: a subsystem that has no DOCX analogue

`packages/xlsx/src/formula/` is a complete five-stage pipeline:

```
lexer  →  parser (precedence-climbing)  →  AST  →  evaluator  →  recalc orchestrator
                                                        ↑                ↑
                                               function registry   dependency graph
                                                                   (Tarjan SCC for cycles)
```

DOCX has nothing comparable. Architectural notes worth the
documentation budget:

- **Synchronous evaluation**. No promises in the hot path — Excel
  semantics treat formula evaluation as a synchronous fixpoint.
  Async I/O (e.g. for an LLM-backed function) would be modeled as a
  separate command, never inside the evaluator.
- **`Value` is closed**. The evaluator sees only
  `number | string | boolean | { error: CellErrorCode }` — there is
  no `null` (empty cells coerce per Excel rules) and no
  `undefined`. This makes every operator's truth table a finite
  table.
- **Tarjan strongly-connected components for cycle detection**.
  Cycles within a SCC produce `#REF!` for every cell in the
  cycle; the rest of the dep graph evaluates in topological order
  around the cycle. This is the textbook approach but worth
  flagging because a naive depth-first recursion would either
  miss cycles or stack-overflow on real workbooks.
- **Volatile functions (NOW, RAND, …) are tagged on registration**.
  Any volatile cell re-fires on every `recalcAll`, even if its
  inputs are unchanged. This is the dependency-graph machinery
  earning its keep — a non-volatile cell with the same inputs
  short-circuits.
- **`listRegisteredFunctions`** exists so the formula autocomplete
  popover (Phase 11d) is sourced from the same registry the
  evaluator reads. Single source of truth — the popover _cannot_
  advertise a function the engine cannot evaluate.

## 3. Style table with content-hash deduplication

DOCX runs each carry their own font / colour / underline props
inline. XLSX has a single workbook-level **style table**:

- One entry per _unique combination_ of font / fill / border /
  alignment / numberFormat (the `<cellXfs>` shape from OOXML).
- Cells reference styles by index (`Cell.styleId: number`).
- `xlsx:set-cell-format` computes a content hash of the resulting
  format and **dedups** — applying Bold twice anywhere in the
  workbook produces zero new style entries on the second call.
- The renderer flattens style indices into CSS via
  `flattenCellXf` + `styleForCell` (web-side helpers exported
  from `@officeai/xlsx`).

**Implication**: the patch payload is patch-style — `{ font: { bold: true } }` —
and the handler is responsible for "find or create" semantics.
DOCX's `set-text-format` writes properties directly onto runs; no
dedup table required.

## 4. Pixels in the model, character widths at the OOXML boundary

OOXML's `cols/@width` is in _character units_ (font-dependent).
Storing widths in the model in those units would force every
renderer to know the workbook's default font.

**Decision**: `Sheet.columnWidths` and `Sheet.rowHeights` store
**CSS pixels** (Phase 11g). The renderer is the source of truth;
the serializer is responsible for the back-conversion at the
OOXML boundary.

**Known caveat (documented in the Phase 11g build-log entry)**:
the back-conversion is currently lossy in one direction — a
foreign reader (Excel desktop) opening a workbook we resized
will see the default column width until a follow-up writer maps
pixels → characters. We accept this for the editor-first MVP
because it keeps the renderer trivial and the resize UX
millisecond-responsive.

## 5. Variable geometry via prefix sums + binary search

DOCX's paged renderer can address the scrollable surface with
page-level offsets — pages are a coarse-grained unit. XLSX cells
have variable column widths and row heights, and the visible
window can fall anywhere.

**Mechanism**:

```ts
const colXs = [0, w0, w0+w1, w0+w1+w2, ...];   // prefix sums
const rowYs = [0, h0, h0+h1, h0+h1+h2, ...];

// visible window:
const startCol = lower_bound(colXs, scroll.left);
const endCol   = lower_bound(colXs, scroll.left + viewport.width);
```

Memoised on `sheet.columnWidths`, `sheet.rowHeights`, and the
transient `colDrag` / `rowDrag` state, so unchanged sheets reuse
the prefix arrays. Visible-window math is `O(log n)` per scroll
event regardless of sheet size.

## 6. Two-surface focus model (cell ↔ formula bar)

DOCX has one focusable editing surface (the PM view). XLSX has
two — the focused cell **and** the formula bar — and Excel users
expect them to behave as a single editing context:

- **Type-to-edit**: a printable key on a focused cell redirects
  the keystroke into the formula bar and parks the caret at the
  end. Implemented as a top-level keyboard handler in
  `XlsxEditor.tsx` that detects the "not yet editing" state and
  forwards to a programmatic focus + value mutation on the formula
  bar.
- **Click-to-insert-ref**: while the formula bar is editing a
  `=`-prefixed expression, clicking another cell appends its A1
  ref at the formula bar's caret position instead of moving the
  selection. Requires `formulaCaretRef` to survive React renders
  (a ref, not state, so render-time reads don't trigger effects)
  and `e.preventDefault()` on the cell click so DOM focus does
  not leave the formula bar before the insertion runs.

**Why a ref instead of state**: caret position changes on every
keystroke. Storing it in state would cascade re-renders into the
suggestions popover (which depends on prefix-at-caret). The ref +
imperative-position pattern keeps render cost flat in the typing
hot path.

## 7. Diff vocabulary is XLSX-specific

DOCX diffs are mostly text-shaped: `text-inserted`, `text-deleted`,
`style-updated`, `paragraph-inserted`, etc. XLSX adds:

| Diff kind                              | Used by                                                             |
| -------------------------------------- | ------------------------------------------------------------------- |
| `cell-updated`                         | every command that mutates values, including recalc cascade entries |
| `formula-updated`                      | `set-cell-formula`, insert/delete-row/column (formula text rewrite) |
| `format-updated`                       | `set-cell-format`                                                   |
| `style-added`                          | `set-cell-format` when a new `xfId` is appended                     |
| `rows-inserted` / `rows-deleted`       | structural ops                                                      |
| `columns-inserted` / `columns-deleted` | structural ops                                                      |
| `referenced-cell-deleted`              | `delete-row` / `delete-column` per `EC-R2` / `EC-F4`                |
| `merge-added` / `merge-removed`        | merge / unmerge                                                     |
| `sheet-added`                          | `add-sheet`                                                         |
| `sheet-renamed`                        | `rename-sheet`                                                      |
| `comment-added`                        | `add-comment`                                                       |
| `node-updated` (with `meta.kind`)      | sizing commands (Phase 11g)                                         |

Recalc side-effects show up as `cell-updated` entries with a
`source: "recalc"` marker — the diff log distinguishes "the user
set B2 = 5" from "B4's =SUM(B2:B3) re-fired because B2 changed".

## 8. Drag interactions: command on `mouseup`, transient preview locally

DOCX has nothing comparable. Header drag-to-resize (Phase 11g)
needs to feel rubber-band responsive without saturating the
command bus.

**Approach**: the Grid keeps `colDrag` / `rowDrag` _local_ state
that follows every `mousemove`. The visible width / height
during drag is read from this transient state, not from the
agent snapshot. **One** command (`xlsx:set-column-width` or
`xlsx:set-row-height`) is dispatched on `mouseup` with the final
value. The diff log stays usable, undo restores the pre-drag
state in one step, and the grid stays at 60 fps during drag.

The same pattern is the right answer for any future continuous
input (range fill handles, drag-to-move selection, conditional
formatting paint brush): preview locally, commit one command
on commit gesture.

## 9. Replace-agent on file open

DOCX presumably hot-swaps documents into the existing PM view.
XLSX (Phase 11a) takes a different stance: opening a `.xlsx` from
disk **constructs a fresh `XlsxAgent`** via `fromBuffer` and
swaps the React agent prop. Reasoning:

- A new file means a new `partHashes` baseline. Mutating the
  existing agent in place would either drop the old baseline
  (corrupting the byte-equality oracle for the new file) or
  carry it over (silently invalidating it).
- A new file means a new undo history. Carrying the old
  command stream would let `Cmd-Z` resurrect deleted cells from
  the previous file — uniformly bad UX.
- Snapshot subscription is rewired in one render pass; the
  revision counter resets to 0.

The drag-and-drop overlay uses the same code path — drop is
just an alternate file picker.

## 10. Phase 7e parallel-agent function-library build

Worth documenting as a **methodology** because we'll do it again
when we add the next 60 functions.

The 89 P0 functions span five categories (math, logic, info,
lookup, text). Building them sequentially would have been ~5
days of typing tests. Instead, Phase 7e launched **5 sub-agents
in parallel**, each owning one category, with these
preconditions:

1. The function registry interface (Phase 7c) was frozen —
   sub-agents could not modify it.
2. The `Value` union and error model (Phase 7a) were frozen —
   no agent could invent a new error code.
3. Each sub-agent owned its own `*-functions.ts` file plus its
   own `*-functions.test.ts`, so there was zero file conflict.
4. The integration step (a single barrel `index.ts` listing all
   five modules) was reserved for the parent agent.

Result: 89 functions + 302 tests landed in one phase, integrated
in a single follow-up commit. The pattern generalises to any
work that decomposes into independent leaves with a frozen
contract.

## 11. Number-format presets are a UX layer over OOXML built-ins

Excel's number formatting is keyed by `numFmtId`: 0–49 are
built-in, 164+ are custom format strings. The full vocabulary is
overwhelming for a toolbar dropdown.

**Layer**: `apps/web/app/xlsx-editor/styles.ts` defines a small
preset list (General, Number, Currency €, Currency $, Percent,
Date) and `presetNumFmtId(presetKey)` maps each to the
underlying built-in `numFmtId`. The dropdown speaks human; the
command speaks OOXML.

**Render path**: `formatCellValue(value, numFmtId)` applies the
format code to the raw value at render time. The model always
stores raw values (numbers as numbers, dates as Excel serials).
This keeps formula evaluation working on the raw value while the
display shows the formatted string — exactly Excel's discipline.

## 12. Merge rendering: oversized top-left + covered set

Merge regions are stored as `{ r1, c1, r2, c2 }` rectangles. The
renderer cannot simply "draw the same content in every covered
cell" — it has to draw a single oversized cell at `(r1, c1)`
spanning the rectangle, and not draw the covered cells at all.

**Mechanism** (`mergeIndex` in `Grid.tsx`):

```ts
const topLeft = new Map<key, MergedRect>(); // (r1,c1) → rect
const covered = new Set<key>(); // every (r,c) inside the rect except (r1,c1)
```

Per-cell loop:

```ts
if (covered.has(key))   continue;          // skip
if (topLeft.has(key))   span = lookup;    // draw oversized
else                    span = 1×1;        // normal cell
```

Merged-cell width / height is computed via the same prefix-sum
arithmetic (item 5) so a merge across resized columns lays out
correctly without special-casing.

## 13. Selection model: single rectangle, intentionally

XLSX selection is `{ anchor, focus }` and represents _exactly one_
rectangle. Excel supports multi-rectangle selection (Ctrl-click
to add disjoint areas), and it's intentionally **not** in scope.

**Why single-rectangle**: 90% of toolbar operations apply uniformly
to a rectangular range. Multi-rectangle adds significant
complexity to the selection algebra (intersect/union, marquee
rendering, merge interactions, range-aware command fan-out) for
the long-tail of UX. Listed in `feature-scope.md` as deferred.

**Implication**: every command that operates on a "selection"
in the web layer fans out one command per cell or operates on
the single `selectionToRange()` rectangle. There is no
multi-range plumbing to bypass when the time comes — the
selection type is the bottleneck, not the command shape.

## 14. Test pyramid is shaped differently

| Layer            | DOCX                                     | XLSX                                                                                                                                                    |
| ---------------- | ---------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Unit             | model / serializer / handlers            | model / serializer / handlers + **formula engine (475 tests)** + style-table dedup                                                                      |
| Integration      | round-trip oracle on real-world fixtures | round-trip oracle + **per-command property tests for the inverse mutation**                                                                             |
| E2E (Playwright) | editor smoke against bundled DOCX        | editor smoke + **drag interactions** (resize, drag-extend selection) + **caret-aware formula bar tests** (autocomplete acceptance, click-to-insert-ref) |

**Why drag-aware e2e is XLSX-specific**: nothing in the DOCX UX
depends on a `mousedown → mousemove → mouseup` sequence with a
specific commit point. XLSX has three (drag-extend selection,
column resize, row resize) and they all test the
local-preview / commit-on-up architecture from item 8.

**Why caret-aware e2e is XLSX-specific**: the formula bar is the
only place in either product where caret position has _semantic_
significance (clicking a cell appends an A1 ref at the caret).

---

## 15. Two-layer formula rendering: strict lexer + permissive scanner

|                     | DOCX                   | XLSX                                                                                                          |
| ------------------- | ---------------------- | ------------------------------------------------------------------------------------------------------------- |
| Display tokeniser   | n/a                    | `tokenizeForDisplay()` in `packages/xlsx/src/formula/highlight.ts` — never throws, contiguous-cover guarantee |
| Evaluator tokeniser | n/a                    | `lex()` in `packages/xlsx/src/formula/lexer.ts` — strict, throws on malformed input, source of truth for AST  |
| Why two             | only one consumer (PM) | the highlighter is **always** asked to render mid-typing input (`=A1+`, `=SUM("hello`); the evaluator is not  |

**Why two scanners**: forcing the strict lexer to be permissive
would erode its evaluator contract; teaching the highlighter to
swallow exceptions would scatter throw/catch noise across React
render paths. Splitting them keeps each one's invariants tight.

**Cost**: a small amount of duplication in regex catalogues for
references and operators. **Benefit**: `tokenizeForDisplay` can
guarantee `tokens[i].end === tokens[i+1].start` (contiguous
coverage) — which the formula bar overlay relies on to align
glyphs with the underlying transparent input character-by-character.

The `assignRefColors` companion hashes by a normalised `refKey`
(uppercased, `$`-stripped, sheet-qualified), so `A1`, `$A$1`, and
`a1` share a colour while distinct addresses cycle through the
8-colour `DEFAULT_REF_COLORS` palette. DOCX has no parallel
machinery — there are no "expressions" inside docs that need
chromatic-distinct rendering.

---

## 16. Two-surface input model with a transparent overlay

|                      | DOCX                             | XLSX                                                                                                            |
| -------------------- | -------------------------------- | --------------------------------------------------------------------------------------------------------------- |
| Where the user types | the document surface (PM editor) | one of three: cell `<input>`, formula bar `<input>`, or the surface `<div>` (type-to-edit redirects to formula) |
| Coloured rendering   | n/a                              | `FormulaHighlight.tsx` overlay layered behind the formula bar — same font/spacing, transparent input text       |
| Caret ownership      | PM owns the caret                | the `<input>` keeps the caret; overlay is `pointer-events:none, aria-hidden`                                    |

**Why an overlay instead of `contentEditable`**: the formula bar
already owns a long history of careful caret behaviour (P11
click-to-insert-ref, autocomplete acceptance, type-to-edit
redirection). Migrating to `contentEditable` would force re-
verifying every one of those flows. The overlay pattern keeps
the input intact and only changes how it _looks_; the only
synchronisation primitive is mirroring `scrollLeft` on the
overlay via `transform: translateX` so long formulas stay aligned.

**Trade-off**: IME composition is mildly fragile — half-composed
characters briefly show in the default text colour before the
overlay catches up. Acceptable for English-formula scope; revisit
when CJK formula authors arrive.

---

## 17. Ref-rectangle highlighting: separate visual layer

|                 | DOCX                     | XLSX                                                                          |
| --------------- | ------------------------ | ----------------------------------------------------------------------------- |
| Selection paint | PM decoration on the doc | `<div data-testid="grid-marquee">` absolutely positioned over `colXs / rowYs` |
| Other overlays  | n/a                      | `refRects` rendered as 2px dashed coloured borders (zIndex 3, below marquee)  |

**Why a separate visual layer for refs**: selection and ref-
highlight have different lifecycles (selection is persistent;
ref-highlight only exists while the formula bar is focused with
a `=` formula) and different colour vocabularies (selection is
always violet; refs cycle through the 8-colour palette). Painting
them as siblings keeps z-order obvious and lets the marquee win
on overlap, which is what Excel does.

DOCX has no analogue — there are no "this paragraph is
referenced by that other paragraph" relationships at the
rendering layer.

---

## 18. Keyboard parity: surface vs in-edit modes (P12)

|                              | DOCX                             | XLSX                                                                                                                         |
| ---------------------------- | -------------------------------- | ---------------------------------------------------------------------------------------------------------------------------- |
| Where keyboard handlers live | PM plugins (one tree, one model) | three: surface div (`onSurfaceKeyDown`), formula bar `<input>` (`onKeyDown`), in-cell `<input>` (`onKeyDown`)                |
| Mode dispatch                | PM cursor position               | `formulaFocused` × `formulaDraft.startsWith("=")` × `selection.kind` (single vs range vs whole-row vs whole-col)             |
| Excel-isms shipped           | n/a                              | Arrow nav (+ Shift extend, + Ctrl jump-to-data-edge), Tab/Enter commit-and-move, F2 enter edit, Delete drops whole rows/cols |

**Why three handlers instead of one**: each surface owns
different invariants — the surface div manages selection and
type-to-edit redirection, the formula bar input manages the
draft and click-to-insert-ref, the in-cell input is a short-
lived peer to the formula bar. Folding them into one would
require either lifting all state up or constantly checking
`document.activeElement`. The split keeps each handler's deps
narrow and React's re-render graph predictable.

**Why `Delete` deletes whole rows/columns instead of `Cmd+−`**:
explicit user request, but it also dodges Chromium's zoom-out
shortcut interception in headless Playwright. The behaviour
gracefully degrades — a single-cell Delete still clears, only
whole-row / whole-col selections trigger structural deletion.

---

---

## 14. Clipboard, fill, and undo (Phase 13)

Three Excel-native interactions had no DOCX analogue and forced
new architecture in this layer:

| Surface           | DOCX                                                          | XLSX                                                                                                                                                |
| ----------------- | ------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------- |
| Clipboard         | OS-default cut/copy of the PM selection (no model awareness)  | `XlsxClipboardSnapshot` round-trips through TSV+HTML with a fingerprint, parsed by a headless `parseExternalClipboard` that handles Excel/Sheets/Numbers/CSV without touching the DOM. |
| Fill / drag       | n/a                                                           | `xlsx:fill-range` + a 6-detector series engine in `packages/xlsx/src/fill/series.ts` (numeric, date, weekday, month, text-numeric, repeat).         |
| Multi-step Undo   | Per-mutation `before` snapshot; bus had no Undo until Phase 13 | Same `before` snapshots, plus a `redoStack` and `canUndo`/`canRedo`/`undo`/`redo` on `CommandBus`. `MutationStatus` gains `"undone"`. Both agents proxy. |

**Why headless clipboard parsing**: keeping `parseHtmlTable` /
`parseFingerprintHtml` / `delimitedToSnapshot` in
`@officeai/xlsx` means the same code path handles "user pastes
Excel-Desktop HTML in the browser" and "agent pastes a CSV string in
a Node-only test". No `jsdom`. No browser dependency. Real fixtures
in `packages/xlsx/src/clipboard/__fixtures__/` cover Excel Desktop,
Google Sheets, Apple Numbers, German CSV (semicolon delimiter), and
multiline-quoted CSV.

**Why redo re-runs the handler instead of restoring `after`**:
between an `undo()` and a `redo()` the user (or an agent) can
dispatch other mutations, rebase the pending stack, or move sheets
around. Re-applying the original payload against the *current*
approved snapshot is the only way to keep redo correct in the face
of a rebase. The cost is that pure handlers are mandatory — which
they already are.

**Why Cmd+Y is also accepted**: muscle memory from Excel-on-Windows
where Ctrl+Y is the redo. Cmd+Shift+Z (macOS-native) is the primary
binding.

---

## When this doc should be updated

- A new editor surface lands and shares non-trivial machinery
  with one of the existing two — the contrast table grows a
  PPTX column, or a row collapses if XLSX/DOCX converge.
- A subsystem here is replaced (e.g. the hand-rolled grid is
  swapped for `react-virtualized` or a competitor) — bump the
  relevant section with the new substrate and the migration
  rationale.
- The deferred items in items 4 (OOXML char-width round-trip)
  or 13 (multi-rectangle selection) are picked up — strike the
  caveat and link to the closing build-log entry.
