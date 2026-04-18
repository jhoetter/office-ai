# XLSX — Edge Cases

> Known hard cases and how the editor degrades. Each entry has a
> "trigger" (what the user / agent did) and a "behaviour" (what we ship
> in P0). Things we cannot handle deterministically surface a
> structured error rather than a silent corruption — per the
> "fail loudly" architecture principle.

## Reference handling

### EC-R1: Insert row above a referenced cell

- **Trigger:** `=A5` exists; agent inserts 2 rows at row 3.
- **Behaviour:** All references at or below row 3 shift down by 2. `=A5` becomes `=A7`. Absolute refs (`=$A$5`) also shift (Excel parity). Refs _above_ the inserted band do not shift.
- **Failure mode:** None expected. Deterministic.

### EC-R2: Delete a row containing a referenced cell

- **Trigger:** `=A5` exists; agent deletes row 5.
- **Behaviour:** `=A5` becomes `=#REF!` (the literal error in formula text), and `cachedValue` becomes `#REF!`. The mutation `diff` includes a `referenced-cell-deleted` summary item.

### EC-R3: Delete a column containing a defined name's range

- **Trigger:** Defined name `taxRate = Sheet1!$B$1`; agent deletes column B.
- **Behaviour:** Defined name is preserved with `=#REF!` per Excel. Mutation diff lists the broken defined name.

### EC-R4: Cross-sheet reference where target sheet was renamed

- **Trigger:** Agent renames `Sheet2` → `Inputs`. `=Sheet2!A1` exists in `Sheet1`.
- **Behaviour:** All references to `Sheet2` rewrite to `Inputs`. Quoted-form refs (`='Sheet 2'!A1`) handled identically.

### EC-R5: 3D reference

- **Trigger:** `=SUM(Sheet1:Sheet3!A1)` exists; agent renames `Sheet2`.
- **Behaviour:** P0 — preserve the 3D ref token verbatim and do **not** re-parse it. The formula string is unchanged. Sheet rename does NOT rewrite the 3D ref. Documented as a known deferral; emit a warning in the build log on the first occurrence.

## Formulas

### EC-F1: Circular reference

- **Trigger:** `A1 = =B1`, `B1 = =A1`.
- **Behaviour:** Dep graph builder detects the cycle, sets both cells' `cachedValue` to `#REF!`, attaches `{ kind: "circular", cycle: ["Sheet1!A1", "Sheet1!B1"] }` to the mutation diff.
- **Iterative calc:** explicitly not supported in P0. We do not attempt convergence.

### EC-F2: Unknown function name

- **Trigger:** Agent writes `=DOSOMETHING(A1)`.
- **Behaviour:** Lexer + parser succeed; evaluator returns `#NAME?`. Formula string is preserved. The cell is flagged in the model so the renderer can show a "no such function" tooltip.

### EC-F3: Unparseable formula

- **Trigger:** OOXML imports `=` (empty) or `=A1+++` from a file authored by an Excel internal preview build.
- **Behaviour:** Parser reports a structured `FormulaParseError`. Cell stores the original string in `formula.raw` and a `cachedValue` of `#NAME?`. We do **not** drop the formula string — round-trip preserves it. Mutation that introduces such a formula is rejected with `precheck` failure.

### EC-F4: Reference to deleted sheet

- **Trigger:** Agent deletes `Sheet2`. `=Sheet2!A1` exists.
- **Behaviour:** Formula rewritten to `=#REF!`. `cachedValue` = `#REF!`. Diff lists every affected cell.

### EC-F5: Volatile function on a 50k-row sheet

- **Trigger:** `B1:B50000 = =RAND()`.
- **Behaviour:** Mark all volatile cells dirty on every recalc. Recalc runs in topological order. We assert the perf budget (`< 100 ms` on 10k dependents) but explicitly do **not** assert it on 50k volatile cells — that's a documented "agent should batch" case.

### EC-F6: Shared formula (`<f t="shared" si="0" ref="B2:B100">=A2+1</f>`)

- **Trigger:** Imported from Excel; agent edits one cell in the share group.
- **Behaviour:** On import, expand all shared formulas to per-cell formulas (memory cost OK at our scope). On serialize, optionally re-share contiguous identical-shape groups. Edit to one cell breaks the share group; all others retain their per-cell formulas. Documented in the build log.

### EC-F7: Array formula (`<f t="array" ref="A1:C3">=...</f>`)

- **Trigger:** Imported.
- **Behaviour:** P0 — preserve as opaque-on-the-formula-cell. Do not evaluate. The "anchor" cell stores the raw formula; spilled cells reference the anchor by id. Round-trips clean.

## Storage / model

### EC-S1: Sparse cell distribution

- **Trigger:** A workbook has a cell at `XFD1048576` (Excel max) and one at `A1`.
- **Behaviour:** `Sheet.cells: Map<string, Cell>` handles arbitrary sparseness. Renderer uses windowing; agents read by ref, not iteration.

### EC-S2: 1,048,576-row sheet

- **Trigger:** Excel max rows.
- **Behaviour:** Parse + serialize: O(used cells) not O(rows). Renderer: never iterates beyond viewport.

### EC-S3: Merged cell that overlaps an inserted row

- **Trigger:** `A1:C3` merged; agent inserts row at row 2.
- **Behaviour:** Merge expands to `A1:C4`. If the merge entirely contains the inserted band, the merge expands. If the inserted band crosses the merge boundary partially, P0 rejects the mutation with `precheck` failure (`merge-boundary-crossed`).

### EC-S4: Cell with both a value and a formula

- **Trigger:** Excel writes `<c><f>=A1+1</f><v>42</v></c>`.
- **Behaviour:** `formula.text = "=A1+1"`, `cachedValue = 42`. Our recalc updates `cachedValue` after parsing.

### EC-S5: Cell with no `<v>` (formula stored without cached value)

- **Trigger:** Some authoring tools omit `<v>` on first save.
- **Behaviour:** Compute `cachedValue` on import via the formula engine. If the formula is unsupported, set `cachedValue = #NAME?`.

### EC-S6: Inline string vs shared string table

- **Trigger:** Excel uses `<is><t>foo</t></is>` for some cells, `sst` for others, in the same sheet.
- **Behaviour:** Both shapes parse correctly. On serialize, prefer SST when the string appears more than once in the workbook; inline otherwise.

## OOXML round-trip

### EC-O1: Untouched part with whitespace differences from re-serialize

- **Trigger:** `xl/styles.xml` was reformatted in some Excel build to use 2-space indent.
- **Behaviour:** Untouched parts are NOT re-serialized; they're emitted byte-for-byte from the input zip. Whitespace is preserved by definition.

### EC-O2: Modified part where SheetJS re-emits with attribute reordering

- **Trigger:** Agent sets one cell value; SheetJS re-emits `xl/worksheets/sheet1.xml` with attributes in a different order than Excel.
- **Behaviour:** Acceptable. The "untouched parts byte-equal" criterion only applies to parts not in the dirty set. Modified parts must be **structurally equivalent**, not byte-equal.

### EC-O3: Pivot cache references a deleted column

- **Trigger:** Agent deletes a column that a pivot caches data from.
- **Behaviour:** Pivot is preserved as opaque. We do **not** invalidate the cache; downstream Excel will refresh it on open. Documented as a known limitation.

### EC-O4: Custom XML data binding references our edited range

- **Trigger:** Custom XML `<x:Map>` references `Sheet1!A1`.
- **Behaviour:** Custom XML is opaque. Edit succeeds; binding is preserved. If Excel later detects a binding mismatch, it surfaces a non-blocking warning. Documented.

### EC-O5: External link to another workbook

- **Trigger:** `xl/externalLinks/externalLink1.xml` exists.
- **Behaviour:** Preserved as opaque. Do not evaluate cross-workbook references; surface them as `#N/A` in `cachedValue`.

## Concurrency / staging

### EC-C1: Two agent commands in flight, one references the other's mutation

- **Trigger:** Agent A: `xlsx:set-cell-value B1 = 5`. Agent B: `xlsx:set-cell-formula C1 = "=B1*2"`. Both pending.
- **Behaviour:** Working snapshot includes both. `cachedValue` of `C1` = `10` based on working `B1`. If A is rejected, `C1` recalcs from approved `B1`. The mutation `diff` for B annotates the dependency.

### EC-C2: Agent rejected after a downstream agent depended on it

- **Trigger:** EC-C1 with A rejected.
- **Behaviour:** B's `cachedValue` is recomputed from approved snapshot. Mutation B's `before` and `after` are recomputed (technically, the rebase recomputes B against the new approved). Documented.

## I/O failures

### EC-I1: Truncated zip

- **Trigger:** `XlsxAgent.fromBuffer(corruptedBuf)`.
- **Behaviour:** `XlsxParseError` with structured `cause` (zip-corruption / missing-part / invalid-xml). Never silent.

### EC-I2: Missing required part (`xl/workbook.xml`)

- **Trigger:** Zip is valid but not an Excel file.
- **Behaviour:** `XlsxParseError("missing-required-part: xl/workbook.xml")`.

### EC-I3: XML namespace collision

- **Trigger:** Some authoring tool re-prefixes the main xlsx namespace.
- **Behaviour:** Parser uses namespace URI matching, not prefix matching. Tested via a synthetic fixture.

### EC-I4: Non-UTF-8 part

- **Trigger:** Old Excel build emits Windows-1252 in `xl/comments1.xml`.
- **Behaviour:** Detect via XML decl `encoding="windows-1252"` and decode accordingly. Re-emit as UTF-8 on serialize.

## Renderer

### EC-V1: Cell content longer than column width

- **Trigger:** `A1 = "very long string"`, column A width = 8.
- **Behaviour:** Spill into adjacent empty cells (Excel parity). Stop at the first non-empty neighbour.

### EC-V2: Cell with a date number-format applied to a string value

- **Trigger:** `A1 = "hello"`, format = `m/d/yyyy`.
- **Behaviour:** Render the literal string. Excel does the same.

### EC-V3: Window-edge merged cell

- **Trigger:** Merge `A1:Z1`, viewport shows columns C–H.
- **Behaviour:** Renderer paints the merge band at its anchor (A1) clipped to the viewport. Clicking anywhere in the band selects the anchor cell.

### EC-V4: 50,000-row scroll

- **Trigger:** User flings the scrollbar to the bottom.
- **Behaviour:** Only viewport rows render. Scroll is smooth (verified by the perf budget in `acceptance-criteria.md` G5).

## Out-of-scope features that the agent might attempt

If the agent dispatches a command for an out-of-scope feature, the bus
rejects with a structured `not-supported` mutation:

- `xlsx:edit-pivot` → rejected (`pivot-editing-out-of-scope`).
- `xlsx:add-chart` → rejected (P0; deferred to P1).
- `xlsx:set-defined-name` → rejected (P0; deferred to P1).
- `xlsx:run-macro` → rejected (`macros-not-executed`).

The MCP layer does not expose tools for these in P0, so the only path
to surface the rejection is a hand-crafted `xlsx_apply` call. The
rejection makes the failure mode visible rather than corrupting the
file.
