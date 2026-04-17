# XLSX — Agent Commands

> The 13 P0 commands every `xlsx_*` MCP tool, every CLI invocation,
> and every renderer interaction dispatch through the bus
> ([`spec/shared/command-bus.md`](../shared/command-bus.md)). Each
> entry below carries the typed payload, the handler step list, the
> OOXML impact, the `precheck` rejection cases, an example, and the
> inverse mutation description (for the property test in
> [`acceptance-criteria.md`](acceptance-criteria.md) `G3`).
>
> Required reading first:
>
> - [`feature-scope.md`](feature-scope.md) — what's in P0.
> - [`edge-cases.md`](edge-cases.md) — every command cites the
>   relevant `EC-*` rules.
> - [`formula-engine.md`](formula-engine.md) — recalc invariants for
>   every command that touches cells.
> - [`spec/docx/agent-commands.md`](../docx/agent-commands.md) —
>   shape this doc mirrors.
> - [`prompt.md`](../../prompt.md) lines 396–410 — the original
>   command list this fully types.

---

## 0. Common types

```typescript
import type { Command } from "../shared/command-bus.md";

export interface XlsxCommandBase<TType extends string, TPayload> extends Command<TType, TPayload> {}

/** Cell value union — every primitive Excel can store in `<c>`. */
export type CellValue = number | string | boolean | null | { error: CellErrorCode };

export type CellErrorCode = "#DIV/0!" | "#NAME?" | "#VALUE!" | "#NUM!" | "#N/A" | "#REF!" | "#NULL!";

/** Patch-style format payload — undefined fields are left unchanged.
 *  Mirrors the `TextFormatPayload` discipline from DOCX. */
export interface CellFormat {
  font?: {
    family?: string;
    size?: number; //  pt
    bold?: boolean;
    italic?: boolean;
    underline?: boolean;
    strike?: boolean;
    color?: string; //  RRGGBB no #
  };
  fill?: {
    color?: string; //  RRGGBB no #
    pattern?: "solid" | "none";
  };
  border?: {
    top?: BorderSide;
    right?: BorderSide;
    bottom?: BorderSide;
    left?: BorderSide;
  };
  alignment?: {
    horizontal?: "left" | "center" | "right" | "fill" | "justify";
    vertical?: "top" | "middle" | "bottom";
    wrapText?: boolean;
    indent?: number;
  };
  numberFormat?: string; //  built-in id (as string) or custom format string
}

export interface BorderSide {
  style?: "thin" | "medium" | "thick" | "double" | "dashed" | "dotted" | "none";
  color?: string; //  RRGGBB no #
}
```

### `precheck` contract

Every handler that has a precondition exposes a `precheck(snapshot,
payload, ctx) → { ok: true } | { ok: false; reason: string;
suggestedFix?: string }`. The bus calls `precheck` before `apply`;
on `{ ok: false }` it produces a `Mutation` with `status: "rejected"`
whose `diff` carries the `reason` and `suggestedFix`. Per
[`analysis.md`](analysis.md) §1.3: this is the structured "no, but
here's why" surface that lets an LLM recover.

The single shape:

```typescript
export interface PrecheckFailure {
  ok: false;
  /** Stable kebab-case id; agents should switch on this. */
  reason: string;
  /** Human-readable hint with a concrete next step. */
  suggestedFix?: string;
}
```

### Inverse mutation discipline

Per [`acceptance-criteria.md`](acceptance-criteria.md) `G3`: every
handler ships an `inverse(before, payload, after) → Command` so the
property test in `handlers.test.ts` can assert
`apply(redo) ∘ apply(undo) === identity` against a Vitest property
generator. Every entry below names the inverse explicitly.

---

## 1. `xlsx:set-cell-value`

```typescript
type SetCellValuePayload = {
  sheet: string; //  sheet name (case-sensitive; reject on miss)
  ref: string; //  A1 single cell, e.g. "B2"
  value: CellValue;
};
```

### Behaviour

1. Resolve the sheet by name; reject `unknown-sheet` on miss.
2. Parse `ref` via `references.parseA1`; reject `invalid-ref` on
   malformed text.
3. If the cell currently holds a formula, drop the formula and
   register the new literal: `engine.addCell(ref, null, value)`.
4. Otherwise overwrite the cell value: `engine.addCell(ref, null,
value)`.
5. Call `engine.recalc()` to update every dependent's
   `cachedValue`.
6. Mark `dirty.parts.add("xl/worksheets/sheetN.xml")`.
7. If `value` is a string and the SST is in use, intern via
   `sst.intern(value)` and set the cell's `t = "s"`.
8. Emit a single `cell-updated` change to the diff plus one
   `cell-updated` per cell whose `cachedValue` changed in step 5.

### OOXML impact

- Dirties: `xl/worksheets/sheetN.xml`. May dirty
  `xl/sharedStrings.xml` if a new string was interned.
- Creates: nothing.
- Untouched parts (every other `<c>` on the sheet plus all other
  workbook parts) round-trip byte-identically per
  [`feature-scope.md`](feature-scope.md) "File I/O".

### Excel-parity (per `prompt.md` line 396)

`CellValue = number | string | boolean | null | { error }` covers
every primitive form Excel stores in `<c>`. Setting `null` clears the
cell (drops the `<c>` element entirely, not "empty string").

### `precheck`

| `reason`            | When                                          | `suggestedFix`                                                        |
| ------------------- | --------------------------------------------- | --------------------------------------------------------------------- |
| `unknown-sheet`     | `sheet` not in workbook                       | "Use one of: \[sheet names]"                                          |
| `invalid-ref`       | `ref` not a single A1 cell                    | "Use a single cell ref like 'B2', not a range"                        |
| `merged-non-anchor` | `ref` is a non-anchor cell of a merged region | "Target the anchor cell '\[anchor]' or call xlsx:unmerge-cells first" |
| `formula-string`    | `value` is a string starting with `=`         | "Use xlsx:set-cell-formula for formulas"                              |

The `merged-non-anchor` rejection is per `EC-S3`-adjacent semantics:
Excel allows writes to non-anchor merged cells but discards them on
save; we surface the loss instead.

### Example

```json
{
  "type": "xlsx:set-cell-value",
  "payload": { "sheet": "Q1", "ref": "B2", "value": 42 }
}
```

Before: `B2` is blank. After: `B2 = 42`. Diff:
`{ kind: "cell-updated", path: ["sheets", "Q1", "cells", "B2"],
   before: null, after: 42 }`.

### Inverse

`xlsx:set-cell-value` with `value = before.cell?.value ?? null`. If
the previous cell had a formula, the inverse is the equivalent
`xlsx:set-cell-formula` carrying the original formula text — the
property-test fixture stores the full `before.cell` snapshot and
chooses the right inverse type accordingly.

### Edge cases cited

- `EC-S1` (sparse cells): write to `XFD1048576` is supported.
- `EC-S3` (merged cells): rejected via precheck above.
- `EC-S4`/`EC-S5`: behaviour above is consistent with importing a
  cell that has both a value and a formula — the value is the
  cached one, this command replaces both.

---

## 2. `xlsx:set-cell-formula`

```typescript
type SetCellFormulaPayload = {
  sheet: string;
  ref: string; //  A1 single cell
  formula: string; //  with or without leading '='
};
```

### Behaviour

1. Resolve the sheet; reject `unknown-sheet` on miss.
2. Parse `ref`; reject `invalid-ref` on malformed text.
3. Strip an optional leading `=` from `formula`.
4. Call `engine.parse(formula, { row, col, sheet })`. On
   `FormulaParseError`:
   - If `kind === "empty-formula"`, this becomes a
     `xlsx:set-cell-value` with `value = null`.
   - Otherwise reject with `precheck` `formula-parse-error` and
     attach `parseError.kind` to the rejection metadata. The cell
     is NOT silently flagged as `#NAME?` here — that path only fires
     on **import** (per `EC-F3`); agent-issued unparseable formulas
     are rejected loudly so the LLM gets feedback.
5. `engine.addCell(ref, formula, null)` → `engine.recalc()`.
6. Write `formula.text` to the cell's `<f>` element; write the
   recalc'd `cachedValue` to `<v>`.
7. Dirty `xl/worksheets/sheetN.xml`. May dirty `xl/calcChain.xml`
   if present (or drop it; we currently re-emit `calcChain.xml` on
   any calculation change for safety).
8. Diff: one `formula-updated` change for the source cell, one
   `cell-updated` for every dependent whose `cachedValue` changed.
   On `recalc.cycles` non-empty, the diff also includes a
   `circular` change with `meta.cycle = [...refs]` per `EC-F1`.

### OOXML impact

- Dirties: `xl/worksheets/sheetN.xml`, optionally `xl/calcChain.xml`.
- Creates: nothing new in P0.
- Untouched parts byte-preserved.

### Excel-parity (per `prompt.md` line 397)

The formula text is preserved verbatim (we don't canonicalise
whitespace or quote-style). The cached value is what Excel would
display after a recalc-all. The structured cycle metadata is a
deliberate **superset** of Excel's behaviour, which silently
rewrites `#REF!` cycles back to a default value.

### `precheck`

| `reason`                | When                                         | `suggestedFix`                                          |
| ----------------------- | -------------------------------------------- | ------------------------------------------------------- |
| `unknown-sheet`         | sheet missing                                | "Use one of: \[sheet names]"                            |
| `invalid-ref`           | ref malformed or a range                     | "Use a single cell ref like 'B2'"                       |
| `formula-parse-error`   | parser threw                                 | repeats `parseError.kind` and span                      |
| `merged-non-anchor`     | ref is a continuation cell of a merge        | "Target the anchor cell '\[anchor]'"                    |
| `unsupported-construct` | parser raised a `*-not-supported` error kind | hint specific to the construct (3D ref, structured ref) |

### Example

```json
{
  "type": "xlsx:set-cell-formula",
  "payload": { "sheet": "Q1", "ref": "C2", "formula": "=SUM(B2:B10)" }
}
```

Before: `C2` blank, `B2:B10 = [1,2,3,4,5,6,7,8,9]`. After: `C2 = 45`,
formula `=SUM(B2:B10)`. Diff includes the formula-update and the
cached-value change.

### Inverse

Mirror of §1: `xlsx:set-cell-formula` with `formula = before.cell.formula.text`,
or `xlsx:set-cell-value` with `value = before.cell.value` if the
previous cell carried a literal.

### Edge cases cited

- `EC-F1` (circular): handled in step 8.
- `EC-F2` (unknown function): parser succeeds; eval returns `#NAME?`;
  the `cachedValue` reflects this. Not a precheck failure.
- `EC-F3` (unparseable): rejected with `formula-parse-error`.
- `EC-F4` (deleted-sheet ref): the reference resolves at parse time;
  if the named sheet doesn't exist, the formula is parsed but the
  cached value will be `#REF!` after recalc. Not a precheck failure
  (cross-sheet refs to future sheets are intentional sometimes).
- `EC-F6` (shared formula): the resulting per-cell formula is stored
  expanded; the serializer's optional re-share pass may compact it
  into a `<f t="shared">` group on save.
- `EC-R5` (3D ref): rejected via `unsupported-construct`.

---

## 3. `xlsx:set-range-values`

```typescript
type SetRangeValuesPayload = {
  sheet: string;
  range: string; //  A1 range, e.g. "A1:C3"
  values: CellValue[][]; //  row-major; dimensions MUST equal range
};
```

### Behaviour

1. Resolve sheet and range; reject `unknown-sheet` / `invalid-range`.
2. Validate dimensions: `values.length === rangeRows` and every
   `values[i].length === rangeCols`. Reject `dimension-mismatch`
   on any divergence.
3. For each `(r, c)` in the range, dispatch internally to the same
   logic `xlsx:set-cell-value` uses — but **batch** the recalc:
   call `engine.addCell` per cell, then a **single** `engine.recalc()`
   at the end. This is the §17 perf-budget hot path for bulk paste.
4. Strings starting with `=` are **values, not formulas** in this
   command (consistent with §1; the agent must use a separate command
   for bulk formulas in P0). Document loudly.
5. Dirty `xl/worksheets/sheetN.xml`; intern strings into the SST as
   needed.
6. Diff: one `cell-updated` per cell whose value changed (skip
   no-ops); plus dependent recalc changes.

### OOXML impact

- Dirties: `xl/worksheets/sheetN.xml`, possibly
  `xl/sharedStrings.xml`.
- Creates: nothing.

### Excel-parity (per `prompt.md` line 407)

This is the "CSV paste" surface. Excel's clipboard paste-as-values
follows the same rectangular-must-match-target semantics; we
mirror it.

### `precheck`

| `reason`             | When                                     | `suggestedFix`                                             |
| -------------------- | ---------------------------------------- | ---------------------------------------------------------- |
| `unknown-sheet`      | sheet missing                            | sheet name list                                            |
| `invalid-range`      | range malformed                          | "Use 'A1:C3' shape; whole rows/columns not supported here" |
| `dimension-mismatch` | values dims don't match range            | reports expected vs actual dimensions                      |
| `merge-overlap`      | range overlaps a merged region partially | "Call xlsx:unmerge-cells first or shrink range"            |
| `cell-cap-exceeded`  | row × col > 100,000 (configurable cap)   | "Split into smaller batches"                               |

`merge-overlap` is per `EC-S3`: a range that fully contains or fully
sits outside every merge in the sheet is fine; partial overlap is
rejected.

### Example

```json
{
  "type": "xlsx:set-range-values",
  "payload": {
    "sheet": "Q1",
    "range": "A1:B2",
    "values": [
      ["name", "score"],
      ["Alex", 42]
    ]
  }
}
```

Before: A1:B2 blank. After: header + one row written. Diff: 4
`cell-updated` changes plus any dependent recalc.

### Inverse

A `xlsx:set-range-values` with `values` = the matrix of `before`
values for the same range. Cells that were blank before re-emit as
`null`. The property test asserts the round-trip value-by-value.

### Edge cases cited

- `EC-S1`/`EC-S2`: large sparse ranges OK up to the cap.
- `EC-S3`: partial-merge overlap rejected.
- `EC-V1`: long-string spill is a renderer concern, not a model
  concern; this command does not validate display width.

---

## 4. `xlsx:set-cell-format`

```typescript
type SetCellFormatPayload = {
  sheet: string;
  range: string; //  A1 range OR single cell
  format: Partial<CellFormat>; //  patch; absent fields = leave unchanged
};
```

### Behaviour

1. Resolve sheet and range.
2. For each cell in the range:
   - Read current `styleId` → resolve to a `CellFormat`.
   - Merge with `payload.format` patch: per-leaf `undefined` =
     leave unchanged, `null` (where allowed by sub-type) = clear.
   - Compute the **content hash** of the merged format.
   - Look up an existing `cellXfs` entry by content hash; reuse if
     found, otherwise mint a new `xfId` and append.
   - Set `cell.styleId = newXfId`.
3. Dirty `xl/worksheets/sheetN.xml` and `xl/styles.xml`.
4. Diff: one `format-updated` change per cell whose effective style
   hash actually changed (skip no-ops to keep diffs honest).

### OOXML impact

- Dirties: `xl/worksheets/sheetN.xml` (for the changed `s` attribute
  per cell), `xl/styles.xml` (for new `cellXfs` entries).
- Creates: nothing structural.

The **content-hashed style id** scheme is what makes the bulk-format
case cheap: a `format-range A1:Z1000 with bold:true` that promotes
1000 cells to "bold + their existing other-fields" only allocates a
small constant number of new `xfId`s — one per pre-existing style
that didn't already include `bold`.

### Excel-parity (per `prompt.md` line 398)

`CellFormat` covers every dimension from
[`feature-scope.md`](feature-scope.md) "Cell editing":
font/fill/border/alignment/number-format. The patch semantics match
how the renderer's format-bar buttons emit commands (`bold` toggle
sends `{ font: { bold: true } }`, not the full font object).

### `precheck`

| `reason`           | When                                  | `suggestedFix`                                      |
| ------------------ | ------------------------------------- | --------------------------------------------------- |
| `unknown-sheet`    | sheet missing                         | sheet list                                          |
| `invalid-range`    | range malformed                       | A1 range hint                                       |
| `invalid-format`   | malformed `format` (e.g. bad colour)  | "color must be 6-hex-digit RRGGBB without '#'"      |
| `unknown-style-id` | a referenced number-format id missing | "Use a built-in id 0..49 or a custom format string" |

### Example

```json
{
  "type": "xlsx:set-cell-format",
  "payload": {
    "sheet": "Q1",
    "range": "A1:E1",
    "format": { "font": { "bold": true }, "fill": { "color": "FFEEAA" } }
  }
}
```

Before: row 1 default style. After: row 1 bold with cream
background. Diff: 5 `format-updated` changes (one per cell), plus a
`style-added` change if a new `cellXfs` entry was appended.

### Inverse

`xlsx:set-cell-format` with `format` = a patch that **restores** each
cell's previous format. Because the patch is per-cell, the inverse
is materialised as **N** sub-payloads internally; the bus stores the
list and replays in order. (For uniform-before ranges this collapses
to one inverse.)

### Edge cases cited

- `EC-S6`: SST/inline string mix — irrelevant here; this command
  doesn't touch values.
- `EC-V2` (date format on a string): the format applies; renderer
  handles the "render the literal string" rule. Not rejected.

---

## 5. `xlsx:insert-row`

```typescript
type InsertRowPayload = {
  sheet: string;
  at: number; //  1-based, like A1; insertion is BEFORE this row
  count: number; //  >= 1
};
```

### Behaviour

1. Resolve sheet; reject `unknown-sheet`.
2. Validate `at >= 1`, `count >= 1`, `at + count - 1 <= 1048576`
   (Excel max row).
3. Walk `sheet.cells`: for every cell with `row >= at - 1`
   (0-based internal), shift `row += count`.
4. Walk `sheet.merges`: for every merge that overlaps the
   insertion band, expand `r1 += count`. For merges whose top is at
   or below `at - 1`, also shift `r0 += count`. Per `EC-S3`: a merge
   that the insertion **partially crosses** (insertion lands inside,
   not at the boundary, AND the merge would otherwise be split)
   triggers `merge-boundary-crossed` precheck failure.
5. For every formula in the workbook: call
   `references.adjustForInsertRow(ref, sheet, at - 1, count)`.
   Re-emit formula text via `serializeRangeRef`. Re-parse and call
   `engine.addCell` to update the dep graph.
6. Adjust conditional formats, hyperlinks, comments, defined names
   that target the affected sheet via the same `adjustForInsertRow`
   path.
7. Insert blank `<row>` placeholders into the sheet's
   `<sheetData>` (the serializer drops empty rows; the typed model
   doesn't actually need a placeholder, but the row index map gets
   a `count` bump).
8. Call `engine.recalc()`.
9. Diff: one `rows-inserted` summary change `{ at, count }` plus
   per-cell shift summaries (`shifted_cells: number`); plus
   recalc-side `cell-updated` changes for any dependent whose value
   changed.

### OOXML impact

- Dirties: `xl/worksheets/sheetN.xml`.
- Indirectly dirties any cross-sheet formula references → potentially
  every other `xl/worksheets/sheetM.xml`.
- May dirty `xl/comments*.xml`, `xl/sharedStrings.xml` (no), and
  anywhere else a row reference lives (defined names, conditional
  formats).
- Untouched parts (`xl/styles.xml`, every chart and drawing,
  `xl/sharedStrings.xml`, etc.) round-trip byte-identically.

### Excel-parity (per `prompt.md` line 399)

Excel's "Insert rows above" is exactly this: shift downward,
absolute refs adjust (not preserved), hyperlinks/comments anchored
to shifted cells follow, formulas re-emit with new ref text. We
match.

### `precheck`

| `reason`                 | When                                                   | `suggestedFix`                                                                         |
| ------------------------ | ------------------------------------------------------ | -------------------------------------------------------------------------------------- |
| `unknown-sheet`          | missing sheet                                          | sheet list                                                                             |
| `invalid-position`       | `at < 1` or `at > 1048576`                             | row range hint                                                                         |
| `invalid-count`          | `count < 1` or would overflow Excel max row            | "count must be ≥ 1 and at + count − 1 ≤ 1048576"                                       |
| `merge-boundary-crossed` | per `EC-S3` — insertion would split a merge mid-region | "Insert above the merge top, below the merge bottom, or call xlsx:unmerge-cells first" |

### Example

```json
{
  "type": "xlsx:insert-row",
  "payload": { "sheet": "Q1", "at": 3, "count": 2 }
}
```

Before: rows 1–10 populated. After: rows 1–2 unchanged, rows 3–4
blank (newly inserted), original rows 3–10 now 5–12. Diff:
`rows-inserted` summary plus `formula-updated` for every formula
whose ref shifted.

### Inverse

`xlsx:delete-row` with `at` and `count` from the original payload.
The property test seeds a workbook, applies `insert-row`, applies
the inverse, and asserts cell-by-cell equality with the original
snapshot. Because `before` is captured up-front, even formulas that
were rewritten by the insert come back to their original text on
the inverse.

### Edge cases cited

- `EC-R1` (insert above a referenced cell): the reference shifts
  per Excel. Tested.
- `EC-R5` (3D refs): not rewritten by this handler; the lexer
  preserves them verbatim.
- `EC-S2` (1,048,576-row sheet): explicit invalid-count guard.
- `EC-S3` (merged cell crossing): rejected via precheck.

---

## 6. `xlsx:insert-column`

```typescript
type InsertColumnPayload = {
  sheet: string;
  at: number; //  1-based column index (A=1, B=2, ...)
  count: number;
};
```

### Behaviour

Mirror of §5 with rows replaced by columns:

1. Resolve sheet; validate `at`, `count` (Excel max column
   `XFD = 16384`).
2. Shift every cell with `col >= at - 1` by `count`.
3. Adjust merges (analogous to row case; per `EC-S3` mid-merge
   insertion is rejected).
4. Adjust every formula's column refs via
   `adjustForInsertColumn`.
5. Adjust comments, hyperlinks, conditional formats, defined names.
6. Adjust `<col>` width definitions in
   `xl/worksheets/sheetN.xml`'s `<cols>`: split any `<col>` whose
   `min..max` straddles `at` and shift the rest.
7. `engine.recalc()`; diff includes `columns-inserted` summary.

### OOXML impact

- Dirties: `xl/worksheets/sheetN.xml` for the affected sheet plus
  every sheet that has cross-sheet column refs to it.
- Untouched parts byte-preserved.

### Excel-parity (per `prompt.md` line 400)

Same model as §5 — Excel "Insert columns to the left" maps directly.

### `precheck`

Same shape as §5 with `XFD = 16384` as the max-column ceiling.

### Example

```json
{
  "type": "xlsx:insert-column",
  "payload": { "sheet": "Q1", "at": 2, "count": 1 }
}
```

Before: cols A–E populated. After: A unchanged, B blank,
original B–E now C–F. Diff: `columns-inserted` summary plus
ref-shift changes.

### Inverse

`xlsx:delete-column` with the same `at` and `count`.

### Edge cases cited

`EC-R1`/`EC-R3`/`EC-R5`/`EC-S3` — same set as §5 transposed.

---

## 7. `xlsx:delete-row`

```typescript
type DeleteRowPayload = {
  sheet: string;
  at: number; //  1-based; rows at..at+count-1 are removed
  count: number;
};
```

### Behaviour

1. Resolve sheet; validate `at`, `count` (must be inside used range).
2. **Collect** the cells about to be deleted: for every formula in
   the workbook whose `dependencies` reference a cell or range
   inside the deletion band, mark its handler-side `casualties` set
   so the diff can list `referenced-cell-deleted` summaries per
   `EC-R2`.
3. Drop every cell with `at - 1 <= row < at - 1 + count`.
4. For every cell with `row >= at - 1 + count`, shift `row -= count`.
5. Adjust merges (a merge fully inside the deletion band is dropped;
   a merge that straddles the boundary triggers
   `merge-boundary-crossed` precheck). Merges entirely below the
   band shift up.
6. For every formula: call `adjustForDeleteRow`. The helper either
   returns an adjusted ref or `Errors[CellErrorKind.REF]` for refs
   that fall entirely inside the deletion band. Where `#REF!` is
   returned, the formula text is rewritten with a literal `#REF!`
   token (matching Excel's behaviour and per `EC-R2`/`EC-F4`).
7. Adjust comments, hyperlinks, conditional formats, defined names
   (per `EC-R3` defined names referencing a deleted column/row are
   rewritten with `#REF!` and the diff lists them).
8. `engine.recalc()`; cells that now evaluate to `#REF!` get their
   `cachedValue` updated and emit `cell-updated` changes.
9. Diff: `rows-deleted` summary `{ at, count, deleted_cells: N,
shifted_cells: M }`; plus per-cell `referenced-cell-deleted`
   summaries; plus formula updates.

### OOXML impact

- Dirties: every sheet whose formulas referenced the affected band.
- Removes cells from `xl/worksheets/sheetN.xml`.
- Strings whose only reference was a deleted cell may become unused
  in the SST; we **do not** GC the SST in P0 (preserve unused
  entries — Excel does the same).

### Excel-parity (per `prompt.md` line 401)

Excel's "Delete rows" applies the same shift+rewrite logic. Our
extra: the structured `referenced-cell-deleted` summary in the
diff. Excel silently rewrites; we surface.

### `precheck`

| `reason`                   | When                                      | `suggestedFix`                                            |
| -------------------------- | ----------------------------------------- | --------------------------------------------------------- |
| `unknown-sheet`            | missing sheet                             | sheet list                                                |
| `invalid-position`         | `at < 1`                                  | "use 1-based row index"                                   |
| `invalid-count`            | `count < 1` or `at + count - 1 > 1048576` | bounds hint                                               |
| `merge-boundary-crossed`   | deletion would split a merge              | "Delete the entire merge band or call xlsx:unmerge-cells" |
| `would-empty-required-row` | (reserved — currently unused in P0)       | —                                                         |

### Example

```json
{
  "type": "xlsx:delete-row",
  "payload": { "sheet": "Q1", "at": 5, "count": 1 }
}
```

Before: row 5 contains `[10, 20, 30]`; cell `=A5` exists in `B1`.
After: row 5 deleted, original rows 6–N shift up; `B1`'s formula
becomes `=#REF!` with `cachedValue = #REF!`. Diff includes
`rows-deleted` summary AND `referenced-cell-deleted` for `B1`.

### Inverse

`xlsx:insert-row` with the same `at` / `count`, **then** a batch
`xlsx:set-range-values` (per §3) restoring the deleted cells'
values, **then** `xlsx:set-cell-formula` for any formula cell.
Because the inverse is multi-step, the property test stores the
captured "deleted slice" in the mutation's `before` payload and the
inverse handler dispatches the sequence atomically. This is the
hardest-shaped inverse in the P0 set; it is the one most worth
the property-test investment.

### Edge cases cited

- `EC-R2` (delete a referenced cell): handled per spec.
- `EC-R3` (defined-name range deletion): handled.
- `EC-S3` (merge-boundary cross): rejected via precheck.
- `EC-F4` (delete sheet referenced by a formula): not this command —
  see `xlsx:delete-sheet` (out of P0).

---

## 8. `xlsx:delete-column`

```typescript
type DeleteColumnPayload = {
  sheet: string;
  at: number;
  count: number;
};
```

### Behaviour

Mirror of §7 with columns. Same casualty-collection, same
`#REF!` rewriting, same precheck shape with `XFD = 16384` as the
column ceiling.

### OOXML impact

- Same as §7 transposed; additionally the `<col>` width entries in
  `<cols>` shift / split to compensate.

### Excel-parity (per `prompt.md` line 402)

Same as §7 for the column axis.

### Example

```json
{
  "type": "xlsx:delete-column",
  "payload": { "sheet": "Q1", "at": 2, "count": 1 }
}
```

Before: cols A–E. After: B deleted; original C–E now B–D. Refs to
column B become `#REF!`; refs to C–E shift down by one.

### Inverse

`xlsx:insert-column` + restoration sequence (mirror of §7).

### Edge cases cited

`EC-R2`/`EC-R3`/`EC-S3` — same set as §7 transposed.

---

## 9. `xlsx:merge-cells`

```typescript
type MergeCellsPayload = {
  sheet: string;
  range: string; //  A1 range, e.g. "A1:C3"
};
```

### Behaviour

1. Resolve sheet and range.
2. Validate the range is a proper rectangle ≥ 2 cells.
3. Check overlap with existing merges: any partial overlap rejects
   `overlap-with-existing-merge`. Full containment of an existing
   merge inside the new range is also rejected (Excel allows this
   but we want the agent to call `xlsx:unmerge-cells` first to make
   the intent explicit; document loudly).
4. **Preserve only the anchor cell's value**: every non-anchor cell
   in the range has its value cleared (Excel parity — merging
   discards non-anchor values silently; we surface a
   `cells-cleared-by-merge` diff entry per cell).
5. Add the merge to `sheet.merges` (the `IndexedRanges<Merge>`
   structure from [`analysis.md`](analysis.md) §1.1).
6. Dirty `xl/worksheets/sheetN.xml`.
7. Diff: `merge-added` change `{ range }` + per-cleared-cell
   `cell-updated` summaries.

### OOXML impact

- Dirties: `xl/worksheets/sheetN.xml` (writes `<mergeCells>` with
  the new `<mergeCell ref="A1:C3"/>`).

### Excel-parity (per `prompt.md` line 403)

Excel's "Merge & Center" with values discarded — we match the
discard side; the "center" alignment is a separate format command
(`xlsx:set-cell-format` with `alignment.horizontal: "center"`).

### `precheck`

| `reason`                      | When                                       | `suggestedFix`                             |
| ----------------------------- | ------------------------------------------ | ------------------------------------------ |
| `unknown-sheet`               | missing sheet                              | sheet list                                 |
| `invalid-range`               | range malformed or single cell             | "Range must be ≥ 2 cells (e.g. 'A1:B2')"   |
| `overlap-with-existing-merge` | partial overlap, or contains another merge | "Call xlsx:unmerge-cells on \[refs] first" |

### Example

```json
{
  "type": "xlsx:merge-cells",
  "payload": { "sheet": "Q1", "range": "A1:C1" }
}
```

Before: A1=`"Title"`, B1=C1=blank. After: A1:C1 merged with anchor
value `"Title"`. Diff: `merge-added` plus zero `cell-updated` (B1,
C1 were already blank).

### Inverse

`xlsx:unmerge-cells` with the same `range`. The inverse-after-clear
case (where merging cleared B1=`"x"` and C1=`"y"`) requires an
additional `xlsx:set-cell-value` per cleared cell to restore — same
multi-step pattern as `delete-row`.

### Edge cases cited

- `EC-S3` (merge crossing inserted band): structural commands
  reject merging-related cases; this command itself doesn't fire
  EC-S3.
- `EC-V3` (window-edge merge): renderer concern, not model.

---

## 10. `xlsx:unmerge-cells`

```typescript
type UnmergeCellsPayload = {
  sheet: string;
  range: string; //  must exactly match an existing merge
};
```

### Behaviour

1. Resolve sheet and range.
2. Find a merge in `sheet.merges` whose ref **exactly equals** the
   payload range (we don't unmerge "the merge containing this cell"
   in P0 — the agent must specify the merge precisely; the diff
   from `xlsx:merge-cells` is structured so this is easy).
3. Remove the merge.
4. Dirty `xl/worksheets/sheetN.xml`.
5. Diff: `merge-removed` `{ range }`.

### OOXML impact

- Dirties: `xl/worksheets/sheetN.xml` (drops the `<mergeCell>`).

### Excel-parity (per `prompt.md` line 404)

Excel's "Unmerge cells". The non-anchor cells re-appear blank
(matching their state at merge time — we did not retain the cleared
values).

### `precheck`

| `reason`          | When                               | `suggestedFix`                            |
| ----------------- | ---------------------------------- | ----------------------------------------- |
| `unknown-sheet`   | missing sheet                      | sheet list                                |
| `invalid-range`   | range malformed                    | A1 range hint                             |
| `merge-not-found` | no exact-match merge at that range | "Use xlsx_inspect to list current merges" |

### Example

```json
{
  "type": "xlsx:unmerge-cells",
  "payload": { "sheet": "Q1", "range": "A1:C1" }
}
```

Before: A1:C1 merged. After: three independent cells, A1 keeps the
anchor value, B1 / C1 blank.

### Inverse

`xlsx:merge-cells` with the same `range`. Symmetric and clean.

### Edge cases cited

- `EC-S3`: irrelevant here.
- `EC-V3`: renderer concern.

---

## 11. `xlsx:add-sheet`

```typescript
type AddSheetPayload = {
  name: string;
  at?: number; //  0-based insert position; default = append
};
```

### Behaviour

1. Validate `name`: 1–31 chars, no `[`/`]`/`*`/`?`/`:`/`/`/`\\`,
   not the reserved `History` (Excel rule), not a duplicate of an
   existing sheet (case-insensitive comparison).
2. Mint a new sheet id (next free integer).
3. Construct an empty `Sheet` with default dimensions (1 row, 1
   col), no merges, no comments, no formulas.
4. Insert at `at ?? sheets.length` in `workbook.sheets`.
5. Mint an OOXML part path: `xl/worksheets/sheet{id}.xml`.
6. Update `xl/workbook.xml`'s `<sheets>` list and
   `xl/_rels/workbook.xml.rels` to add the new sheet rel.
7. Update `[Content_Types].xml` to add the
   `application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml`
   override for the new part.
8. Diff: `sheet-added` `{ name, at, id }`.

### OOXML impact

- Creates: `xl/worksheets/sheet{N}.xml` (a near-empty sheet).
- Dirties: `xl/workbook.xml`, `xl/_rels/workbook.xml.rels`,
  `[Content_Types].xml`.
- Untouched parts: every existing sheet, styles, SST, etc.

### Excel-parity (per `prompt.md` line 405)

Standard "Insert sheet" behaviour. Tab colour, freeze panes, etc.
default to absent and are set by separate format-style commands
(out of P0 for tab-colour authoring).

### `precheck`

| `reason`           | When                             | `suggestedFix`                                   |
| ------------------ | -------------------------------- | ------------------------------------------------ |
| `invalid-name`     | name violates Excel naming rules | "Use 1–31 chars; no \\ / ? \* \[ ] : characters" |
| `duplicate-name`   | name (case-insensitive) exists   | "Pick a unique name; existing: \[sheet names]"   |
| `invalid-position` | `at` out of `[0, sheets.length]` | "Position must be between 0 and \[count]"        |

### Example

```json
{
  "type": "xlsx:add-sheet",
  "payload": { "name": "Q2", "at": 1 }
}
```

Before: workbook has \[Q1]. After: \[Q1, Q2]. Diff: `sheet-added`.

### Inverse

`xlsx:delete-sheet` (out of P0 deliberately, per
[`analysis-agent-patterns.md`](analysis-agent-patterns.md) §8.4 — sheet
deletion is unrecoverable in-session). For property-test purposes,
the test fixture invokes a private `removeSheet` helper that maps
1:1 onto a future `xlsx:delete-sheet` payload; the property still
holds.

### Edge cases cited

- `EC-O1`/`EC-O2`: round-trip integrity for the new sheet plus all
  others. Tested.

---

## 12. `xlsx:rename-sheet`

```typescript
type RenameSheetPayload = {
  name: string; //  current name
  newName: string;
};
```

### Behaviour

1. Resolve current sheet by `name`; reject `unknown-sheet` on miss.
2. Validate `newName` per the same rules as `xlsx:add-sheet`
   (length, charset, uniqueness — the **current** sheet is
   excluded from the duplicate check so renaming to the same name
   is a no-op-but-valid bump).
3. Update `sheet.name`.
4. **Rewrite cross-sheet references**: walk every formula in every
   sheet; for each formula whose token stream contains a sheet
   prefix matching the old name (bare or quoted form per
   `EC-R4`), rewrite the prefix and re-parse the formula. Update
   the dep graph.
5. Update defined names that reference the renamed sheet.
6. Per `EC-R5`: 3D references (`Sheet1:Sheet3!A1`) are **not**
   rewritten; the lexer preserves them verbatim. The diff records a
   warning for each 3D ref that may be stale.
7. Dirty `xl/workbook.xml` (the `<sheet name="...">` attribute), the
   modified sheet's `xl/worksheets/sheetN.xml` (no actual change to
   that file, but we keep the dirty flag for audit), and every
   sheet whose formulas changed.
8. Diff: `sheet-renamed` `{ from, to }` plus per-formula
   `formula-updated` summaries.

### OOXML impact

- Dirties: `xl/workbook.xml`, plus every sheet whose formulas
  needed rewriting.
- Untouched parts: `xl/sharedStrings.xml` (the sheet name is **not**
  in the SST), styles, charts, etc.

### Excel-parity (per `prompt.md` line 406)

Excel rewrites every formula reference on rename. Quoted vs bare
form is normalised per `EC-R4`. We match.

### `precheck`

| `reason`         | When                                       | `suggestedFix`                       |
| ---------------- | ------------------------------------------ | ------------------------------------ |
| `unknown-sheet`  | `name` not in workbook                     | sheet list                           |
| `invalid-name`   | `newName` violates rules                   | naming rules hint                    |
| `duplicate-name` | another sheet has `newName` (case-insens.) | "Pick a name not in: \[sheet names]" |

### Example

```json
{
  "type": "xlsx:rename-sheet",
  "payload": { "name": "Sheet2", "newName": "Inputs" }
}
```

Before: Sheet1 contains `=Sheet2!A1`. After: rename to `Inputs`,
`Sheet1` now contains `=Inputs!A1`. Diff: `sheet-renamed` plus
one `formula-updated` for the rewritten cell.

### Inverse

`xlsx:rename-sheet` swapping `name` and `newName`. The property
test asserts that the formula text (and dep graph) round-trips
exactly, including quoted-form refs.

### Edge cases cited

- `EC-R4` (cross-sheet rename): handled.
- `EC-R5` (3D refs): preserved verbatim with diff warning.

---

## 13. `xlsx:add-comment`

```typescript
type AddCommentPayload = {
  sheet: string;
  ref: string; //  A1 single cell
  text: string;
  author: string;
};
```

### Behaviour

1. Resolve sheet and ref.
2. If a comment already exists on this cell, reject
   `comment-exists` (replacement requires `xlsx:edit-comment` —
   deferred — or delete-then-add).
3. Mint a new `commentId` (next free integer per sheet's
   `xl/comments{N}.xml`).
4. Append a `<comment>` to `sheet.comments` with `{ ref, text,
author, threadId: undefined }` (threaded comments — replies — are
   a separate command, also deferred to P1).
5. Append the author to the sheet's `<authors>` list if absent.
6. If `xl/comments{N}.xml` does not exist yet for this sheet,
   create it (and add the relationship in
   `xl/worksheets/_rels/sheetN.xml.rels` and the override in
   `[Content_Types].xml`). Mirrors the `ensureCommentsPart` pattern
   in [`spec/docx/serializer.md`](../docx/serializer.md).
7. Dirty `xl/comments{N}.xml`, possibly
   `xl/worksheets/_rels/sheetN.xml.rels` and
   `[Content_Types].xml`.
8. Diff: `comment-added` `{ sheet, ref, commentId, author }`.

### OOXML impact

- Dirties: `xl/comments{N}.xml`.
- May create: `xl/comments{N}.xml`, the corresponding `_rels`
  entry, and the content-types override (only when the sheet had
  no prior comments).
- Untouched parts: the sheet itself
  (`xl/worksheets/sheet{N}.xml`) — comments live in their own
  part. This is the "modified parts versus untouched parts" rule
  per `EC-O2` paying off.

### Excel-parity (per `prompt.md` line 408)

Excel's classic comment surface — the threaded-comments rewrite
(2018+ "modern" comments) lives in `xl/threadedComments/` and is a
P1 deferral. P0 ships the classic note shape. Round-trip for
existing threaded comments is preserved per
[`feature-scope.md`](feature-scope.md) "Comments / notes".

### `precheck`

| `reason`         | When                           | `suggestedFix`                               |
| ---------------- | ------------------------------ | -------------------------------------------- |
| `unknown-sheet`  | missing sheet                  | sheet list                                   |
| `invalid-ref`    | ref malformed                  | single A1 cell hint                          |
| `comment-exists` | the cell already has a comment | "Call xlsx:delete-comment first then re-add" |
| `empty-text`     | `text === ""`                  | "Comment text must be non-empty"             |
| `empty-author`   | `author === ""`                | "Author is required"                         |

### Example

```json
{
  "type": "xlsx:add-comment",
  "payload": { "sheet": "Q1", "ref": "B7", "text": "Verify with finance", "author": "OfficeAI" }
}
```

Before: B7 has no comment. After: B7 carries a comment by
"OfficeAI". Diff: `comment-added` (and `part-added` for
`xl/comments1.xml` if the sheet had no prior comments).

### Inverse

`xlsx:delete-comment` with `{ sheet, commentId }` (the delete
command is in scope for P1 reply/resolve/delete per
[`feature-scope.md`](feature-scope.md) "Comments / notes" but not
listed in the 13 P0 commands; the property test calls a private
`removeComment` helper that maps 1:1 to the future payload).

### Edge cases cited

- `EC-O1`/`EC-O2`: comments part round-trip. Tested.
- `EC-I4` (non-UTF-8 comments): not relevant here — we always
  emit UTF-8.

---

## Diff format per command — summary

Every handler returns a `DocumentDiff` whose `changes` are sorted
by `path` for deterministic serialisation. The kinds used:

| Kind                            | Used by                                                                                          |
| ------------------------------- | ------------------------------------------------------------------------------------------------ |
| `cell-updated`                  | §1, §2, §3 (per cell), recalc side-effects of every command that mutates values                  |
| `formula-updated`               | §2, §5–§8, §12 (formula text rewrite)                                                            |
| `format-updated`                | §4                                                                                               |
| `style-added`                   | §4 (when a new `xfId` is appended)                                                               |
| `rows-inserted`                 | §5                                                                                               |
| `columns-inserted`              | §6                                                                                               |
| `rows-deleted`                  | §7                                                                                               |
| `columns-deleted`               | §8                                                                                               |
| `referenced-cell-deleted`       | §7, §8 — per `EC-R2` / `EC-F4`                                                                   |
| `merge-added` / `merge-removed` | §9 / §10                                                                                         |
| `sheet-added`                   | §11                                                                                              |
| `sheet-renamed`                 | §12                                                                                              |
| `comment-added`                 | §13                                                                                              |
| `part-added`                    | any command that creates a new OOXML part (e.g. §11 sheet, §13 comments part)                    |
| `circular`                      | §2 (and any command whose recalc surfaces a cycle) — carries `meta.cycle: CellRef[]` per `EC-F1` |

`comment-added`, `merge-added`, etc. are emitted via the same
`Diff.change` construction the `@officeai/core` differ uses for
DOCX (see [`spec/docx/agent-commands.md`](../docx/agent-commands.md)
§"Diff format per command"); the renderer can share its diff-display
component across formats.

---

## Registration order in `registry.ts`

Per `acceptance-criteria.md` `G3`, every P0 handler is registered
in `packages/xlsx/src/commands/registry.ts`. The registration order
matches this doc's section order so the file reads like a numbered
table of contents:

```typescript
import { register } from "@officeai/core/command-bus";
import * as h from "./handlers";

export function registerXlsxCommands(bus: CommandBus<XlsxSnapshot>): void {
  bus.register(h.setCellValue); //  §1
  bus.register(h.setCellFormula); //  §2
  bus.register(h.setRangeValues); //  §3
  bus.register(h.setCellFormat); //  §4
  bus.register(h.insertRow); //  §5
  bus.register(h.insertColumn); //  §6
  bus.register(h.deleteRow); //  §7
  bus.register(h.deleteColumn); //  §8
  bus.register(h.mergeCells); //  §9
  bus.register(h.unmergeCells); //  §10
  bus.register(h.addSheet); //  §11
  bus.register(h.renameSheet); //  §12
  bus.register(h.addComment); //  §13
}
```

Out-of-scope commands (`xlsx:edit-pivot`, `xlsx:add-chart`,
`xlsx:set-defined-name`, `xlsx:run-macro` — see
[`edge-cases.md`](edge-cases.md) "Out-of-scope features that the
agent might attempt") are **not** registered. A dispatch with one
of those `type` values surfaces the bus's default
`no-handler-registered` rejection — exactly the structured "no" the
edge-cases doc promises.

---

## Phase 7 closure gates

What gates this doc's parent phase as "done":

1. Every command in §1–§13 has:
   - A typed payload in `packages/xlsx/src/commands/payloads.ts`.
   - A handler module in `packages/xlsx/src/commands/handlers/`.
   - An entry in `registry.ts` matching the order above.
   - Vitest coverage: happy path + at least one edge case + the
     `precheck` rejection path.
   - An inverse documented per the "Inverse" subsection.
2. The property test in `handlers.test.ts` asserts
   `apply(redo) ∘ apply(undo) === identity` on a generator-built
   workbook for every command.
3. The diff schema kinds enumerated in the summary table are
   defined in `packages/xlsx/src/diff/types.ts` and accepted by
   the renderer's diff-display component (shared with DOCX).
4. The MCP layer ([`analysis-agent-patterns.md`](analysis-agent-patterns.md)
   §8) wires `xlsx_apply` to dispatch any `xlsx:*` typed command
   end-to-end; the dedicated tools (`xlsx_set_cell`,
   `xlsx_set_formula`, …) collapse to the equivalent `xlsx_apply`
   payload internally so there is one code path.

When all four are green, [`acceptance-criteria.md`](acceptance-criteria.md)
`G3` flips, `docs/build-log/xlsx.md` records the close, and the
phase is done.
