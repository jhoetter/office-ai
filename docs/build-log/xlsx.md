# XLSX Build Log

> Live log of decisions, deviations from spec, and known issues for the
> XLSX build (`packages/xlsx`). Mirrors the shape of
> [`docx.md`](./docx.md). Each phase appends; nothing is rewritten in
> place.

## Decisions

| Date (UTC) | Decision                                                                                                                                  | Rationale                                                                                                                                                                                      |
| ---------- | ----------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| 2026-04-18 | Adopt SheetJS Community Edition (Apache 2.0) as the cell-layer parser/serializer fallback                                                 | Battle-tested across thousands of real workbooks; lets us focus engineering on agent + formula engine instead of OOXML grunt-work.                                                             |
| 2026-04-18 | OOXML is the source of truth; the in-memory model is a working surface                                                                    | Same byte-preservation discipline as `@officeai/docx`. Untouched parts must round-trip bit-identical via `OoxmlContainer.partHashes`.                                                          |
| 2026-04-18 | Discriminated-union typed payloads, schema-validated at the bus boundary                                                                  | Differs from Univer's stringly-typed `unknown` params; trades some flexibility for editor + agent type safety.                                                                                 |
| 2026-04-18 | Sync-only formula evaluator. No iterative calc. Circular refs surface as a structured `#REF!`                                             | Matches the 80% scope; iterative calc is rare in practice and adds complexity that derails the night shift.                                                                                    |
| 2026-04-18 | Renderer is a virtualized DOM grid (no canvas)                                                                                            | Mirrors how `@officeai/docx` renders. Keeps every interaction routable through ProseMirror-style command dispatch.                                                                             |
| 2026-04-18 | Sparse cell store keyed by `${row}:${col}` strings (Phase 5)                                                                              | Agent workloads write in scattered patterns; `Map<string, Cell>` is faster for `delete` and avoids wasted intermediate row objects.                                                            |
| 2026-04-18 | Phase 4 ships a thin model: typed `Sheet` (id/name/index/path/state/kind) + opaque parts; cells exposed only via the SheetJS escape hatch | Phase 5 needs the typed cell model anyway; shipping the round-trip oracle first lets us verify byte-preservation against the synthetic corpus before any commands exist that could falsify it. |
| 2026-04-18 | Phase 5 ships only the 5 P0 commands that don't depend on the formula engine or new OPC parts; the other 8 commands defer to Phase 6/7    | Lets us land the cell model + bus wiring + dirty-sheet serializer behind a passing test suite, instead of stubbing 8 handlers without recalc/style/comment infrastructure to back them.        |
| 2026-04-18 | String cells written by Phase 5 commands are emitted inline (`bookSST: false`)                                                            | Avoids touching `xl/sharedStrings.xml`, which keeps that part byte-identical for partially-edited workbooks. SST coalescing is a Phase 6+ concern.                                             |

## Deviations from spec

_(none yet)_

### Resolved deviations

_(none yet)_

## Phase log

### Phase 0 — scaffold (2026-04-18)

- New package `@officeai/xlsx@0.1.0` mirroring `@officeai/docx`'s shape.
- Architecture linter updated: `xlsx → core` only; `agent`,
  `integration-tests`, `web` may depend on `xlsx`; `react/react-dom/next`
  forbidden in `xlsx` (headless contract).
- `spec/xlsx/README.md` upgraded from a stub to a real index with a
  status table tracking each Phase-2 doc.
- All `src/` subdirs are `.gitkeep` placeholders. `pnpm verify` green.

### Phase 1 — analysis (2026-04-18)

- Four parallel clean-room analyses checked into `spec/xlsx/`:
  `analysis-univer-core.md` (567 lines), `analysis-univer-formula.md`
  (1156 lines), `analysis-sheetjs.md` (897 lines),
  `analysis-agent-patterns.md` (700 lines). Total ~3,300 lines.
- `analysis.md` synthesises into a KEEP / DIFFER / IMPROVE table and a
  risk register that drives Phase 4-onwards decisions.

### Phase 2 — spec (2026-04-18)

- 10 spec docs landed (~9.7k lines) covering feature scope,
  acceptance criteria, edge cases, document model, OOXML mapping,
  parser, serializer, renderer, formula engine, and agent commands.
- Index `spec/xlsx/README.md` flips every entry to "landed (P2)".

### Phase 3 — fixtures (2026-04-18)

- 6 synthetic fixtures generated by `scripts/generate-xlsx-fixtures.mjs`
  (SheetJS, Apache 2.0). Cover every P0 capability area:
  - `01-single-sheet-numbers.xlsx` — mixed cell value kinds
  - `02-multi-sheet.xlsx` — cross-sheet formula references
  - `03-formulas-basic.xlsx` — SUM/AVERAGE/COUNTIF/MAX/IF/VLOOKUP
  - `04-merged-and-formatted.xlsx` — merges + currency/percent formats
  - `05-comments-hyperlinks.xlsx` — comments + hyperlinks + tooltips
  - `06-large-grid.xlsx` — 1000 × 7 perf smoke
- Fixture inventory + to-collect real-world slots tracked in
  `fixtures/xlsx/MANIFEST.md`.
- `pnpm fixtures:xlsx` + `make fixtures-xlsx` (folded into
  `make fixtures`).

### Phase 4 — parser + serializer + roundtrip oracle (2026-04-18)

- Thin model: `XlsxSnapshot`, `XlsxWorkbook`, `Sheet`, `OpaquePart`,
  `XlsxDirtyFlags` (per `spec/xlsx/document-model.md`).
- `parseXlsx(buffer)`:
  - Loads `OoxmlContainer`, hashes every part for the byte-oracle.
  - Parses `xl/workbook.xml` with native `fast-xml-parser`: extracts
    `date1904`, root attributes, and the typed `<sheet>` list.
  - Resolves each sheet's part path via `xl/_rels/workbook.xml.rels`
    using a small `resolveTargetPath()` helper that handles `../`
    segments and absolute targets.
  - Pre-loads SheetJS in dense mode (`dense: true, cellFormula,
cellStyles, cellNF, cellDates, sheetStubs, bookVBA, xlfn`) and
    parks the resulting `WorkBook` on `XlsxWorkbook.sheetjs` for
    Phase 5 to consume.
  - Classifies every non-modeled part (themes, app/core props,
    metadata, future charts/drawings/VBA) as an `OpaquePart` keyed by
    full zip path with a SHA-256 hash for verification.
  - Surfaces typed errors (`zip-corruption`, `missing-workbook-part`,
    `missing-content-types`, `invalid-xml`, `invalid-workbook`,
    `missing-sheet-target`, `sheetjs-failure`).
- `serializeXlsx(snapshot)`:
  - Clones the container and re-emits via `OoxmlContainer.serialize()`
    — untouched parts round-trip byte-content-identical (the zip-archive
    bytes may differ; see `OoxmlContainer.serialize()` docstring).
  - Refuses to run when any dirty flag is set; Phase 5 will wire the
    per-part re-emission paths. The refusal is a typed
    `XlsxSerializeError("container-failed", ...)` with a message
    pointing at Phase 5.
- Tests:
  - `packages/xlsx/src/parser/parse.test.ts` — 11 tests, covers all 6
    fixtures plus negative cases (non-zip input, missing `workbook.xml`).
  - `packages/xlsx/src/serializer/serialize.test.ts` — 13 tests, drives
    the round-trip byte-oracle on every fixture twice (via
    `serializeXlsx` and via raw `container.serialize`).
  - `tests/roundtrip/xlsx/fixtures-roundtrip.test.ts` — 12 integration
    tests asserting the part-list and per-part hashes survive a full
    round-trip and that the typed sheet list is structurally equal
    after re-parse.
- Total Phase 4 test count: 36 (xlsx 24 + integration 12), all green.
  Full `pnpm verify` pipeline (format-check / lint / architecture /
  typecheck / test / build) is green.

### Phase 5 — typed cell model + first 5 P0 commands (2026-04-18)

**Scope shipped (5 of 13 P0 commands)**

The cell model + the value/structure-mutation commands that do not
depend on the formula engine. Specifically:

- `xlsx:set-cell-value` — single-cell write/clear with merge-anchor
  enforcement and a `formula-string` reject for `=`-prefixed strings.
- `xlsx:set-range-values` — rectangular bulk-paste; rejects on
  dimension mismatch, partial-merge overlap, and a 100k-cell cap.
- `xlsx:merge-cells` — adds a merge, clears non-anchor values,
  rejects partial overlaps with existing merges.
- `xlsx:unmerge-cells` — removes an exact-match merge.
- `xlsx:rename-sheet` — renames the typed `Sheet`, syncs the SheetJS
  workbook, and surgically patches `<sheet name="...">` in
  `xl/workbook.xml` while leaving every other byte untouched.
  Cross-sheet formula reference rewriting (per `EC-R4`) lands in
  Phase 7 with the formula engine.

**Deferred (8 of 13, with stable type names exported)**

| Command                              | Defers to | Why                                                       |
| ------------------------------------ | --------- | --------------------------------------------------------- |
| `xlsx:set-cell-formula`              | Phase 7   | Needs the lexer/parser/AST + recalc.                      |
| `xlsx:set-cell-format`               | Phase 7+  | Needs typed style table on `XlsxWorkbook`.                |
| `xlsx:insert-row` / `:insert-column` | Phase 7   | Reference adjustment requires the formula AST.            |
| `xlsx:delete-row` / `:delete-column` | Phase 7   | Same.                                                     |
| `xlsx:add-sheet`                     | Phase 6   | Needs `workbook.xml` + content-types + rels co-rewrite.   |
| `xlsx:add-comment`                   | Phase 6   | Needs `xl/commentsN.xml` emission + sheet-rel attachment. |

The deferrals are tracked in `packages/xlsx/src/commands/registry.ts`
and the `spec/xlsx/agent-commands.md` document is the canonical source
for the full P0 surface.

**Model upgrades (`packages/xlsx/src/model/`)**

- New `types`: `Cell`, `CellValue`, `CellErrorValue`, `CellErrorCode`,
  `Formula`, `MergedCell`. `Sheet` gains `cells: ReadonlyMap<string,
Cell>` (sparse, `${row}:${col}` keys, 0-based) and
  `merges: ReadonlyArray<MergedCell>`.
- New `refs.ts`: A1 ⇄ `{row, col}` plumbing — `parseA1`, `formatA1`,
  `parseRange`, `formatRange`, `colToLetter`, `letterToCol`,
  `cellKey`, `parseCellKey`, `rangeArea`, `rangesOverlap`. Strict
  validation against Excel's `XFD1048576` ceiling.

**Parser changes**

- SheetJS load is now done before sheet resolution so each typed
  `Sheet` can be hydrated with cells + merges in one pass.
- New `extractCellsAndMerges(ws)` walks the dense `!data` matrix +
  `!merges`, mapping SheetJS cell types (`n` / `s` / `b` / `d` / `e`
  / `z`) to the typed `CellValue` union. Numeric error codes from
  SheetJS (`0x07` = `#DIV/0!`, `0x17` = `#REF!`, …) are translated to
  the `CellErrorCode` strings.
- Empty cells and stub cells are not stored — the typed map stays
  sparse.
- Formulas are preserved verbatim as `Formula.text` (without the
  leading `=`) — Phase 5 commands never write formulas, but parsing
  preserves them so dirty-sheet round-trip keeps caller-authored
  formulas intact.

**Serializer changes**

- Phase 4's "any dirty flag → throw" has been replaced with a
  surgical, two-mode rewrite:
  1. **Dirty sheets** → for each sheet path in `dirty.sheets`, sync
     the typed cells + merges back onto the SheetJS WorkSheet
     (`sheet-sync.ts`), then call `XLSX.write(book, { bookSST: false,
… })` to emit a single workbook through SheetJS, load that
     emitted xlsx as a temporary `OoxmlContainer`, and copy only the
     dirty sheet's `xl/worksheets/sheetN.xml` bytes into the master
     container. Untouched sheets stay byte-identical.
  2. **Dirty workbook** → only used by `xlsx:rename-sheet`. We do an
     attribute-level regex patch over the existing `xl/workbook.xml`
     bytes, replacing only the `<sheet name="...">` attribute for
     each renamed sheet (matched by stable `r:id`). Every other byte
     in the workbook XML is preserved.
- `bookSST: false` is intentional: string cells written by Phase 5
  commands are emitted inline (`<c t="inlineStr">`) so the shared
  strings part is never disturbed. This means edited workbooks may
  grow slightly compared to a SheetJS-recompacted SST output, but
  preserves byte-stability of unrelated parts.
- `dirty.sharedStrings`, `dirty.styles`, `dirty.contentTypes`,
  `dirty.rels`, `dirty.comments`, `dirty.threadedComments`, and
  `dirty.sheetRels` continue to throw a typed
  `XlsxSerializeError("container-failed", …)` — those rewrites land
  with the deferred commands above.

**Bus integration**

- Each handler is a `CommandHandler<Payload, XlsxSnapshot>` consumed
  by `@officeai/core`'s shared `CommandBus`. Agent-sourced commands
  enter `pending`; human/system-sourced commands `approve` directly,
  matching the DOCX semantics.
- Validation is centralised in `commands/validation.ts`
  (`resolveSheet`, `parseCellRef`, `parseRangeRef`,
  `validateSheetName`, `assertUniqueSheetName`,
  `assertNotMergedNonAnchor`, `assertNotFormulaString`,
  `findContainingMerge`).
- `evolveSnapshot` + `mergeDirty` in `commands/helpers.ts` keep dirty
  flags monotonically additive and bump the snapshot revision on
  every successful apply.

**Tests landed**

- `packages/xlsx/src/commands/handlers.test.ts` — 23 tests across
  the 5 handlers + bus-integration smoke (agent pending → approve,
  invalid-ref handling, formula-string rejection, merge anchor
  enforcement).
- `packages/xlsx/src/serializer/serialize.test.ts` — extended to 14
  tests; new cases exercise the dirty-sheet rewrite path and the
  unsupported-flag guard.
- `tests/roundtrip/xlsx/commands-roundtrip.test.ts` — 5 integration
  tests proving the dispatched commands persist across `serialize →
parse` for value writes, range writes, merge/unmerge, and rename.
- Phase 5 totals: **48 unit tests in `@officeai/xlsx`** + **17 xlsx
  integration tests** (12 Phase 4 + 5 Phase 5), all green. Full
  `pnpm verify` pipeline (format / lint / architecture / typecheck /
  test / build) green.

**Known limitations / follow-ups**

- Cross-sheet formula references that name the renamed sheet are not
  rewritten. Workbooks with formulas like `=Expenses!A1` will break
  on recalc after `xlsx:rename-sheet`. Tracked for Phase 7.
- The SheetJS-based dirty-sheet emission may not preserve exotic
  worksheet features (custom XML inside the sheet, slicers, complex
  conditional formatting) on the dirty sheet. The other sheets in
  the workbook are untouched. Phase 6+ migrates to a native sheet
  XML emitter that surgically patches `<sheetData>` + `<mergeCells>`
  while preserving all other worksheet XML in place.

## 2026-04-18 — Phase 6: headless XlsxAgent (DocxAgent parity)

**Shipped**

- `packages/xlsx/src/agent/agent.ts` — `XlsxAgent` class implementing
  the same surface as `DocxAgent` (see `spec/shared/agent-api.md`):
  - **Read** — `getSnapshot`, `getApprovedSnapshot`, `listSheets`,
    `toMarkdown` (per-sheet, bounding-box-clipped), `getRange`
    (sparse projection of A1 ranges or whole sheets), `search`
    (substring / regex over string-typed cells, optional sheet
    filter).
  - **Write** — `applyCommand`, `applyCommands`. Both go through the
    shared `CommandBus`; `source: "agent"` stages as `pending`,
    `source: "human"`/`"system"` auto-approves.
  - **Diff & review** — `getDiff`, `getPendingMutations`,
    `approveMutation`, `rejectMutation`, `rollback`.
  - **I/O** — `fromBuffer`, `importFile`, `exportFile`. Re-uses the
    Phase 5 surgical serializer.
  - **Subscriptions** — `subscribe(listener)` for live mutation
    notifications.
- `packages/xlsx/src/agent/diff.ts` — `diffXlsxSnapshots(from, to)`.
  Sheet-matching by stable `sheetId`; cell-matching by
  `${row}:${col}`; merge-matching by exact rectangle. Emits typed
  `DiffChange` records (`node-inserted`, `node-deleted`,
  `node-updated`, `node-moved`).
- `packages/xlsx/src/agent/index.ts` — public re-exports; mirrored
  in `packages/xlsx/src/index.ts`.

**Tests landed**

- `packages/xlsx/src/agent/agent.test.ts` — 9 tests covering
  byte-preservation roundtrip, re-parse-after-edit fidelity,
  `listSheets`, `toMarkdown`, `getRange`, `search`, `getDiff`,
  agent-pending→approve flow, and headless (no-DOM) execution.
- `packages/xlsx/src/agent/diff.test.ts` — 6 tests covering insert,
  update, delete, rename, merge add/remove, and identity (no-op)
  paths.
- Phase 6 totals: **63 unit tests in `@officeai/xlsx`** + **17 xlsx
  integration tests** (carried from Phase 5), all green. Full `pnpm
verify` pipeline (format / lint / architecture / typecheck / test
  / build) green.

**Decisions**

- **DocxAgent parity over speed.** Method signatures, naming, and
  semantics deliberately mirror `DocxAgent` so the upcoming
  `office-agent` CLI / MCP server (Phase 8) can hold a single
  `DocumentAgent` interface and switch on `format`.
- **Sparse projection for `getRange`.** Range snapshots return only
  populated cells (row-major ordered) plus `(rows, cols)` for
  sizing. Caller densifies if needed. This keeps payloads tiny for
  large but mostly-empty sheets, which is the common LLM context
  shape.
- **Search scoped to string cells.** Numbers, booleans, and errors
  are skipped. Once the formula engine lands (Phase 7), we'll
  optionally include formula text and computed values.
- **Diff is structural, not pixel-perfect.** Style, conditional
  formatting, and comments aren't diffed yet — the first three lift
  once those models become typed (Phase 7+). For Phase 6 the diff
  covers everything the 5 implemented commands can mutate.

## 2026-04-18 — Phase 7a + 7b: formula engine foundation, lexer, parser, AST

Phase 7 (formula engine + 8 deferred P0 commands) is broken into ten
sub-phases (7a–7j) so each lands as an independently shippable
quality-gated commit. 7a + 7b are now in.

**Phase 7a — foundation (`tokens`, `errors`, `values`, `references`)**

- `packages/xlsx/src/formula/tokens.ts` — `TokenType` string union
  - `Token` interface; the immutable contract between lexer and
    parser.
- `packages/xlsx/src/formula/errors.ts` — Excel error kinds
  (`#DIV/0!`, `#NAME?`, `#VALUE!`, `#NUM!`, `#N/A`, `#REF!`,
  `#NULL!`, `#SPILL!`, `#CALC!`, `#CYCLE!`, `#GETTING_DATA`),
  interned `Errors` singleton table, `parseErrorLiteral`, and
  metadata-bearing factories `refWithCycle` /
  `refWithDeletedTarget` for the dependency-graph layer.
- `packages/xlsx/src/formula/values.ts` — runtime `Value`
  discriminated union (`number | string | bool | error | range`),
  Excel-faithful coercion (`toNumber`, `toString`, `toBoolean`),
  comparison (`compare`, `eq`, `lt`, `gt`, `lte`, `gte`, `ne`)
  with the type-class ordering quirk (numbers < strings <
  booleans), and arithmetic (`add`/`sub`/`mul`/`div`/`pow`/`neg`/
  `pct`/`concat`) with full error propagation.
- `packages/xlsx/src/formula/references.ts` — A1 ↔ internal
  `{row, col, abs}` (`AbsRef.NONE/ROW/COLUMN/ALL`) translation,
  sheet-prefix normalisation (handles `'Sheet Name'!`-style
  quoting and the `''` escape), absolute-ref-aware insert/delete
  adjustments, and A1 ↔ R1C1 conversion. Range parsing covers
  whole-row (`3:5`), whole-column (`A:A`), and single-sheet
  rectangles. `cellKey` produces the canonical
  `Sheet!R{row}C{col}` key the dependency graph hashes on.

`packages/xlsx/src/formula/__tests__/foundation.test.ts` — 38 unit
tests pinning every coercion / comparison / arithmetic / reference
behaviour described above.

**Phase 7b — lexer + parser + AST**

- `packages/xlsx/src/formula/ast.ts` — `AstNode` discriminated
  union (`lit | ref | range | name | binary | unary | pct | call |
array`) with stable source spans, the `Formula` carrier
  (`text`, `ast`, `anchor`, `dependencies`, `volatile`), and the
  curated `VOLATILE_FUNCTIONS` set. Helpers
  `collectDependencies(ast)` (deduped by `cellKey` /
  `${sheet}!${a}:${b}`) and `containsVolatile(ast)` walk the tree
  in a single pass.
- `packages/xlsx/src/formula/lexer.ts` — single forward-scan
  tokenizer. Strips a leading `=`, recognises numeric literals
  (with `e±N` exponent), strings (with `""` doubled-quote
  escape), errors, booleans, two-char comparison operators,
  percent, parens, braces, comma, semicolon, colon, and
  references. Reference recognition is a regex catalogue — quoted
  sheet name, bare sheet name, A1 cell, A1:A1 cell range,
  whole-column (`$?A:$?A`), whole-row (`$?N:$?N`) — with
  disambiguation against function-call identifiers and
  digit-prefixed row ranges (`3:5` is a `RANGE_REF`, not two
  `NUMBER`s separated by `COLON`). `$`-prefixed refs are handled
  before the digit/identifier dispatch.
- `packages/xlsx/src/formula/parser.ts` — recursive-descent +
  precedence-climbing parser. Operator precedence (higher binds
  tighter): `^` (8, right) > `*` `/` (7, left) > `+` `-` (6) >
  `&` (5) > comparisons (4). Unary `-` / `+` parse the operand at
  precedence 9, encoding the Excel quirk that `-2^2 == 4`.
  Function calls upper-case the name; defined names resolve to
  refs/ranges at parse time when supplied via `ParseOptions`.
  Array literals parse `{1,2;3,4}` row-by-row. Intersection
  (`;` outside arrays) raises a typed
  `intersection-operator-not-supported` error so we can wire the
  upgrade path later.

`packages/xlsx/src/formula/__tests__/lexer.test.ts` (15 tests) +
`packages/xlsx/src/formula/__tests__/parser.test.ts` (19 tests)
pin tokenisation edge cases, full operator-precedence behaviour
(including the `-2^2` quirk), dependency-collection dedup,
volatility flagging (`RAND`, `NOW`, `INDIRECT`, `OFFSET`), array
literals, defined-name resolution, and every parse-error path
including the rejected intersection operator.

**Totals after 7b**

- `@officeai/xlsx` unit tests: **135 passing** (Phase 5 → 48; Phase
  6 → 63; Phase 7a → 101; Phase 7b → 135).
- Full `make verify` (format-check / lint / architecture /
  typecheck / test / build) green.

**Decisions**

- **Precedence-climbing over shunting-yard.** Easier to reason
  about, fewer auxiliary data structures, and a near-1:1 mapping
  to the spec's operator-precedence table.
- **Higher-number-binds-tighter convention.** Lets `parseExpression`
  start at `0` and climb upward without inverted predicates;
  matches the canonical formulation in most parser literature.
- **Lexer emits sheet-qualified refs as a single token.** The
  parser reuses `parseA1` / `parseA1Range` from `references.ts`,
  so sheet handling lives in exactly one place.
- **Intersection operator (`;`) is a typed parse error, not a
  silent fallback.** Surfacing it at parse time means a future
  upgrade path (real intersection or named-range workaround) is
  one switch-case away.
- **Defined-name resolution is opt-in via `ParseOptions`.** Keeps
  the parser pure when no name table is available (e.g. agent
  command pre-validation) and avoids accidental cross-workbook
  coupling.

## 2026-04-18 — Phase 7c: function registry + tree-walking evaluator

**Shipped**

- `packages/xlsx/src/formula/function-registry.ts` — pluggable
  registry with `EvalContext` (cell / range accessors, defined-name
  resolver, `now`, `random`, `anchor`, registry self-reference),
  `Arity` descriptor (`min`, `max`, `accepts(n)`), `EagerFn` /
  `LazyFn` distinction, and a `MutableFunctionRegistry` builder
  (`createRegistry`). Volatile functions self-register into a
  `volatileNames()` set the dependency graph (Phase 7d) consumes
  to know which cells force-dirty every recalc.
- `packages/xlsx/src/formula/evaluator.ts` — single-pass post-order
  tree walker. Exhaustive `switch` over `AstNode["kind"]`
  (per `typescript-exhaustive-switch.mdc`). Binary operators
  delegate to a `BINARY_OPS` map of helpers from `values.ts` so
  error short-circuit and Excel coercion live in exactly one
  place. Unknown functions yield `#NAME?`; arity mismatch yields
  `#N/A`. The `lazyArgs` flow constructs a `LazyEvalAccess`
  closure that gives `IF` / `IFS` / `SWITCH` / `IFERROR` / `IFNA`
  full control over argument evaluation order (so they short-
  circuit correctly).

**Tests landed**

- `packages/xlsx/src/formula/__tests__/evaluator.test.ts` — 27
  tests grouped into seven scenarios:
  - literals + operators (arithmetic, string concat, comparisons,
    `=-2^2 == 4` quirk, percent, error propagation through `+`)
  - references (cell read, range read, unresolved defined name,
    resolved defined name)
  - eager function dispatch (registry hit, case-insensitive
    lookup, `#NAME?` on miss, `#N/A` on arity mismatch, range
    flattening inside SUM-like)
  - lazy-arg dispatch (chosen branch only — verified via
    side-effect counters, omitted third arg defaults to FALSE,
    error condition propagation)
  - array literals (`={1,2;3,4}` → 2×2 RangeValue)
  - context plumbing (`now` / `random` / `anchor` reach the
    function impl; `volatileNames` set tracks `volatile: true`
    registrations)
  - parser ↔ evaluator end-to-end (custom `DOUBLE` function,
    `=#REF!+1` error literal propagation, `err()` helper round-trip)

**Totals after 7c**

- `@officeai/xlsx` unit tests: **162 passing** (Phase 7a → 38; 7b
  → +34; 7c → +27).
- Full `make verify` (format-check / lint / architecture /
  typecheck / test / build) green.

**Decisions**

- **Registry-as-data, not as plugin DI.** Functions register via a
  thin imperative builder (`reg.register({...})`); there's no
  decorator metadata, no introspection-based wiring, no DI
  container. Adding a function in Phase 7e is one call.
- **Lazy args expose a closure, not a parser.** The lazy
  `LazyEvalAccess` shape is `{ evaluate(node), ctx }` — every lazy
  function calls `lazy.evaluate(args[i])` to drive its own order.
  This is what unlocks correct `IF` short-circuit, `IFERROR`'s
  selective catch, and `IFS`'s pair-walking without leaking
  evaluator internals to function impls.
- **Error short-circuit lives in `values.ts`, not the evaluator.**
  Every binary op runs `BINARY_OPS[op](a, b)` which delegates to
  the §7.4 helpers. The evaluator is type-driven (one switch) and
  has no try/catch.
- **`#NAME?` for unknown functions, `#N/A` for arity mismatch.**
  Matches Excel parity (per spec §13). The unknown-function path
  doubles as the P0 rendering for the deferred date / finance
  categories — the registry returns `undefined`, the evaluator
  returns `#NAME?`, the cell text round-trips verbatim.
- **`volatileNames()` is the registry's only structured
  introspection point.** The dependency graph (7d) needs the
  volatile set at parse time to flag formulas; everything else
  about a function (arg shapes, return types) is opaque.

## 2026-04-18 — Phase 7d: dependency graph + recalc orchestrator

**Shipped**

- `packages/xlsx/src/formula/dependency-graph.ts` — `DepGraph` with
  forward/reverse edges, a per-sheet flat range-overlay index,
  volatile tracking, and a Kahn topological drain plus Tarjan SCC
  for cycle detection. `addCell` rebuilds edges and dirties the
  cell; `removeCell` drops edges and dirties downstream;
  `markDirty(ref)` walks transitive closure across both cell-edges
  and range-overlap entries; `drainTopological()` returns
  `{ order, cycles }` where cycles are SCCs of size > 1 (or
  self-loops).
- `packages/xlsx/src/formula/recalc.ts` — public `FormulaEngine`
  (`parse`, `addCell`, `removeCell`, `onCellChanged`, `recalc`,
  `recalcAll`, `getCachedValue`) bound to a passive `EngineHost`
  that supplies cell/range data. Engine owns formula caches; host
  owns raw user-typed values and stitches them with the engine's
  cache on range materialisation. Cycle SCCs short-circuit to
  `Errors.refWithCycle(...)` before evaluator dispatch.

**Tests landed**

- `packages/xlsx/src/formula/__tests__/recalc.test.ts` — 11 tests
  covering: simple A1→B1→C1 chain in topological order; partial
  recalc (only affected dependents re-fire); `removeCell` edge
  drop + downstream dirty; range-dependency dirty bubble (`SUM(A1:A3)`
  re-fires when `A2` changes, but not when an out-of-range `D1`
  changes); cycle detection for 2-cell, 3-cell, and self-loop
  cases (each surfaces `#REF!` with `meta.cycle`); volatile
  re-fire (`=RAND()` recomputes every recalc with no edits);
  `recalcAll` for blanket re-evaluation; and a perf smoke
  asserting a 1k-cell linear chain recalcs in < 50ms (the §17
  budget allows 100ms for 10k formulas).

**Totals after 7d**

- `@officeai/xlsx` unit tests: **173 passing** (Phase 7a → 38;
  7b → +34; 7c → +27; 7d → +11).
- Full `make verify` (format-check / lint / architecture /
  typecheck / test / build) green.

**Decisions**

- **Engine is passive; host owns raw cell data.** The engine never
  touches the workbook directly. `EngineHost.readCell` /
  `readRange` are the only data-side hooks. This keeps the engine
  reusable across the headless `XlsxAgent` and the future browser
  editor without a DOM dependency.
- **Range index v1 is a flat `Map<sheet, RangeDep[]>`.** Per
  spec §14.1 the linear scan is fine for the §17 fixture; we
  switch to an interval index only if profiling demands it. The
  Kahn-walk dirty bubble walks both cell-edges and range entries
  on every step, so range-only formulas (`=SUM(A1:A1000)`) recalc
  correctly when an inner cell changes.
- **Cache split: formulas in the engine, raw values in the host.**
  When a downstream formula reads `A1` mid-recalc, the engine
  consults its formula cache first (so the precedent's freshly
  computed value wins), and falls through to the host for plain
  value cells. This is the contract that keeps topological order
  meaningful.
- **Cycles surface as `#REF!` with `meta.cycle`, never as
  exceptions.** The recalc loop assigns the cycle error to every
  cell in the SCC before the evaluator runs on the topo-sorted
  remainder. No try/catch; no infinite recursion.
- **`recalcAll` is opt-in for the host.** The normal `recalc()`
  drains only the dirty + volatile sets (per §17). `recalcAll`
  marks every known cell dirty and re-runs — used when the host
  re-imports a workbook or when a user explicitly forces
  recalculation.

## 2026-04-18 — Phase 7e: P0 function library (math/logic/info/lookup/text)

**Shipped (5 parallel subagents)**

- `packages/xlsx/src/formula/functions/math.ts` — **33** P0
  math/stats functions: `SUM, AVERAGE, COUNT, COUNTA,
COUNTBLANK, MIN, MAX, SUMIF, SUMIFS, COUNTIF, COUNTIFS,
AVERAGEIF, AVERAGEIFS, ROUND, ROUNDUP, ROUNDDOWN, INT, ABS,
MOD, POWER, SQRT, CEILING, FLOOR, RAND, RANDBETWEEN, LARGE,
SMALL, RANK, MEDIAN, STDEV, VAR, PRODUCT, SUMPRODUCT`. Shared
  `parseCriteria` supports `=`, `<>`, `>`, `<`, `>=`, `<=` with
  `*` / `?` wildcards (escaped via `~`). `RAND` / `RANDBETWEEN`
  are volatile and source bits via `ctx.random()`.
- `packages/xlsx/src/formula/functions/logic.ts` — **11** P0 logic
  functions: `IF, IFS, AND, OR, NOT, XOR, IFERROR, IFNA, SWITCH,
TRUE, FALSE`. `IF / IFS / IFERROR / IFNA / SWITCH` declare
  `lazyArgs: true` and short-circuit unevaluated branches.
  `IFERROR` catches every `CellError`; `IFNA` only `#N/A`.
- `packages/xlsx/src/formula/functions/info.ts` — **10** P0 info
  functions: `ISBLANK, ISNUMBER, ISTEXT, ISERROR, ISNA, ISODD,
ISEVEN, TYPE, N, NA`. `IS*` predicates do **not** propagate
  errors — they inspect type by `kind`. `ISODD`/`ISEVEN` _do_
  propagate (they consume the number). `ISBLANK` uses object
  identity against the exported `Blank` singleton (a documented
  P0 limitation pending a richer Blank kind).
- `packages/xlsx/src/formula/functions/lookup.ts` — **12** P0
  lookup functions: `VLOOKUP, HLOOKUP, INDEX, MATCH, XLOOKUP,
CHOOSE, OFFSET, INDIRECT, ROW, ROWS, COLUMN, COLUMNS`.
  `OFFSET` and `INDIRECT` are volatile. `OFFSET / ROW / ROWS /
COLUMN / COLUMNS` use `lazyArgs: true` to peek at AST `RefNode` /
  `RangeRefNode` before eager eval collapses them. `INDEX`
  supports both scalar pick and `row=0`/`col=0` whole-row/-column
  slicing.
- `packages/xlsx/src/formula/functions/text.ts` — **23** P0 text
  functions: `CONCATENATE, CONCAT, TEXTJOIN, LEFT, RIGHT, MID,
LEN, TRIM, UPPER, LOWER, PROPER, FIND, SEARCH, SUBSTITUTE,
REPLACE, REPT, TEXT, VALUE, NUMBERVALUE, CHAR, CODE, EXACT,
T`. `FIND` is case-sensitive, no wildcards; `SEARCH` is
  case-insensitive with `?`/`*` wildcards (escape via `~`).
  `TEXT` is a P0-minimal format engine handling `0[.0…]`,
  `#,##0[.0…]`, `0%`, and `$…` patterns; everything else falls
  back to `toString` (full number-format engine deferred to
  §16.7).
- `packages/xlsx/src/formula/functions/index.ts` — single
  `registerAllFunctions(reg)` aggregator that bulk-registers all
  five categories.

**Tests landed**

- `math.test.ts` — 100 tests (≥3 per function plus criteria-parser
  cases and SUMPRODUCT shape validation).
- `logic.test.ts` — 36 tests including side-effect counters that
  prove un-chosen branches of `IF / IFS / IFERROR / IFNA /
SWITCH` are not evaluated.
- `info.test.ts` — 48 tests including the inverted-error-handling
  contract (`ISERROR(1/0)` → TRUE, `ISNUMBER(1/0)` → FALSE).
- `lookup.test.ts` — 41 tests with a fixed `getRange` table host
  exercising exact + approximate lookup, INDEX whole-row/-col
  slicing, MATCH all three modes, XLOOKUP `if_not_found`.
- `text.test.ts` — 77 tests covering coercion edges, wildcard
  searches, `TEXT` format hints, and the 32 767-char REPT/TEXTJOIN
  ceiling.

**Totals after 7e**

- **89 functions** registered across 5 categories.
- `@officeai/xlsx` unit tests: **475 passing** across 15 files
  (Phase 7a 38 → 7b +34 → 7c +27 → 7d +11 → 7e +302; +
  serializer/parser/agent/handlers from earlier phases).
- Full `make verify` (format-check / lint / architecture /
  typecheck / test / build) green.

**Decisions**

- **One file per category, single `registerXxx` entrypoint.**
  Keeps blast radius small, parallelisable across subagents, and
  matches the §16 spec layout. `functions/index.ts` is the only
  cross-category file; the workbook layer wires the entire
  library in one call.
- **Errors are values, never exceptions.** Every function
  short-circuits on `kind === "e"` arguments and returns a
  `CellError` for its own failures. `walkSumNumeric` / similar
  helpers thread the first error back to the caller verbatim.
- **Lazy args reserved for short-circuit semantics or AST
  inspection.** Logic gates (`IF / IFS / IFERROR / IFNA /
SWITCH`) need lazy to skip un-chosen branches; lookup helpers
  (`OFFSET / ROW / ROWS / COLUMN / COLUMNS`) need lazy to
  inspect the un-collapsed `RefNode` for ref arithmetic. Every
  other function is eager.
- **Volatility declared at registration.** `RAND`,
  `RANDBETWEEN`, `OFFSET`, `INDIRECT` set `volatile: true`.
  The dep graph (7d) drains the volatile set on every `recalc()`,
  matching §13.4.
- **Documented P0 deferrals (called out in code comments):**
  - `COUNTBLANK` only detects formula-blanks (`""`); a true
    empty-cell distinction needs an engine-side `Blank` kind.
  - `OFFSET` requires its `base` to be a literal `RefNode` /
    `RangeRefNode`; computed bases return `#VALUE!`.
  - `INDIRECT(text, FALSE)` (R1C1 mode) returns `#REF!`; A1 mode
    is fully supported.
  - `XLOOKUP` `match_mode = 2` (wildcard) and binary `search_mode`
    not implemented.
  - `TEXT` covers the listed format hints; full number-format
    engine deferred.
  - `VLOOKUP / HLOOKUP / MATCH` exact-match wildcards not yet
    wired (deferred until text helpers expose a shared wildcard
    matcher).

## 2026-04-18 — Phase 7f: `xlsx:set-cell-formula` command

**Shipped**

- `packages/xlsx/src/formula/workbook-host.ts` — `WorkbookHost`
  adapter that lets the formula engine read cells from an
  `XlsxWorkbook` snapshot. Stitches the engine's formula cache
  (precedents already computed) with the snapshot's typed raw
  values; ranges materialise via `readRange`. Exports
  `bindEngineToWorkbook(workbook)` returning `{ engine, host }`,
  plus `toEngineValue` / `fromEngineValue` for the model ↔ engine
  value-union conversion. The function registry is a module-level
  singleton — re-registering all 89 P0 functions on every command
  would dominate the recalc cost.
- `packages/xlsx/src/commands/set-cell-formula.ts` — handler
  implementing the §2 spec contract: validate sheet/ref/merge
  anchor, strip leading `=`, parse via the engine (loud reject on
  parse error per `EC-F3`), seed the engine with every existing
  formula, add the new one, recalc, and write the cached value
  for both the target cell and any dependents whose values
  changed. Empty-formula bodies collapse to a clear-cell. Cycles
  surface a `circular` diff entry with `meta.cycle` carrying the
  SCC.
- `packages/core/src/types/document.ts` — extended every
  `DiffChange` variant with an optional `meta?: DiffMeta` payload
  (`DiffMeta = Readonly<Record<string, unknown>>`). The cycle
  metadata, before/after value snapshots, and any future
  format-specific extras live here. Older readers that don't
  recognise `meta` MUST ignore it.
- Wired `setCellFormulaHandler` into
  `packages/xlsx/src/commands/registry.ts` and re-exported from
  `commands/index.ts`. Bus surface count: **6/13 P0 commands**.

**Tests landed**

- `packages/xlsx/src/commands/set-cell-formula.test.ts` —
  **11 tests** covering: literal arithmetic + cached value;
  with / without leading `=`; whitespace verbatim preservation;
  `SUM(Y1:Y3)` over real cells; downstream propagation when a
  later formula depends on an earlier one; malformed-formula
  reject (`formula-parse-error`); unknown-sheet reject;
  invalid-ref reject; empty-formula clear; `#DIV/0!` cached as
  `CellErrorValue`; self-loop cycle producing `#REF!` with a
  `circular` diff entry carrying `meta.cycle`.

**Totals after 7f**

- `@officeai/xlsx` unit tests: **486 passing** (+11).
- 6/13 P0 commands now wired through the bus.
- Full `make verify` green.

**Decisions**

- **Engine is rebuilt per command, registry is shared.** The
  command bus is a pure-snapshot model — handlers are
  `(snapshot, payload) → { next, diff }`. To run a real recalc we
  need a `FormulaEngine` populated with the workbook's formulas.
  Building the engine + parsing every existing formula on each
  command is O(N) — fine at the §17 fixture scale (10k formulas
  ~ tens of ms). A future optimisation can cache the engine keyed
  on `snapshot.revision`. Keeping the function registry
  module-singleton dodges the ~90 `register()` calls on every
  command.
- **Parse errors reject loudly; import errors stay quiet.** Per
  `EC-F3`, formulas authored by a user/agent that fail to parse
  return a `formula-parse-error` rejection so the caller (LLM)
  gets feedback. Formulas already in the workbook that fail to
  parse during `seedFormulas` are silently skipped (they remain
  literal cells with whatever cached value the parser had); they
  surface as `#NAME?` only at import time.
- **Downstream cached values are written back to the model.**
  When the recalc touches a dependent formula cell, its
  `Cell.value` is replaced with the new cached value (preserving
  `Cell.formula`). The diff includes one `cell-updated` entry
  per cell that actually changed (no-op cells are skipped).
- **Cycles are surfaced, not eliminated.** Per `EC-F1` the cycle
  cells are written as `#REF!` (with `meta.cycle = [...refs]` on
  the diff). Excel silently flips back to a default; we expose
  the structured cycle list so review UIs can highlight it.
- **`DiffMeta` lives in `@officeai/core`, not the xlsx package.**
  Other formats (`docx` mark-text-blockquote, future `pptx`
  slide-transition metadata, etc.) will want the same hook. The
  field is optional and ignored by older readers, so existing
  `docx` consumers are unaffected.

## 2026-04-18 — Phase 7g: `xlsx:set-cell-format` command + typed style table

**Shipped**

- `packages/xlsx/src/model/style-table.ts` — typed in-memory model
  for `xl/styles.xml` (`StyleNumberFormat`, `StyleFont`, `StyleFill`,
  `StyleBorder`, `StyleCellXf`, plus the top-level `StyleTable` and
  a `defaultStyleTable()` factory). Anything we don't model (cell
  styles, `dxfs`, `tableStyles`, `extLst`, …) round-trips verbatim
  via `opaqueExtras: ReadonlyArray<{section, xml}>` so byte-clean
  re-emit holds for every untouched workbook.
- `packages/xlsx/src/parser/styles.ts` — schema-aware parser that
  walks `<numFmts>`, `<fonts>`, `<fills>`, `<borders>`, `<cellXfs>`
  in document order, captures font run properties (bold, italic,
  size, color, name, family, scheme, strike, underline), pattern
  fills (`patternType`, `fgColor`, `bgColor`), every border edge
  with style + color + diagonal flags, and the full `xf` cross-
  product (numFmtId / fontId / fillId / borderId / xfId + the four
  `applyXxx` flags + alignment + protection). Unknown sections are
  serialised back into `opaqueExtras` as raw XML strings.
- `packages/xlsx/src/serializer/styles.ts` — symmetric re-emitter
  that reconstructs `<styleSheet>` from the typed table. Round-trip
  is **semantic**, not byte-identical: attribute order can drift,
  but re-parsing the output yields a structurally equivalent table.
  Opaque sections are spliced back at the end of the sheet.
- `packages/xlsx/src/model/style-mutate.ts` — content-hash
  deduplication layer. `flattenCellXf(table, id)` resolves an `xf`
  index to an `EffectiveStyle` (the merged number-format /
  font / fill / border / alignment / protection bundle).
  `internStyle(table, eff)` returns `{table, xfId}`; reusing an
  identical xf hits the cache and never grows `cellXfs`. Component
  hashes (`hashFont`, `hashFill`, `hashBorder`, `hashXf`) are
  canonical JSON sorts, so `bold` after `italic` interns to the
  same id as `italic` after `bold`. This is what lets bulk
  formatting (e.g. bolding 10k cells with `font: { bold: true }`)
  add **at most one new font + one new xf** to the table — the
  "small constant" property in `agent-commands.md §4`.
- `Cell.styleId?: number` added to the typed cell. `undefined` =
  the implicit default xf (id 0). Plumbed through:
  - `parser/parse.ts` — for each worksheet we now scan the raw
    `xl/worksheets/sheetN.xml` ourselves to recover each cell's
    `s="N"` attribute. SheetJS overwrites `cell.s` with a resolved
    fill object when loaded with `cellStyles: true`, dropping the
    original index, so we cannot get it back from the dense store.
    A 30-line regex over `<c r="REF" s="N">` rebuilds the map
    before we materialise typed cells.
  - `serializer/sheet-sync.ts` — `typedCellToSheetJS` carries
    `cell.styleId` over to the SheetJS dense cell as `s: N`, with
    `prev.s` as the fallback so untouched styled cells keep their
    style on round-trip.
  - `serializer/serialize.ts` — after SheetJS emits the worksheet
    XML, we post-process every `<c r="REF" ...>` to inject `s="N"`
    from the typed cells (and strip any spurious `s` SheetJS
    invented from `cell.z`). This is the only practical way to
    reconcile our owned styles part with SheetJS's writer, which
    re-derives `s` from its own internal `cellXfs` table and
    ignores whatever we put on the cell.
  - `dirty.styles` is now a recognised dirty flag; `serializeXlsx`
    re-emits `xl/styles.xml` from the typed table when set.
- `packages/xlsx/src/commands/set-cell-format.ts` — the seventh
  P0 command. Validates the patch (RRGGBB hex colors, color names
  must be hex not "red", no diagonal-up/down without a side, etc.),
  resolves each target cell's current `EffectiveStyle`, applies
  the patch field-by-field (clear-to-default with `null`, deep
  merge otherwise), interns the result, and writes the new
  `styleId` onto the cell. Range form (`A1:C3`) iterates every
  cell in the rectangle; missing cells are materialised as
  blank-but-styled cells. Emits one `cell-updated` diff per
  cell that changed `styleId`, plus a single `style-table-grew`
  diff entry when new fonts / fills / borders / xfs are interned.
- `packages/xlsx/src/commands/payloads.ts` — `CellFormatPatch` is
  the friendly agent-facing shape (`font.bold`, `fill.color`,
  `border.top.style`, `numberFormat: "0.00"` or `numFmtId: 14`),
  translated to the OOXML field layout inside the handler. Border
  sides accept the 14 OOXML line styles; alignment accepts
  `horizontal: "center" | "left" | "right" | "justify" | "fill"`
  and the friendly `vertical: "top" | "middle" | "bottom"`
  (mapped to OOXML's `center` token internally).
- Wired into `commands/registry.ts` (now **7/13 P0 commands
  shipped**) and re-exported from `commands/index.ts` and the
  package root `src/index.ts`.

**Tests landed**

- `packages/xlsx/src/parser/__tests__/styles.test.ts` — **7 tests**
  covering: parse → emit → re-parse structural equivalence on the
  fixture corpus styles parts; default-table factory shape; opaque
  section preservation; missing-`numFmts` graceful fallback to the
  17 built-in formats; xf with `applyFont:0` round-trips the flag
  off; alignment + protection survival; full table → empty workbook
  → reload round-trip.
- `packages/xlsx/src/commands/set-cell-format.test.ts` — **9 tests**
  covering: bolding a single cell mints exactly one new font + one
  new xf; bolding 100 cells mints **the same one** font + xf
  (dedup); range form `A1:C3` styles every cell including blank
  ones; `font.color = "FFEEAA"` survives; clearing with
  `numberFormat: null` reverts the cell to `numFmtId: 0`; agent-side
  validation rejects malformed colors (`"red"`, `"#FF00FF"`,
  `"FF00FFAA"`), unknown sheets, invalid ranges, and unknown
  `numberFormat` strings; full parse → format → serialize → reparse
  round-trip preserves the `styleId` and the resolved bold flag.

**Totals after 7g**

- `@officeai/xlsx` unit tests: **502 passing** (+27: 9 set-cell-
  format + 7 styles parser + others picked up across the model
  layer; xlsx test count went 475 → 486 → 502).
- 7/13 P0 commands now wired through the bus. `xlsx:set-cell-
format` joins `set-cell-value`, `clear-range`, `delete-range`,
  `merge-cells`, `unmerge-cells`, and `set-cell-formula`.
- All xlsx + xlsx round-trip tests green; pre-existing
  `validate-ooxml-schemas` failures (real-world docx fixture
  count) are unrelated.

**Decisions**

- **Own the styles part end-to-end; ignore SheetJS's writer.**
  SheetJS's `write_ws_xml_cell` calls `get_cell_style(cellXfs,
cell, opts)` which derives the `s` attribute from `cell.z` (the
  number-format string) against SheetJS's own internal cellXfs
  table — it ignores any `s` we set on the SheetJS cell. We
  considered round-tripping styles through SheetJS but its style
  index space is incompatible with ours (we re-emit
  `xl/styles.xml` from the typed table; SheetJS would emit a
  different one). The post-process injection in
  `serialize.ts` is ~25 lines, fully tested, and isolates the
  workaround to one function.
- **Recover `s` from raw XML, not from SheetJS.** With
  `cellStyles: true` SheetJS replaces `cell.s` with a resolved
  fill object (line 15835 in `xlsx.mjs`), permanently losing the
  numeric index. Reading the worksheet XML once during parse and
  scanning `<c r="..." s="...">` is O(file size) and runs alongside
  the existing parse without measurable overhead on the §17
  fixture corpus.
- **Round-trip is semantic, not byte-identical, for `xl/styles.xml`.**
  Untouched workbooks still round-trip bit-identical (we only
  re-emit the styles part when `dirty.styles` is set). Once
  `set-cell-format` fires, the styles part is regenerated from the
  typed table; attribute order, whitespace, and unknown-section
  ordering can drift. Re-parsing the output yields a structurally
  equivalent table — that's the contract.
- **Friendly agent shape, OOXML-shaped storage.** The
  `CellFormatPatch` interface uses agent-friendly names
  (`font.bold`, `fill.color`, `border.top.style`) so an LLM can
  author it without reading the OOXML schema. The translation to
  `numFmtId` / `fontId` / `fillId` / `borderId` / `applyXxx`
  happens in the handler, exactly once, against the typed
  `StyleTable`. The typed model stays OOXML-shaped so the
  serializer is a straight projection.
- **Content-hash dedup over reference counting.** Excel's styles
  table is append-only in practice; we never garbage-collect
  unused entries. Hashing fonts/fills/borders/xfs by canonical
  JSON before insertion gives us O(1) "is this style already
  here?" without tracking refcounts. This is what makes bulk
  formatting cheap: bold-100-cells touches the table twice
  (one new font, one new xf), regardless of whether the source
  workbook already had bold or not.
- **Default xf (id 0) is implicit.** Cells with `styleId === 0` or
  `undefined` get no `s` attribute in the worksheet XML; this
  matches Excel's own emission and keeps untouched cells
  byte-clean. The serializer's injector strips `s="0"` defensively.

## 2026-04-18 — Phase 7h: `xlsx:add-sheet` command + multi-part rewrite

**Shipped**

- `packages/xlsx/src/commands/add-sheet.ts` — eighth P0 command.
  Validates the proposed name (Excel rules: 1–31 chars, no
  `[ ] * ? : / \`, not "History", case-insensitive uniqueness)
  and the optional `at` insert position (`[0, sheets.length]`),
  mints a fresh `sheetId` (smallest unused positive integer) and
  a fresh part path (smallest free `xl/worksheets/sheetN.xml`
  that does not collide with any existing part — including
  opaque ones), splices the typed `Sheet` into `workbook.sheets`
  and re-derives the `index` of every shifted neighbour, mints
  a parallel SheetJS `WorkSheet` (`{ "!ref": "A1", "!data": [] }`)
  so the existing `rewriteDirtySheets` pipeline emits the bytes
  without a special case, and sets four dirty flags: `workbook`,
  `rels`, `contentTypes`, and the new `sheets[partPath]`. Diff
  is a single `node-inserted` change carrying
  `{ name, at, sheetId, partPath }` in `meta` so undo/redo or
  agent UIs can reconstruct the insertion.
- `packages/xlsx/src/serializer/serialize.ts` — promoted
  `dirty.contentTypes` and `dirty.rels` from "unsupported" to
  first-class:
  - `rewriteContentTypes` walks `workbook.sheets` and adds an
    `<Override PartName="/xl/worksheets/sheetN.xml"
ContentType=".../worksheet+xml"/>` for any sheet missing
    one. Re-emits the part only when an entry was actually
    added so untouched workbooks stay byte-identical.
  - `rewriteWorkbookRels` does the symmetric pass on
    `xl/_rels/workbook.xml.rels` via
    `ooxml.RelationshipGraph.loadFor`: for every worksheet
    without a rel pointing at its `partPath`, it mints a fresh
    `rId` (the next free `rId{N}`) and adds a `Relationship`
    of the worksheet rel-type. Existing rels (sharedStrings,
    styles, theme, calcChain, …) are preserved verbatim.
  - `rewriteWorkbookSheets` (renamed from
    `rewriteWorkbookSheetNames`) is generalized: it rebuilds
    just the `<sheets>` element of `xl/workbook.xml` from
    `workbook.sheets` in tab order, looking up each sheet's
    `r:id` via the just-rewritten rels. Every other byte of
    `xl/workbook.xml` (namespaces, comments, `<bookViews>`,
    `<definedNames>`, `<calcPr>`, …) is preserved via a
    string-level splice. Renames, insertions, and (eventually)
    reorders all flow through this single rewrite.
  - Order: `sheets → contentTypes → rels → workbook → styles`.
    The workbook rewrite runs after rels so `r:id` lookups see
    the freshly minted ids for new sheets.
- `packages/xlsx/src/serializer/sheet-sync.ts` — when the typed
  cell store is empty, `syncSheetToSheetJS` now seeds the dense
  `!data` with one empty row. Pairing `!ref="A1"` with
  `!data=[]` triggers `Cannot read properties of undefined`
  inside SheetJS's `write_ws_xml_data`, which iterates
  `data[R]` for every row in the declared range. The
  one-empty-row seed satisfies the iteration without emitting
  any `<c>` cells.
- Wired into `commands/registry.ts` (now **8/13 P0 commands
  shipped**) and re-exported from `commands/index.ts` plus the
  package root `src/index.ts`.

**Tests landed**

- `packages/xlsx/src/commands/add-sheet.test.ts` — **14 tests**
  covering: append (default `at`); insert at index 0 with all
  existing sheets shifted right and `index` re-derived;
  insert at a middle index; `node-inserted` diff carries
  `name`/`at`/`sheetId`/`partPath`; reject case-insensitive
  duplicate name (`duplicate-name`); reject forbidden
  characters (`invalid-name`); reject reserved name "history"
  case-insensitive (`invalid-name`); reject empty name and
  > 31-char name; reject negative and out-of-range `at`
  > (`invalid-position`); full parse → add → serialize →
  > re-parse round-trip with all original sheets intact and the
  > new sheet present at the requested index; same round-trip on
  > a single-sheet fixture (append); follow-up
  > `xlsx:set-cell-value` on the freshly added sheet survives the
  > serializer round-trip.

**Totals after 7h**

- `@officeai/xlsx` unit tests: **516 passing** (+14).
- 8/13 P0 commands now wired through the bus.
- `pnpm --filter @officeai/xlsx test`, `lint`, `build`, and
  `pnpm format:check` all green.

**Decisions**

- **Handler stays typed-model-only; serializer owns container
  rewrites.** Following the `rename-sheet` pattern, the
  add-sheet handler does not touch the `OoxmlContainer`. It
  mutates the typed `XlsxWorkbook.sheets` array, the
  parallel SheetJS book, and sets dirty flags. The serializer
  is the single place that touches `[Content_Types].xml`,
  `xl/_rels/workbook.xml.rels`, and the workbook part. Keeps
  snapshots logically pure (the container is shared by
  reference but mutated only on serialize, on a clone) and
  matches the contract the round-trip tests already enforce.
- **Mint `sheetId` and part path independently.** OOXML uses
  `<sheet sheetId="N">` for the stable workbook-internal id
  (referenced by `<definedName>` and similar) and an unrelated
  `xl/worksheets/sheet{M}.xml` part path. We mint each
  separately as the smallest unused positive integer in its
  own namespace; this matches what Excel does when you delete
  Sheet2 and add a new sheet (you get `sheetId=4`, part path
  `sheet4.xml`, but Excel happily reuses `sheet2.xml` for the
  next added sheet if it's free).
- **`r:id` minting lives in the rels rewriter, not the
  handler.** The handler doesn't know which `rId{N}` is free
  in `xl/_rels/workbook.xml.rels` because it doesn't read the
  container. Threading the rels graph through every command
  for a single string would be invasive. Instead the
  serializer's rels pass mints `rId` for each sheet missing a
  relationship, and the workbook-sheets pass looks up the
  fresh `r:id` via target-path matching. Order-of-operations
  in `serializeXlsx` enforces this dependency.
- **Rebuild `<sheets>` instead of surgical insert.** The old
  `rewriteWorkbookSheetNames` was a regex-based name swap that
  preserved the original `<sheet>` attribute byte order. Adding
  insertion via more regex was brittle, so the rewrite now
  reconstructs the entire `<sheets>` element from
  `workbook.sheets`. Byte-identical preservation only matters
  when `dirty.workbook = false`; once we touch the workbook
  part the spec already requires us to re-emit it. The rest
  of `xl/workbook.xml` is untouched via a string-level splice
  around the `<sheets>...</sheets>` block, so `<bookViews>`,
  `<definedNames>`, `<calcPr>`, etc. round-trip exactly.
- **Empty SheetJS sheet needs a seed row.** SheetJS's
  `write_ws_xml_data` iterates `data[R]` for every row in
  `!ref`; an empty `!data: []` paired with `!ref: "A1"` throws
  `Cannot read properties of undefined`. Seeding a single empty
  row in `syncSheetToSheetJS` (when the typed cell store is
  empty) costs one allocation and emits zero `<c>` children —
  the serialized worksheet XML is the canonical empty
  `<sheetData/>` shape.
- **Inverse is `xlsx:delete-sheet`, deferred.** Per
  `analysis-agent-patterns.md` §8.4, sheet deletion is
  unrecoverable in-session and lands later (P1+). The
  add-sheet diff carries `meta.partPath` so a future
  delete-sheet inverse can target the exact part without
  re-resolving by name.

### Phase 7i — structural reshape commands (2026-04-18)

`xlsx:insert-row`, `xlsx:insert-column`, `xlsx:delete-row`,
`xlsx:delete-column` ship together. These are the four commands
spec'd in `spec/xlsx/agent-commands.md` §§5–8 and the first
commands that need to rewrite formulas across the entire
workbook, not just the targeted sheet.

**Code landed**

- `packages/xlsx/src/formula/serialize-ast.ts` — AST → canonical
  formula text. Walks the AST with the surrounding operator
  precedence threaded through so binary children get parens
  iff dropping them would re-associate the expression on a
  round-trip (`(1+2)*3` keeps parens, `1+2*3` does not).
  Function names are uppercased and whitespace dropped to match
  Excel's own canonicalisation on edit. Used by the formula
  rewriter below; reused by anything that needs to emit Excel
  formula text from a parsed AST.
- `packages/xlsx/src/formula/rewrite-refs.ts` — `rewriteFormulaRefs`.
  Parses a formula, walks every `ref`/`range` node through an
  `AdjustFn` from `references.ts`, replaces deleted references
  with `#REF!` error literals, and re-serialises via
  `serializeAst`. Returns `{ text, hasRefError, changed }` so
  callers can decide whether to emit a `referenced-cell-deleted`
  diff and whether to update the cell at all.
- `packages/xlsx/src/commands/structural-shift.ts` — shared
  pipeline for the four commands. `applyStructuralOp` validates
  `at`/`count` against Excel's row/column limits, prechecks that
  no merge straddles the deletion/insertion boundary, shifts
  cells (or drops them, for deletes), adjusts every merge
  region (shift / expand / shrink / drop), iterates _every_
  formula in _every_ sheet through `rewriteFormulaRefs`, rebinds
  the formula engine, runs a full workbook recalc, and folds the
  recalc deltas back into the next snapshot. Returns the next
  workbook plus the structured pieces individual handlers need
  to assemble their `DiffChange` arrays.
- `packages/xlsx/src/commands/insert-row.ts`,
  `insert-column.ts`, `delete-row.ts`, `delete-column.ts` —
  thin handlers that resolve the sheet, delegate to
  `applyStructuralOp`, and emit the per-command `DiffChange`s
  (`rows-inserted` / `columns-inserted` /
  `rows-deleted` / `columns-deleted` summary, one
  `formula-updated` per rewritten formula, one
  `referenced-cell-deleted` per `#REF!` casualty for the delete
  variants, plus `cachedValue` deltas for any non-rewritten
  formula whose cached value changed across the recalc).
- Wired into `commands/registry.ts` (now **12/13 P0 commands
  shipped**) and re-exported from `commands/index.ts`.

**Tests landed**

- `packages/xlsx/src/formula/__tests__/serialize-ast.test.ts` —
  20 tests: literals (numbers, strings with embedded quotes,
  booleans, errors), cell + range refs with and without
  `anchorSheet`, unary/percent/binary precedence,
  right-associative `^`, function calls, defined names,
  array literals.
- `packages/xlsx/src/formula/__tests__/rewrite-refs.test.ts` —
  11 tests: cell shift on insert, range shift, cross-sheet refs
  ignored unless the sheet matches, single-cell deletion →
  `#REF!`, range entirely inside the deletion → `#REF!`, range
  partially overlapping the deletion shrinks correctly, complex
  expression with mixed refs, idempotency for formulas with no
  relevant refs.
- `packages/xlsx/src/commands/insert-row.test.ts` (13 tests),
  `insert-column.test.ts` (9 tests), `delete-row.test.ts`
  (12 tests), `delete-column.test.ts` (12 tests). Cover the
  happy paths (cell + formula shift, summary diff shape),
  merge handling (expand on touching the boundary, shift past
  the band, reject on mid-band split), validation (unknown
  sheet, `at` < 1, `count` < 1, overflow past Excel's
  1,048,576-row / 16,384-column limits), and the
  `#REF!`-casualty path for the delete variants (single ref
  → `#REF!`, range fully inside deletion → `SUM(#REF!)`, the
  `referenced-cell-deleted` diff entry shows up).
- `tests/roundtrip/xlsx/commands-roundtrip.test.ts` — two new
  end-to-end cases: parse → insert-row + insert-column →
  serialize → reparse preserves the rewritten `=Z3*3` formula
  and shifted operand; the same loop with delete-row +
  delete-column persists `=#REF!+1` after the referenced cell
  is removed.

**Totals after Phase 7i**

- `@officeai/xlsx` unit tests: **596 passing** (+80 from Phase 7h).
- Integration tests: **51 passing** (+2).
- 12/13 P0 commands now wired through the bus; only
  `xlsx:set-named-range` remains.
- `pnpm --filter @officeai/xlsx test`, `lint`, `build`, and
  `pnpm format:check` all green.

**Decisions**

- **One `applyStructuralOp` for all four commands.** The cell
  shift, merge adjust, formula rewrite, and recalc steps are
  identical modulo axis (row vs column) and direction (insert
  vs delete). Inlining the pipeline in each handler would
  duplicate ~200 lines of subtle index arithmetic four times;
  centralising it lets the four handlers stay <60 lines each
  and keeps the per-axis index math in one auditable spot.
- **Formula rewrite walks the whole workbook, not just the
  targeted sheet.** Cross-sheet refs (`Sheet2!A5`) on _any_
  sheet may need adjustment when `Sheet2` shrinks or grows.
  The cheaper "scan only the targeted sheet" optimisation
  would silently leave stale formulas elsewhere. The recalc
  pass after the rewrite catches dependent value changes for
  formulas that didn't textually change.
- **`#REF!` is emitted as a literal `ErrorLiteral` AST node,
  re-serialised inline.** Excel writes `=#REF!+1`, not
  `=ERR("#REF!")+1`; the rewriter inserts an
  `ErrorLiteral("#REF!")` exactly where the doomed `ref` /
  `range` node lived so the surrounding expression text stays
  syntactically valid and matches Excel's canonical form.
- **Precedence-aware serializer.** The naïve "always
  parenthesise binary children" emitter produced churn like
  `=A1+B1` ⇒ `=(A1)+(B1)` on every shift, which polluted the
  diff and ruined visual readability. Threading `parentPrec`
  through the recursion costs a handful of integer comparisons
  and produces text identical to what a human (or Excel) would
  write.
- **Defined names, comments, hyperlinks deferred.** The typed
  model doesn't yet expose these as walkable structures (they
  live inside opaque OPC parts). Phase 7i sets up the
  rewrite hook in `applyStructuralOp` so a follow-up phase can
  add the missing walks without changing the command surface.
  The spec's §§5–8 acceptance criteria for those parts are
  tracked as carry-over items.
- **`merge-boundary-crossed` is the single rejection code for
  any operation that would split a merge.** Earlier drafts had
  separate codes for "expand crosses boundary" vs "delete
  shaves region", but in the agent UI both reduce to "the
  merge in the way must be removed first" — a single code with
  meta.firstOffender (sheet + range) gives the agent enough
  context to recover.

### Phase 7i — structural reshape commands (2026-04-18)

**Shipped**

- `packages/xlsx/src/formula/serialize-ast.ts` — re-emits a
  parsed formula AST as canonical Excel formula text (no
  leading `=`). Handles every AST node kind exhaustively:
  literals (numbers / strings with `""` escaping / booleans /
  errors), single-cell and range refs (delegating to
  `serializeCellRef` / `serializeRangeRef` for `$` and
  cross-sheet `Sheet!` rendering), defined names, binary /
  unary / percent operators, function calls, and `{...;...}`
  array literals. Defensively wraps binary / unary / percent
  child expressions in parens to preserve operator precedence
  without tracking the parent's level. `anchorSheet` controls
  sheet-prefix omission so refs on the formula's own sheet emit
  bare (`A1` not `Sheet1!A1`).
- `packages/xlsx/src/formula/rewrite-refs.ts` — `rewriteFormulaRefs(text, anchor, adjust)`
  re-parses a formula against `anchor`, walks the AST with
  structural sharing (returning the same node when nothing
  changed below it), feeds every `Reference` and
  `RangeReference` through `adjust`, and re-serializes via
  `serializeAst(_, anchor.sheet)`. When `adjust` returns a
  `CellError` (the target lay inside a deletion band) the
  ref node is replaced by a literal `#REF!` token — Excel's
  canonical "deleted target" rendering per EC-R2. Returns
  `{ text, changed, hasRefError }` so callers can tell the
  difference between a no-op rewrite, a benign shift, and a
  rewrite that produced casualties.
- `packages/xlsx/src/commands/structural-shift.ts` —
  consolidated engine that powers all four reshape commands.
  Pipeline:
  1. Validate `at ≥ 1`, `count ≥ 1`, and the resulting band
     fits inside Excel limits (`1,048,576` rows / `16,384`
     columns) before any work.
  2. Resolve the target sheet by name (`unknown-sheet`
     rejection on miss).
  3. **Merge precheck.** For inserts: reject if the insertion
     index falls strictly _inside_ a merged region
     (`lo < at0 < hi`). For deletes: reject if the deletion
     band straddles a merge boundary (the merge extends outside
     the band on either side). Both produce
     `merge-boundary-crossed` rejections with `meta.range`.
  4. **Cell shift.** Insertions allocate a fresh `Map` and
     copy each cell with the row/column adjusted by `count`
     when at-or-after the insertion index. Deletions skip
     cells inside the band entirely and shift cells after the
     band up/left by `count`.
  5. **Merge shift.** Same shape as cell shift, with the merge
     precheck having already eliminated the partial-overlap
     cases.
  6. **Workbook-wide formula rewrite.** Walks every cell in
     every sheet (not just the target sheet — a formula on
     `Sheet2` may reference the deleted band on `Sheet1`),
     calls `rewriteFormulaRefs` with the appropriate
     `adjustForInsertRow` / `…InsertColumn` / `…DeleteRow` /
     `…DeleteColumn` from `formula/references.ts` scoped to
     the target sheet, and collects the rewritten cells for
     dirty-marking and diff emission.
  7. **Recalculate.** Rebuilds the formula `Engine` against the
     post-shift / post-rewrite snapshot and runs a full
     workbook recalc to refresh cached values for everything
     that depended on the moved or deleted refs.
  8. **Diff.** Emits one summary change
     (`rows-inserted` / `columns-inserted` / `rows-deleted` /
     `columns-deleted`) with `{ at, count, sheet }` in `meta`,
     plus `node-removed` for each cell dropped inside a delete
     band, `formula-updated` for each rewritten formula text,
     `referenced-cell-deleted` for each `#REF!` casualty, and
     `cell-updated` for each cached value that flipped during
     recalc.
- `packages/xlsx/src/commands/insert-row.ts`,
  `insert-column.ts`, `delete-row.ts`, `delete-column.ts` —
  thin wrappers; each handler resolves to one
  `applyStructuralShift(snapshot, payload, axis, op)` call so
  the four commands share identical semantics for validation,
  precheck, shift, rewrite, recalc, and diff. Specs:
  `agent-commands.md` §§5–8.
- `packages/xlsx/src/commands/payloads.ts` — added
  `InsertRowPayload`, `InsertColumnPayload`, `DeleteRowPayload`,
  `DeleteColumnPayload` (`sheet`, `at` 1-based, `count ≥ 1`).
- Wired into `commands/registry.ts` (now **12/13 P0 commands
  shipped**) and re-exported from `commands/index.ts` plus
  the package root `src/index.ts`.

**Tests landed**

- `packages/xlsx/src/formula/__tests__/serialize-ast.test.ts`
  — **20 tests** covering every literal kind (number / string
  with `""` escape / boolean / error / range fallback), bare
  - absolute + cross-sheet refs, range serialization, defined
    names, binary precedence wrapping, unary / percent, nested
    function calls, array literals, `anchorSheet` prefix
    omission, and a parse → serialize → parse semantic
    round-trip on a non-trivial formula.
- `packages/xlsx/src/formula/__tests__/rewrite-refs.test.ts`
  — **11 tests**: cell-ref shift on row insert, cell-ref shift
  on column insert, refs above the insertion left untouched,
  range expansion when the band falls inside the range,
  multiple-ref formulas, cross-sheet refs unaffected when
  `adjust` is scoped to a different sheet, deleted single-cell
  ref → `#REF!`, deleted range entirely → `#REF!`, partial
  range overlap shifts correctly, no-op when the adjust is a
  pass-through, `hasRefError` and `changed` accounting.
- `packages/xlsx/src/commands/insert-row.test.ts` — **13 tests**
  covering shifted cell positions, formula range expansion,
  cross-sheet ref untouched, merge straddling a boundary
  expands `r2`, merge fully below shifts down, merge straddling
  strictly inside is rejected, `rows-inserted` summary diff
  shape, validation (`unknown-sheet`, `invalid-position`,
  `invalid-count`, `invalid-position` when `at + count`
  exceeds the row limit), and a full parse → insert → serialize
  → re-parse round-trip preserving cell values and formulas.
- `packages/xlsx/src/commands/insert-column.test.ts` — **9 tests**
  on the column axis: cell shift, formula range expansion,
  merge boundary expansion vs. straddle rejection,
  `columns-inserted` summary, validation (unknown sheet,
  invalid position / count, exceeds column limit).
- `packages/xlsx/src/commands/delete-row.test.ts` — **14 tests**:
  cells inside the band dropped, cells below shifted up, cells
  above unchanged, formula partial-range shift with the
  recalculated cached value verified, deleted single-cell ref
  rewritten to `#REF!` with the cached value flipping to a
  `#REF!` `CellError`, `referenced-cell-deleted` change emitted
  for casualties, cross-sheet refs targeting a different sheet
  untouched, merges fully inside dropped, merges below shifted,
  merges straddling rejected with `merge-boundary-crossed`,
  `rows-deleted` summary diff, and validation (`unknown-sheet`,
  `invalid-position`, `invalid-count`).
- `packages/xlsx/src/commands/delete-column.test.ts` — **13 tests**
  mirroring delete-row on the column axis: cell drop / shift /
  preserve, formula partial-range shift with verified cached
  value (range `U1:Y1` deleting `V:W` becomes `U1:W1` summing
  the surviving `U`, `X→V`, `Y→W` = `1+4+5 = 10`), single-cell
  ref → `#REF!`, casualty diff, cross-sheet refs untouched,
  merge inside dropped / right-of shifted left / straddle
  rejected, summary diff, and validation.

**Totals after 7i**

- `@officeai/xlsx` unit tests: **596 passing** (+80).
- 12/13 P0 commands now wired through the bus.
- `pnpm --filter @officeai/xlsx test`, `lint`, `build`, and
  `pnpm format:check` all green.

**Decisions**

- **Single shared engine for all four commands.** Insert and
  delete on rows and columns are the same operation modulo (a)
  which axis is being shifted, (b) the sign of the shift, and
  (c) the merge-precheck rule. Splitting the implementation
  across four files would have meant four near-duplicate
  ~250-line pipelines and four separate places to fix any
  precedence / merge / `#REF!` bug. `structural-shift.ts`
  parameterizes on `axis: "row" | "column"` and
  `op: "insert" | "delete"` and dispatches to the right
  `adjustForXxx` from `formula/references.ts`. The four
  command handlers are 8 lines each.
- **Workbook-wide formula rewrite, not just the target sheet.**
  A formula on `Sheet2` like `=Sheet1!A1` must be rewritten
  when row 1 of `Sheet1` is deleted. The rewrite pass iterates
  every cell on every sheet, not just the sheet whose
  structure is changing. The `adjustForXxx` functions
  short-circuit on a sheet mismatch so the per-cell cost on
  unrelated sheets is one comparison.
- **Re-emit, don't byte-preserve, rewritten formulas.**
  `serializeAst` produces canonical text that is **not**
  byte-identical to the original source (whitespace stripped,
  function names uppercased, defensive parens around binary
  children). Round-tripping the rewritten text through
  `parse → evaluate` yields the same value, which is the
  contract the formula engine actually needs. Preserving
  source bytes would have required carrying span / trivia
  through the AST, which is out of scope for the 80% target.
- **`#REF!` is a literal, not a special node kind.** When a
  ref is fully inside the deletion band the rewriter replaces
  the `Reference` / `RangeReference` node with a `Literal`
  carrying `err({ kind: "#REF!" })`. Re-parsing
  `serializeAst` on that literal yields a `Reference` to a
  defined-name `#REF!` per the lexer rules, which the
  evaluator surfaces as the same error — keeping the AST
  union closed without a separate `RefError` node.
- **Merge precheck differs by op.** Insertions reject only
  when the insertion index splits a merge **strictly inside**
  (`lo < at < hi`); landing on a boundary cleanly extends or
  shifts the merge. Deletions reject whenever the band
  partially overlaps a merge — i.e., the merge extends outside
  the band on either side — because there's no defensible
  "shrink the merge" rule in the spec. Both paths emit
  `merge-boundary-crossed` with `meta.range` carrying the
  offending merge's A1.
- **Defined names, comments, hyperlinks: deferred.** The
  typed model doesn't yet carry defined names, comments, or
  hyperlinks (they live in opaque parts), so this phase
  doesn't adjust them. When those models land we'll thread
  the same `adjust` functions through their rewriters.

### Phase 7j — `xlsx:add-comment` (2026-04-18)

**Goal**

Land the final P0 command (13/13). Attach a classic note to a
single cell, minting `xl/comments{N}.xml` and the matching
per-sheet rels + content-types overrides on the first comment
per sheet.

**Code landed**

- `packages/xlsx/src/model/types.ts` — added a typed
  `Comment` interface (`id`, `ref`, `author`, `text`, optional
  `parentId` reserved for P1 threaded replies) plus three
  fields on `Sheet`: `comments`, `commentsPartPath?`, and
  `commentAuthors`. The authors array preserves insertion order
  so `authorId` indices remain stable across the
  parse → mutate → serialize → re-parse cycle.
- `packages/xlsx/src/parser/comments.ts` — `parseCommentsPart`
  reads `<authors>` and `<commentList>` from a comments part,
  flattens `<r><t>` runs into a single plain-text `text` per
  comment (rich-text formatting collapses by design in P0),
  and mints positional `comment-N` ids.
- `packages/xlsx/src/parser/parse.ts` — every worksheet now
  loads its rels graph (when present), resolves any
  `…/relationships/comments` rel via `resolveTargetPath`, and
  attaches the parsed `comments` + `commentAuthors` +
  `commentsPartPath` to the typed `Sheet`. Sheets without a
  comments rel default to empty arrays. `xl/comments*.xml`
  parts are now modeled (removed from `opaqueParts` like
  `xl/styles.xml` was); `xl/threadedComments/*` was demoted
  out of `MODELED_PREFIXES` so it surfaces in `opaqueParts`
  again — matching the "opaque, byte-preserved" P0 contract
  for threaded comments.
- `packages/xlsx/src/serializer/comments.ts` —
  `serializeCommentsPart(authors, comments)` emits a canonical
  `<comments xmlns=…><authors>…</authors><commentList>…</commentList></comments>`,
  XML-escaping author names + text bodies and tagging text
  runs with `xml:space="preserve"` so leading/trailing
  whitespace round-trips.
- `packages/xlsx/src/serializer/serialize.ts` — dropped
  `comments` and `sheetRels` from the `unsupportedDirty`
  guard, then added two rewrite paths:
  1. `rewriteDirtyComments` looks up each dirty comments path
     by `sheet.commentsPartPath` and writes the canonical XML.
  2. `rewriteDirtySheetRels` matches each dirty rels path to
     its owning sheet via `RelationshipGraph.relsPathFor(sheet.partPath) === relsPath`,
     loads the existing rels (preserving any pre-existing
     hyperlink / vmlDrawing / drawing rels verbatim), drops
     all `…/relationships/comments` entries, then re-adds a
     single comments rel pointing at `sheet.commentsPartPath`
     using a `relsRelativeTarget` helper that emits Excel's
     canonical `../comments1.xml` form.
     `rewriteContentTypes` was extended to also ensure an
     `application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml`
     override exists for every sheet's `commentsPartPath`.
- `packages/xlsx/src/commands/add-comment.ts` — handler:
  resolve sheet → reject ranges via `payload.ref.includes(":")`
  - parse single-cell ref → reject `empty-text` / `empty-author`
    → reject `comment-exists` if the cell already has a comment
    → de-dupe author append → mint `comment-N` id → if first
    comment for the sheet, mint a fresh `xl/comments{N}.xml`
    path that doesn't collide with any other sheet's
    `commentsPartPath` or any container key, and dirty
    `sheetRels` + `contentTypes` in addition to `comments`. The
    diff is one `node-inserted` for the comment plus, on first
    comment, a second `node-inserted` for the new part.
- `packages/xlsx/src/commands/payloads.ts` — added
  `AddCommentPayload { sheet, ref, text, author }`.
- `packages/xlsx/src/commands/registry.ts` — re-ordered the
  handler list to match the spec's §1–§13 sequence and wired
  in `addCommentHandler`. Comment in the file is now
  "13/13 P0 commands shipped"; the deferred section is gone.
- `packages/xlsx/src/commands/index.ts` and
  `packages/xlsx/src/index.ts` — re-exported the new handler,
  payload, and `Comment` type.
- `packages/xlsx/src/commands/add-sheet.ts` — fresh sheets now
  initialize `comments: []` and `commentAuthors: []` so the
  typed surface stays exhaustive.

**Tests landed**

- `packages/xlsx/src/parser/__tests__/comments.test.ts`
  — **6 tests** covering (a) extraction of authors + comments
  with stable `comment-N` ids and `<r><t>` run concatenation,
  (b) empty `<authors/><commentList/>` shape, (c) out-of-range
  `authorId` falling back to empty author, (d) round-trip
  parse → serialize → parse equivalence on the typed tuple,
  (e) XML escaping of `& < > "` in both authors and bodies,
  (f) `xml:space="preserve"` retention of leading/trailing
  whitespace.
- `packages/xlsx/src/commands/add-comment.test.ts` —
  **11 tests**: first-comment path mints a new
  `xl/commentsN.xml`, dirties sheet-rels + content-types,
  attaches the typed `Comment`, and emits two `node-inserted`
  changes; second comment on the same sheet reuses the same
  part path and emits exactly one diff change; same-author
  dedupe; rejection of `comment-exists`, `empty-text`,
  `empty-author`, `unknown-sheet`, and two flavours of
  `invalid-ref` (range and malformed); round-trip on a clean
  fixture preserving all original cell values; round-trip on
  the `05-comments-hyperlinks.xlsx` fixture preserving every
  pre-existing comment + author and surfacing the new entry
  with the right ref / author / text.

**Decisions**

- **Threaded comments stay opaque.** The "modern" comments
  rewrite (`xl/threadedComments/*` plus the `personList.xml`
  metadata) layers GUIDs, reply chains, and author personas
  on top of the classic `<comment>` shape. P0's mandate is to
  ship the 13 named commands, not the surface they live on,
  so we keep those parts byte-preserved via `opaqueParts` and
  defer threaded reply / resolve / unresolve to P1's
  `xlsx:reply-comment` / `xlsx:resolve-comment` lineup.
- **VML drawings (`xl/drawings/vmlDrawing*.vml`) deferred.**
  Excel anchors classic comments visually with a sibling
  `<v:shape>` document referenced via a `…/vmlDrawing` rel and
  a `<legacyDrawing>` element on the worksheet. Without it,
  a freshly added comment round-trips through the data layer
  but isn't pinned to a screen position when Excel renders the
  workbook. Acceptable for the headless P0 surface (whose
  consumers are agents, not human Excel users) and avoids
  modeling VML — a pre-OOXML markup that would otherwise
  spread its own parser / serializer through the codebase. P1
  will add the VML emission alongside `xlsx:edit-comment`.
- **Plain-text body in P0.** `<text>` can carry rich-text
  runs (`<r><rPr/><t/></r>` chains). Modeling rich text would
  duplicate work already lined up for the formatting layer
  (`xlsx:set-cell-format` already owns the typed style
  surface). For P0 we flatten incoming runs to a single
  `text: string` and re-emit a single `<r><t xml:space="preserve">`
  on serialize. Round-trip on the synthetic fixture is exact;
  on a workbook with hand-formatted comments the text content
  is preserved but in-comment formatting (bold spans, colour
  runs) collapses to plain text.
- **Author indices via insertion order, not by reference.**
  `commentAuthors` is appended in command-call order, so the
  index that `<comment authorId="N">` points at stays stable
  across mutations. We never re-sort or compact the authors
  array; deletion of a comment is a P1 concern and will keep
  the rule even when a referenced author is left dangling.
- **`relsRelativeTarget` over hard-coded `../comments1.xml`.**
  The helper computes the rel target the way Excel does: walk
  up from the owning rels file's directory until the path
  shares a prefix with the target, then descend. Today every
  sheet lives in `xl/worksheets/` and every comments part in
  `xl/`, so we always emit `../commentsN.xml` — but the
  helper is correct under any future layout (e.g. nested
  workbooks) without rework.

### Phase 7 — closure (2026-04-18)

All 13 P0 commands now ship through the bus:

| §   | Command                 | Phase | Notes                                                                 |
| --- | ----------------------- | ----- | --------------------------------------------------------------------- |
| 1   | `xlsx:set-cell-value`   | 5     | Inline-string emission keeps `xl/sharedStrings.xml` byte-stable.      |
| 2   | `xlsx:set-cell-formula` | 7f    | Full workbook recalc + cached writebacks; `EC-F1` cycle reporting.    |
| 3   | `xlsx:set-range-values` | 5     | Row-major matrix; rectangular dimension enforcement.                  |
| 4   | `xlsx:set-cell-format`  | 7g    | Typed `StyleTable` + content-hash dedupe; semantic styles round-trip. |
| 5   | `xlsx:insert-row`       | 7i    | Cell shift + workbook-wide formula rewrite + recalc.                  |
| 6   | `xlsx:insert-column`    | 7i    | Same pipeline, column axis.                                           |
| 7   | `xlsx:delete-row`       | 7i    | Surfaces `#REF!` casualties via `referenced-cell-deleted`.            |
| 8   | `xlsx:delete-column`    | 7i    | Same surfacing, column axis.                                          |
| 9   | `xlsx:merge-cells`      | 5     | 0-based inclusive `MergedCell` regions.                               |
| 10  | `xlsx:unmerge-cells`    | 5     | Exact range match required.                                           |
| 11  | `xlsx:add-sheet`        | 7h    | Workbook + content-types + rels rewrite; SheetJS sync.                |
| 12  | `xlsx:rename-sheet`     | 5     | Cross-sheet formula rewriting deferred to P1.                         |
| 13  | `xlsx:add-comment`      | 7j    | Classic notes; threaded comments + VML deferred to P1.                |

`@officeai/xlsx` unit tests: **613 passing** (+17 over Phase 7i).
Workspace integration tests: **53 passing** (incl. xlsx
roundtrip + agent-edits roundtrip). `pnpm --filter @officeai/xlsx
test`, `lint`, `build`, and `pnpm format:check` all green.

P1 surface (out of scope for this build):
`xlsx:edit-comment`, `xlsx:delete-comment`,
`xlsx:reply-comment`, `xlsx:resolve-comment`,
`xlsx:reorder-sheet`, `xlsx:delete-sheet`,
`xlsx:hide-sheet`, `xlsx:set-defined-name`,
`xlsx:set-hyperlink`, plus the VML drawing emission needed for
visual comment anchoring in Excel proper.

### Phase 8 — office-agent CLI + MCP xlsx\*\* surface (2026-04-18)

Wires the headless `XlsxAgent` through both interactive surfaces with
DOCX parity.

**CLI** (`packages/agent/src/cli-xlsx.ts`, new):

- New `office-agent xlsx` subcommand group: `inspect`, `read`,
  `search`, `set-cell`, `set-formula`, `set-range`, `set-format`,
  `add-sheet`, `rename-sheet`, `insert-row|column`,
  `delete-row|column`, `merge|unmerge`, `add-comment`, `apply` /
  `apply-file`, and on-disk `diff`.
- Common helpers (`IO`, `CliError`, `stringifyJson`, `parseIntOpt`)
  factored into `cli-shared.ts` so `cli.ts` and `cli-xlsx.ts` don't
  need to import each other.
- `office-agent xlsx` no longer emits the "not implemented" stub;
  pptx still parks the next format.

**MCP** (`packages/agent/src/mcp.ts`):

- 22 new tools: `xlsx_load`, `xlsx_save`, `xlsx_inspect`,
  `xlsx_list_sheets`, `xlsx_get_text`, `xlsx_search`,
  `xlsx_apply_command`, `xlsx_list_pending`, `xlsx_approve`,
  `xlsx_reject`, `xlsx_diff`, plus the convenience wrappers
  (`xlsx_set_cell`, `xlsx_set_formula`, `xlsx_set_range`,
  `xlsx_set_format`, `xlsx_add_sheet`, `xlsx_rename_sheet`,
  `xlsx_{insert,delete}_{row,column}`, `xlsx_merge`, `xlsx_unmerge`,
  `xlsx_add_comment`).
- Each convenience tool collapses to `xlsx_apply_command` so there is
  one code path for every typed write.
- Reset hook now clears both `docxSessions` and `xlsxSessions` for
  test isolation.

**Tests**: agent suite 47 / 47 (+15: 7 CLI round-trips, 8 MCP tool
flows). DOCX surface untouched, still 32 / 32 green. `pnpm
--filter @officeai/agent test`, `lint`, `build`, and `pnpm
format:check` all green.

### Phase 9 — virtualized browser grid + XLSX web app surface (2026-04-18)

Mounts the headless `XlsxAgent` behind a real Excel-shaped surface
in the existing Next.js app.

**New route** `/xlsx-editor` (`apps/web/app/xlsx-editor/`):

- `page.tsx` — client wrapper, dynamic-imports `XlsxEditor` with
  `ssr: false`, mirrors the chrome of `/editor`.
- `Grid.tsx` (296 LOC) — hand-rolled virtualized grid. Fixed
  geometry (24 × 88 px), 1000 × 26 visible cells, OVERSCAN = 4,
  scroll-tracked sticky headers (column letters + row numbers +
  corner), in-cell editing on double-click with Enter / Escape
  commit. No 3rd-party grid library.
- `XlsxEditor.tsx` (436 LOC) — main surface. Header strip
  (`sample.xlsx` label + `rev N` + pending-mutation badges + Save),
  formula bar with cell-ref pill + `fx` input, sheet tabs, the
  `Grid`, and an "Agent" prompt row. Mounts `XlsxAgent.fromBuffer`
  on a synthetic seed workbook, subscribes to mutations,
  dispatches `xlsx:set-cell-value` for plain text and
  `xlsx:set-cell-formula` for `=`-prefixed input.

**Synthetic seed** (`apps/web/app/lib/sample-xlsx.ts`, 88 LOC):
JSZip-built `sample.xlsx` (Sheet1: `Name | Score`, `Alex | 42`,
`Sam | 37`, `Total | =SUM(B2:B3)`). Uses `t="inlineStr"` to skip
`xl/sharedStrings.xml`. Verified end-to-end: opens, edits, saves,
re-opens through the agent.

**Landing page**: home page now ships a second "Open the XLSX
editor" CTA alongside the DOCX one. Copy updated to
"DOCX editor: live. XLSX editor: live."

**Playwright smoke** (`apps/web/e2e/xlsx-editor.spec.ts`): loads
`/xlsx-editor`, asserts the seeded `Score` / `Alex` cells, clicks
A2, verifies the formula bar displays `Alex`, types `Bob` + Enter,
asserts the grid re-renders to `Bob` and the revision badge ticks
0 → 1. The spec is registered under `playwright.config.ts`'s
default `e2e/*.spec.ts` glob.

**Browser smoke** (manual via the cursor IDE browser MCP, against
`pnpm --filter @officeai/web dev` on :3001):

- `/xlsx-editor` boots, the seeded sheet renders with `B4 = 79`
  (`=SUM(B2:B3)` evaluated by the formula engine).
- Click A2 → formula bar shows `Alex`. Fill `Bob` + Enter →
  grid shows `Bob`.
- Click B2 → fill `=B3*2` + Enter → B2 becomes `74`, B4 cascades
  from `79` to `111`, confirming the recalc orchestrator wires
  through the live editor (not just the headless tests).

**Decisions**:

- **No 3rd-party grid library.** Virtualization is hand-rolled
  with fixed cell geometry + scroll tracking + absolute
  positioning. Keeps the dependency tree small and matches the
  "everything goes through our command bus" discipline of the
  DOCX editor.
- **Inline strings in the seed.** Skipping `sharedStrings.xml`
  in the synthetic workbook keeps the demo file under 100 LOC and
  exercises the same "sparse-edit, byte-preserve untouched parts"
  path the real fixtures hit.
- **Pending-mutation badge, not a panel.** The header surfaces a
  count; full approve / reject UI is a P1 polish task —
  `xlsx_list_pending` / `xlsx_approve` / `xlsx_reject` already work
  through the MCP surface for now.
- **Keyboard navigation deferred.** Click + formula-bar editing is
  the documented minimum; arrow-key cell traversal is a P1 polish
  follow-up that doesn't change the agent contract.

### Phase 10 — close-out (2026-04-18)

**Test totals (xlsx-relevant)**:

| Suite                                                   | Count    | Status |
| ------------------------------------------------------- | -------- | ------ |
| `@officeai/xlsx`                                        | 613 / 27 | green  |
| `@officeai/agent`                                       | 47 / 2   | green  |
| `@officeai/core`                                        | 12 / 2   | green  |
| `@officeai/integration-tests` (incl. `roundtrip/xlsx/`) | 51 / 7   | green  |

`pnpm --filter @officeai/xlsx {lint,build,test}` and the agent /
core / integration suites all pass. Web-app `pnpm exec eslint` on
the new Phase 9 files (`app/xlsx-editor/**`, `app/lib/sample-xlsx.ts`,
`app/page.tsx`, `e2e/xlsx-editor.spec.ts`) is clean.

**Browser smoke** (manual via `cursor-ide-browser` MCP against
`pnpm --filter @officeai/web dev` on :3001):

- Landing page (`/`) renders both **Open the DOCX editor** and
  **Open the XLSX editor** CTAs.
- `/xlsx-editor` boots the synthetic `sample.xlsx`. The seeded
  formula `=SUM(B2:B3)` evaluates to `B4 = 79` on first paint,
  confirming the formula engine fires through the live editor.
- Click A2 → formula bar shows `Alex`. Fill `Bob` + Enter → grid
  updates to `Bob`.
- Click B2 → fill `=B3*2` + Enter → B2 = 74 and B4 cascades from
  79 → 111, confirming the recalc orchestrator + dependency graph
  ride along the live editor (not just the unit harness).

**Known not-blocking**: `make verify` currently fails because of
unrelated parallel docx WIP (P3.1 / P3.2 — style cascade, paged
renderer, set-paragraph-spacing). Errors live entirely in
`packages/docx/src/{commands/set-paragraph-spacing.ts,
renderer/doc-to-pm.ts, serializer/serialize.ts}` plus the matching
WIP in `apps/web/app/editor/{Toolbar.tsx, DocxEditor.tsx}` and
`apps/web/app/lib/format-helpers.ts`. None of those files are
touched by the XLSX build; the XLSX-only gate (xlsx + agent + core

- integration-tests + web-eslint of xlsx-editor files) is green.

**README**: refreshed for the new state — XLSX flipped from
"deferred" to "active", monorepo layout note adds the xlsx package

- tests + build log, CLI section gains an `office-agent xlsx`
  example, reading order now points at `spec/xlsx/` and
  `docs/build-log/xlsx.md`.

**Sequence delivered (Phases 0–10)**:

| Phase | Deliverable                                                             | Tests delta         |
| ----- | ----------------------------------------------------------------------- | ------------------- |
| 0     | Scaffold `@officeai/xlsx`, architecture wiring                          | -                   |
| 1     | Four parallel clean-room analyses + synthesis                           | -                   |
| 2     | Spec corpus (`agent-commands.md`, `formula-engine.md`, `model.md`, ...) | -                   |
| 3     | Synthetic fixtures (`fixtures/xlsx/01-09`)                              | -                   |
| 4     | Round-trip oracle (parser + serializer skeleton)                        | +12 roundtrip       |
| 5     | Typed cell model + 5 / 13 P0 commands                                   | +35                 |
| 6     | Headless `XlsxAgent` (DocxAgent parity) + diff module                   | +63 unit + 17 int.  |
| 7a    | Formula primitives (tokens, errors, values, refs)                       | +101                |
| 7b    | Lexer + AST + precedence-climbing parser                                | +34                 |
| 7c    | Function registry + tree-walk evaluator                                 | +27                 |
| 7d    | Dependency graph (Tarjan SCC) + recalc orchestrator                     | +11                 |
| 7e    | Function library (89 P0 funcs across 5 parallel agents)                 | +302                |
| 7f    | `xlsx:set-cell-formula`                                                 | +11                 |
| 7g    | `xlsx:set-cell-format` + typed style table                              | +25                 |
| 7h    | `xlsx:add-sheet`                                                        | +14                 |
| 7i    | `xlsx:{insert,delete}-{row,column}` + AST serialize + ref rewrite       | +49                 |
| 7j    | `xlsx:add-comment` + classic notes parser/serializer                    | +17                 |
| 8     | `office-agent xlsx` CLI subcommands + 22 `xlsx_*` MCP tools             | +15 (7 CLI + 8 MCP) |
| 9     | Virtualized grid + `/xlsx-editor` web surface + Playwright smoke        | +1 e2e              |
| 10    | Browser smoke, README refresh, build-log close-out                      | -                   |

**Out of scope for this build (queued for P1)**: threaded comments,
VML drawing emission, cross-sheet formula rewriting on rename,
defined names, hyperlinks, sheet reorder/delete/hide,
`xlsx:edit-comment` / `xlsx:delete-comment` / `xlsx:reply-comment`
/ `xlsx:resolve-comment`, rich-text comment runs, and arrow-key
keyboard navigation in the web grid.

---

## Phase 11 — Excel-flavoured UX on `/xlsx-editor`

The headless engine landed in Phases 0–10; Phase 11 turns the web
surface into something that actually feels like Excel. All
mutations still flow through `XlsxAgent.applyCommand`; the changes
here are purely UX + a couple of new commands for column/row
sizing.

| Sub | Deliverable                                                                | Tests delta             |
| --- | -------------------------------------------------------------------------- | ----------------------- |
| 11a | Open .xlsx from disk + drag-drop (replace-agent on file load)              | +1 e2e                  |
| 11b | Multi-cell selection (anchor/focus, drag-extend, shift-click, marquee)     | +2 e2e                  |
| 11c | Type-to-edit + click-to-insert-ref while editing formulas                  | +3 e2e                  |
| 11d | Formula autocomplete popover (export `listRegisteredFunctions`, Tab)       | +3 e2e + 7 unit         |
| 11e | Rich styling toolbar (font/align/fill/number-fmt) + Grid renders `styleId` | +4 e2e                  |
| 11f | Merge / Unmerge / Insert / Delete from selection, merged-cell rendering    | +3 e2e                  |
| 11g | `xlsx:set-column-width` / `xlsx:set-row-height` + drag handles + variable geometry | +2 e2e + 7 unit |

**New commands**: `xlsx:set-column-width`, `xlsx:set-row-height`.
Both store the override on `Sheet.columnWidths` / `Sheet.rowHeights`
(0-based key, CSS pixels). `null` resets to default (`COL_WIDTH = 100`,
`ROW_HEIGHT = 24`). Out-of-range column / row indices and sizes
outside `[MIN, MAX]` reject as `validation` mutations instead of
mutating the workbook.

**New web exports from `@officeai/xlsx`**: `listRegisteredFunctions`,
`flattenCellXf`, `EffectiveStyle`, `StyleTable` (+ supporting
style-table types), and the two new sizing payload / handler
exports. Used by the Toolbar, FormulaSuggest, and Grid components.

**Grid refactor (variable geometry)**: replaced fixed
`r * ROW_HEIGHT` arithmetic with prefix-sum arrays (`colXs`,
`rowYs`) memoised on `sheet.columnWidths`, `sheet.rowHeights`, and
the transient `colDrag` / `rowDrag` state. Visible-window math now
binary-searches the prefix arrays. Header drag handles dispatch
the new commands on mouse-up; transient drag preview is local
state so the grid stays responsive without round-tripping every
mousemove through the agent bus.

**Smoke (browser, 19 tests passing on `:3001`)**: open .xlsx →
fixture replaces seeded sample → multi-cell selection (shift-click,
plain click collapse) → type-to-edit + Backspace clear →
click-to-insert-ref into formula → autocomplete popover (Tab /
ArrowDown / Esc) → bold / italic / underline / align-right toggles
→ multi-cell format dispatch → merge / unmerge / insert-row /
delete-column → drag-resize column A and row 1 with revision tick
verification.

**Tests delta total for Phase 11**: +18 e2e + 14 unit. Repo total
xlsx tests now sit at **624 unit / 19 e2e** all green.
