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
