# XLSX — Reference-Repo Analysis (clean-room synthesis)

> Canonical synthesis of the four `analysis-*.md` companion notes. This
> document is the bridge between "what the references do" and "what we
> ship in `spec/xlsx/`". It is also the definitive list of decisions
> recorded in this phase: every later spec doc may cite back to a
> section here.
>
> Companion notes (read for citations / depth, do not duplicate here):
>
> - [`analysis-univer-core.md`](analysis-univer-core.md) — Univer architecture, plugins, command/mutation/operation split, range model, skeleton layer
> - [`analysis-univer-formula.md`](analysis-univer-formula.md) — Univer formula engine: lexer-tree → shunting-yard → AST → walking interpreter, range-indexed dependency graph, error model
> - [`analysis-sheetjs.md`](analysis-sheetjs.md) — SheetJS Community Edition: what it parses/serializes vs what it silently drops; the "TODO part" inventory
> - [`analysis-agent-patterns.md`](analysis-agent-patterns.md) — `dream-num/univer-mcp` (doc-only) + `dream-num/skills` (sheet-git, agent-sheet); proposed `xlsx_*` MCP surface
>
> Legal: every analysis pass was clean-room. No code was copied, lightly
> renamed, or transcribed beyond ≤8-line pseudocode snippets where
> needed to convey a pattern. SheetJS Community Edition is permitted as
> a runtime dependency under Apache 2.0 (per `prompt.md` line 45).

## 1. Architecture decisions for `@officeai/xlsx`

### 1.1 What we KEEP from the references

From **Univer core** ([details](analysis-univer-core.md)):

- Three-tier mutation pipeline — `Command` (orchestrates) → `Mutation` (pure, sync, replayable) → `Operation` (ephemeral UI). We already have this in `@officeai/core`'s `CommandBus` + `MutationStore`; the only addition for XLSX is to keep the inverse-pairing discipline (every mutation produces its own `undo` mutation as a pure function of `before`).
- Sparse-cell storage: an `ObjectMatrix<ICellData>` per sheet. We adopt this as `Map<string, Cell>` keyed by `${row}:${col}` (zero-based internal). Trades O(1) lookup for slightly worse iteration; this is correct for the agent-write-heavy workloads we target.
- A single canonical `IRange` shape with absolute-ref flags (`AbsoluteRefType`) carried out-of-band so we can roundtrip `=$A$1` semantics through index storage.
- A central `RefRangeService` pattern — one place every feature with a "range" plugs into so insert/delete row+column adjustment happens in one place. We will ship this as `IndexedRanges<T>` in `packages/xlsx/src/model/`.
- Headless geometry layer (`SheetSkeleton`) lives in core, separately from the rendering surface. Ours will be a **pure function** (their version is 1100 lines and depends on DI, which is wrong for a geometry layer).

From **Univer formula** ([details](analysis-univer-formula.md)):

- Lexer-tree → shunting-yard → AST → tree-walking interpreter pipeline.
- First-class error values (`ErrorValueObject`) that short-circuit through arithmetic.
- Forward-graph dependency tracking with a range index (R-tree-ish in Univer; we'll start with a simpler interval index keyed by sheet + row band — sufficient for our 80%).
- Volatile-function set (`RAND/RANDBETWEEN/NOW/TODAY`) that's force-dirty on every recalc.
- `si + x + y` shared-formula model for OOXML round-trip.
- `AbsoluteRefType` enum for insert/delete reference adjustment.

From **SheetJS** ([details](analysis-sheetjs.md)):

- Use `XLSX.read(buf, { dense: true, cellFormula: true, cellStyles: true, cellNF: true, sheetStubs: true, bookFiles: true, bookVBA: true, xlfn: true, cellDates: false })` to extract the cell layer cheaply and correctly.
- Use `XLSX.writeXLSX(wb, { bookSST: true })` only when re-emitting _modified_ sheets.
- The cell shape `{ t, v, f, F, w, z, s }` is sufficient for our model translation.
- The 25 built-in number-format IDs + remap table — we'll mirror these so we don't reinvent format-string semantics.

From **agent patterns** ([details](analysis-agent-patterns.md)):

- The `(value × raw × formula) × (json × csv × tsv)` orthogonal split for range reads — three projections, three encodings, no special cases.
- A1-with-required-sheet-prefix at the public surface (`Sheet1!A1:B5`); numeric indices are 1-based at the boundary, 0-based only internally. Avoids Univer's documented 0-based footgun.
- Per-tool inline `{ revision, diff }` returns from every write tool — the diff replaces a separate "what changed?" round-trip.

### 1.2 What we DIFFER ON

| Topic                   | Univer / SheetJS                                                      | OfficeAI XLSX                                                                                        |
| ----------------------- | --------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------- |
| Source of truth         | JSON `IWorkbookData`; OOXML is exported on demand                     | OOXML is the source of truth; `XlsxWorkbook` is a working surface                                    |
| Round-trip discipline   | SheetJS regenerates every part; Univer has no XLSX I/O at all         | Untouched parts are byte-preserved via `partHashes` (same trick `@officeai/docx` uses)               |
| Style storage           | Univer: content-addressed shared `Styles` map; SheetJS: `cellXfs`     | We keep OOXML's `cellXfs` table on disk; in-memory we expose a content-hashed style id (collab-safe) |
| Plugin system           | Univer: `@wendellhu/redi` DI + topo-sorted plugins                    | Ours: explicit `docxPlugin`-style registration on the bus. No DI container.                          |
| Mutation typing         | Univer: string-id + `unknown` params                                  | Discriminated-union typed payloads, schema-validated at the bus boundary                             |
| Async / iterative calc  | Univer supports both                                                  | Sync only. No iterative calc. Circular refs surface as a structured error.                           |
| Renderer                | Canvas (Univer) / virtual-DOM grid (others)                           | DOM grid with windowing (mirrors how `@officeai/docx` renders — pure DOM, no canvas)                 |
| MCP transport           | univer-mcp: hosted, multimodal, proprietary                           | Ours: stdio MCP server reusing the existing `packages/agent/src/mcp.ts` shell                        |
| Approval surface        | univer-skills `sheet-git`: Git-shaped per-commit review               | Inline mutation diff per call + the existing approved/pending/working tri-state in `MutationStore`   |
| Formula engine ambition | Univer: ~400+ functions, lambda, LET, dynamic arrays, structured refs | OfficeAI: ~150 functions per `prompt.md` line 247. Array fns deferred to "if time"                   |

### 1.3 What we IMPROVE

- **Machine-readable command schemas.** Every `xlsx:*` command exports a Zod schema; the MCP tools and CLI both consume the same schema. LLMs see the validated payload, not a hand-written prompt template.
- **Inverse mutations as a discipline.** Every mutation handler ships an inverse alongside its forward; we add a property test asserting `apply(redo) ∘ apply(undo) === identity` for every mutation.
- **`precheck → { ok, reason, suggestedFix }`** for agent-driven preconditions: an LLM that tries `xlsx:merge-cells` on an already-merged range gets a structured "no, but here's what would work" instead of an exception.
- **Inline `diff` on every mutation** — already in `@officeai/core`, but the XLSX projection adds a per-cell summary so an LLM can audit a 200-cell write without re-parsing the file.
- **Byte-preservation budget.** We commit to: any zip part our `parser.ts` did not deserialize round-trips bit-identical. Verification: per-fixture roundtrip test asserts `zip.parts.set(read).equals(zip.parts.set(write))` for every untouched part.
- **Headless-first.** The `XlsxAgent` runs in Node with zero DOM. No `Injector`, no `LocaleService`, no `ThemeService`. The renderer is a separate package surface.

## 2. Decisions captured for the spec phase

| Spec doc                 | Anchored decisions                                                                                                                                                                                                         |
| ------------------------ | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `feature-scope.md`       | 80% list per `prompt.md` lines 230-275; defer pivots/slicers/sparklines/Power Query/macros (preserve only); 150-fn formula priority list                                                                                   |
| `document-model.md`      | `XlsxWorkbook` (sheets[], styles, sst, definedNames, partHashes); `Sheet` (name, cells: Map, merges: IndexedRanges, conditionalFormats, hyperlinks, comments, freeze, tabColor); `Cell` (value, formula, styleId, comment) |
| `ooxml-mapping.md`       | Per-feature mapping table: `xl/workbook.xml`, `xl/worksheets/sheetN.xml`, `xl/sharedStrings.xml`, `xl/styles.xml`, `xl/comments*.xml`, `xl/_rels/*`, conditional formats, defined names, hyperlinks, drawings (opaque)     |
| `parser.md`              | Two-stage: SheetJS adapter for cells/formulas/SSF; OoxmlContainer for everything else as opaque `Uint8Array`. Hash every part. Algorithm in pseudocode.                                                                    |
| `serializer.md`          | Re-emit dirty parts only. Untouched parts: `OoxmlContainer.parts.get(path)` returns the original bytes. Drives byte-equality post-export for untouched parts.                                                              |
| `renderer.md`            | Virtualized DOM grid; row/col header sticky; formula bar; sheet-tab strip; frozen panes; every interaction → command on the bus. No canvas.                                                                                |
| `formula-engine.md`      | Layout: `lexer.ts → shunting-yard.ts → ast.ts → parser.ts → evaluator.ts → dependency-graph.ts → recalc.ts → errors.ts → values.ts → references.ts + functions/{math,logic,lookup,text,date,finance,info,array}.ts`        |
| `agent-commands.md`      | 13 P0 commands (Phase 5) + formula-aware variants (Phase 7). Zod schemas. Examples per command.                                                                                                                            |
| `edge-cases.md`          | Circular refs, formulas referencing deleted ranges, opaque-part conflicts, very large ranges (50k rows), Excel-vs-Office-vs-LibreOffice deviations                                                                         |
| `acceptance-criteria.md` | All synthetic fixtures roundtrip clean (untouched parts byte-equal); agent CLI smoke; formula correctness suite; perf budget (50k rows under 1 s parse, 16 ms scroll); browser smoke through `make dev`                    |

## 3. Risk register (carried into `docs/build-log/xlsx.md`)

1. **SheetJS opaque part coverage.** SheetJS silently drops pivots, tables (ListObjects), charts, drawings, conditional formatting, data validations, slicers, timelines, custom XML, themes>1, full styles. We capture all of these as opaque `Uint8Array` parts in `OoxmlContainer`. **Validation:** post-export zip inventory diff against import.
2. **Formula correctness.** Excel has ≈10k edge cases. Scope: 80%. We document deltas in `docs/build-log/xlsx.md` as we discover them (e.g. `DATEDIF` undocumented quirks, `VLOOKUP` exact-vs-approx semantics, locale-sensitive number parsing).
3. **Renderer perf.** `prompt.md` line 195: "50k rows smooth". Use windowing (visible viewport only). Budget asserted in a perf script analogous to `scripts/perf-docx.mjs`.
4. **CI for LibreOffice xlsx round-trip.** Extend `scripts/run-libreoffice-roundtrip.mjs` to include `fixtures/xlsx/synthetic/*.xlsx`.
5. **A1 vs R1C1 vs sheet-prefix mismatches.** Pin a single convention at the API boundary (A1 with required `Sheet1!` prefix for cross-sheet refs); add parser/validator helpers in `packages/xlsx/src/model/refs.ts`.

## 4. Subagent contributions

| Subagent | Output                       | Lines | Highlight                                                                                              |
| -------- | ---------------------------- | ----- | ------------------------------------------------------------------------------------------------------ |
| A        | `analysis-univer-core.md`    | 567   | Three-tier C/M/O split + skeleton layer + plugin lifecycle                                             |
| B        | `analysis-univer-formula.md` | 1156  | Lexer-tree → shunting-yard → AST → walker; range-indexed dep graph; volatile fns; `si+x+y` shared form |
| C        | `analysis-sheetjs.md`        | 897   | What SheetJS reads/writes vs the long "TODO" list of dropped parts → drives our opaque-blob list       |
| D        | `analysis-agent-patterns.md` | 700   | univer-mcp is doc-only; agent-sheet is the real reference; 21-tool `xlsx_*` MCP surface proposed       |

Total: ~3,300 lines of analysis ready to drive the spec phase.
