# XLSX Spec

The contract for the XLSX (Microsoft Excel) editor in
[`packages/xlsx`](../../packages/xlsx).

## Reading order

1. [`analysis.md`](analysis.md) — clean-room synthesis of the Univer +
   SheetJS reference repos. Decides what we keep, what we differ on,
   what we improve. Drawn from the four `analysis-*.md` companion notes.
2. [`feature-scope.md`](feature-scope.md) — what is **in** the 80% and
   what is **explicitly out**. No ambiguity.
3. [`document-model.md`](document-model.md) — `XlsxWorkbook`, `Sheet`,
   `Cell`, `RangeRef`, `CellFormat`, `Comment`, `DefinedName`, etc.
4. [`ooxml-mapping.md`](ooxml-mapping.md) — for every in-scope feature:
   which OOXML element / attribute maps to which model field.
5. [`parser.md`](parser.md) — OOXML → `XlsxWorkbook`, including
   opaque-blob preservation for parts we don't model.
6. [`serializer.md`](serializer.md) — `XlsxWorkbook` → OOXML, with
   byte-preservation of untouched parts driven by `partHashes`.
7. [`renderer.md`](renderer.md) — virtualized DOM grid, sheet tabs,
   formula bar, frozen panes; everything is a command.
8. [`formula-engine.md`](formula-engine.md) — lexer, AST, dependency
   graph, evaluator, error model, the ~150-function priority list.
9. [`agent-commands.md`](agent-commands.md) — every `xlsx:*` command:
   payload, behaviour, OOXML impact, examples.
10. [`edge-cases.md`](edge-cases.md) — known hard cases and how we
    degrade gracefully.
11. [`acceptance-criteria.md`](acceptance-criteria.md) — measurable
    done criteria the build must hit before XLSX is declared shipped.

The shared specs in [`../shared/`](../shared) cover infrastructure
(command bus, OOXML utils, agent API, plugin system) which is reused
unchanged from the DOCX phase. The headless `XlsxAgent`
(`packages/xlsx/src/agent/`) implements the same `DocumentAgent`
shape as `DocxAgent`, mirroring methods and pending/approved
semantics; see `spec/shared/agent-api.md` and
[`docs/build-log/xlsx.md` Phase 6 entry](../../docs/build-log/xlsx.md).

## Status

| Doc                      | Status                                                                                      |
| ------------------------ | ------------------------------------------------------------------------------------------- |
| `analysis.md`            | landed (P1)                                                                                 |
| `feature-scope.md`       | landed (P2)                                                                                 |
| `document-model.md`      | landed (P2)                                                                                 |
| `ooxml-mapping.md`       | landed (P2)                                                                                 |
| `parser.md`              | implemented (P4 thin + P5 typed cells)                                                      |
| `serializer.md`          | implemented (P4 byte-oracle + P5 dirty-sheet rewrite)                                       |
| `renderer.md`            | landed (P2)                                                                                 |
| `formula-engine.md`      | landed (P2)                                                                                 |
| `agent-commands.md`      | partial (5/13 implemented in P5; remaining 8 deferred to P7 — see `docs/build-log/xlsx.md`) |
| `edge-cases.md`          | landed (P2)                                                                                 |
| `acceptance-criteria.md` | landed (P2)                                                                                 |
