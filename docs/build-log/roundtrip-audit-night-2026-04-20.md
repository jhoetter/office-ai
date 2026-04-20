# Round-trip attribute-fidelity audit — Night 2026-04-20

This phase delivered a developer-loop friendly audit script that
_counts attribute-class survival_ through a full parse → serialize
→ re-parse cycle for every fixture in the repository.

Unlike the existing `run-libreoffice-roundtrip.mjs` (which gates
"the file still opens cleanly in Office"), the audit specifically
verifies that the **typed snapshot** preserves the formatting
attributes the toolbar exposes — paragraph alignment, run
bold/italic/font/size/color, list ids, page setup, table cells,
chart presence, slide shapes, etc.

## How to run

```
make audit-roundtrip
# or, after the format packages are built:
node scripts/audit-roundtrip.mjs
```

Output is human-readable per fixture plus a machine-readable JSON
summary at `docs/build-log/roundtrip-audit-night.json` that can be
diffed against between runs.

## Result on the first run

Every single fixture round-tripped with **exact attribute match
counts** (no losses, no spurious gains). All 30 bundled fixtures
land at `attrs preserved == attrs after`:

| Format | Fixtures | Exact-match | Notes                                                         |
| ------ | -------- | ----------- | ------------------------------------------------------------- |
| DOCX   | 11/11    | 11/11       | incl. landscape section, multi-column, TOC SDT, callout table |
| XLSX   | 6/6      | 6/6         | incl. 7012-attribute large-grid sheet, formulas, merges       |
| PPTX   | 13/13    | 13/13       | incl. real-world decks, charts, animations, multi-shape       |

Headline numbers from the JSON summary:

- **DOCX 07-toc-sdt**: 927 attribute observations preserved
  exactly — the heaviest fixture in the corpus.
- **XLSX 06-large-grid**: 7012 cell/style/column observations
  preserved exactly.
- **PPTX 08-large-deck**: 150 shape/run/picture observations
  preserved exactly.

This is a strong signal that the work landed by Phase 4 (chart
round-trip) and the prior parser/serializer hardening did not
regress anywhere; the existing per-format vitest suites cover the
finer-grained shape, and this audit covers the workbook-level
reductions that those suites cannot.

## What attributes the audit covers (today)

### DOCX (`tallyDocx`)

```
paragraphs            run-italic
paragraph-alignment   run-font-family
paragraph-list        run-font-size
run-bold              run-color
run-highlight         text-leaves
drawings              page-size
page-margins          page-landscape
```

### XLSX (`tallyXlsx`)

```
sheets    cells       formulas
charts    images      merges
col-widths styled-cells fonts
fills     num-fmts
```

### PPTX (`tallyPptx`)

```
slides    shapes      pictures
charts    paragraph-alignment
run-bold  run-italic  run-font-size
run-color run-font-family
```

## Why count-based, not deep-equal?

A deep equality on snapshots would flag every cosmetic difference
(NodeId remint, attribute order, default vs explicit `xml:space`,
etc.) and drown out the signal we actually want — _did this
formatting attribute survive at all?_ Counting captures that
question precisely: a missing alignment shows up as "12 → 11", and
a duplication shows up as a positive `gain`. Both are flagged.

The vitest packages already do per-feature deep equality on
focused fixtures (e.g. `packages/docx/src/serializer/*.test.ts`),
so the two layers complement: vitest answers "does this single
attribute survive end-to-end?", the audit answers "does the
_aggregate_ across the whole corpus survive?".

## Backlog (cheap one-liners, deferred to follow-up)

- Add `font.color` and `font.fill` to the XLSX walker (they live
  on `EffectiveStyle`, one map lookup per cell).
- Hash a small set of "spot-check" cells/runs (e.g. first 5 styled
  cells) so a swap of two styled IDs that happens to net to zero
  in the totals would still be caught.
- Wire `audit-roundtrip` into the `heavy` make target so CI
  archives the JSON summary on every PR.
- DOCX: add a tally for `<w:hyperlink>` rels (we count `formulas`
  in XLSX; the equivalent for DOCX would be hyperlink ids).

## What broke during the audit (and was fixed)

Nothing. First-run pass.
