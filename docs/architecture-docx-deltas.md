# DOCX architectural deltas vs XLSX

> Companion to [`session-summary.md`](session-summary.md) and
> [`build-log/docx.md`](build-log/docx.md), and a mirror of
> [`architecture-xlsx-deltas.md`](architecture-xlsx-deltas.md). The
> build log is chronological; this doc is _spatial_ — it explains the
> cross-cutting shape of the DOCX product against the XLSX baseline so
> a new contributor (or a future LLM agent) does not have to derive
> these by reading both stacks side-by-side.

The two products share the same headless-first chassis (typed
`CommandBus` from `@officeai/core`, `OoxmlContainer` for byte-
preserving I/O, `DocumentAgent` interface, MCP transport, CLI
shell, Next.js host), but the file formats themselves push against
different architectural axes. The XLSX deltas doc lists what _that_
product did differently; the list below is what DOCX did
_differently_ and _why_.

---

## 1. Renderer: ProseMirror with a bidirectional bridge

|                       | DOCX                                                                            | XLSX                                                                |
| --------------------- | ------------------------------------------------------------------------------- | ------------------------------------------------------------------- |
| Renderer              | `prosemirror-view` + `prosemirror-state`                                        | hand-rolled virtualized grid in `apps/web/app/xlsx-editor/Grid.tsx` |
| Document-model bridge | bidirectional (`renderer/doc-to-pm.ts` + `renderer/transaction-to-commands.ts`) | unidirectional (model → grid; mutations go through the bus)         |
| Selection             | PM `Selection` (positions in a tree)                                            | `{ anchor: {r,c}; focus: {r,c} }` (single rectangle)                |
| Input handling        | PM transactions intercepted, translated, dispatched                             | DOM event handlers dispatch typed commands directly                 |

**Why**: rich text is a flowing, recursively-nested tree with marks,
inline atoms (images, fields, comment markers), block-level
properties (paragraph styles, alignment, indentation) and
collaborative-text-flow concerns (split / merge / wrap). PM's data
model is purpose-built for exactly this shape; bringing the
hand-rolled grid pattern across would mean re-implementing PM's
position algebra, mark coalescing, and DOM ↔ doc mapping for no
gain.

**Cost**: a real bidirectional bridge — `docToPM` projects
`DocxSnapshot → PM doc` at every external mutation, and
`transactionToCommands` walks the PM transaction's steps to emit
typed `docx:*` commands. Both halves have to agree on positional
semantics for the round trip to converge.

**Benefit**: PM gives us native typing feel, undo/redo, IME, native
selection, marks, decorations, plugins, keymaps and all the editor
ergonomics we'd otherwise build by hand.

## 2. The "two-clock" mirroring loop in `mountDocxEditor`

XLSX has one clock — a DOM event becomes a `xlsx:*` command, the bus
applies it, the grid re-renders from the resulting snapshot. DOCX
has two clocks that have to stay in sync without blocking the user:

```
user keystroke
   │
   ▼
PM dispatchTransaction
   ├── view.updateState(tx)             ← clock A (synchronous, optimistic)
   │     • DOM + selection update now; typing feels native
   │
   └── transactionToCommands(tx)        ← clock B (asynchronous, authoritative)
         • translate steps to docx:*
         • pendingFunnelCount += N
         • void agent.applyCommands(...) (fire-and-forget)
                                  │
                                  ▼
                         agent.subscribe(...)
                          ├── pendingFunnelCount > 0 → consume, return  (self-echo)
                          └── otherwise → re-project snapshot, clamp old selection
```

The contract is **"every mutation flows through the bus, but typing
never blocks on it"**. `pendingFunnelCount` is a counter (not a
boolean — handlers run synchronously inside `applyCommands`, so
`subscribe` fires once per command) that suppresses the bus's echo
of changes the funnel just dispatched. External mutations (agent
prompt, comment dispatched outside the funnel, MCP tool call) hit
the `subscribe` callback with `pendingFunnelCount === 0` and
trigger a full re-projection: `docToPM(snapshot)` rebuilds the doc,
the previous `from`/`to` selection is clamped through the new
content size and walked to the nearest valid textblock.

**Known caveat**: an exception inside `applyCommands` decrements
`pendingFunnelCount` back down so the suppression count cannot
leak. We test this end-to-end in `renderer/renderer.test.ts`.

## 3. Style cascade resolver vs flat style table

XLSX styles are a single workbook-level table with content-hash
deduplication — every cell points at one entry that is the absolute
format. DOCX has a real inheritance graph:

```
docDefaults.rPrDefault
   │
   └── styleId chain (basedOn → basedOn → …)
         │
         └── paragraph.pPr.rPr
               │
               └── run.properties (rPr)
```

`agent/style-resolver.ts` walks this chain, merging right-to-left,
with cycle-safety (depth-cap of 16 + visited set) and returns the
_absolute_ `RunProperties` / `ParagraphProperties` that the toolbar
should display.

**Why this matters in the UI**: the toolbar reads
`resolveEffectiveRpr(snapshot, paragraphIndex, runIndex)`, not the
run's own `rPr`. That is why the bold / italic / size / colour /
alignment / indent buttons correctly show "Heading 1" formatting on
a run that has no inline run properties — the value is inherited
through the style chain.

**Implication**: every style-aware command (`set-paragraph-style`,
`format-range`, `set-paragraph-alignment`, `set-paragraph-indent`,
`set-paragraph-spacing`) writes the most-specific level it can,
relying on the cascade to render the rest. We never copy the
resolved value down into the run.

### 3a. Theme-aware font resolution (P3.9)

`<w:rFonts>` is the one cascade slot where Word does **not** do
per-attribute leaf-wins merging. When a child level supplies an
`<w:rFonts>` element of any shape — a literal `w:ascii="Calibri"`
or a theme ref `w:asciiTheme="majorHAnsi"` — the parent's entire
`<w:rFonts>` is discarded. The resolver mirrors this by treating
the trio `(fontFamily, fontFamilyAsciiTheme, fontFamilyHAnsiTheme)`
as a single property: if the child sets any of them, all three
parent values are dropped. See `style-resolver.ts → mergeRpr`.

After the cascade collapses, theme refs are projected into a
literal typeface so the toolbar can display "Aptos Display" rather
than the opaque `majorHAnsi`. Resolution order:

1. **Literal wins**: if the merged rPr has `fontFamily` set, return
   it as-is. (A run-level `<w:rFonts w:ascii="Comic Sans"/>` always
   beats an inherited theme ref.)
2. **`word/theme/theme1.xml`**: parsed by `parser/theme.ts` into a
   typed `ThemePart` carrying `majorFont.latin` / `minorFont.latin`.
   `majorHAnsi` / `majorAscii` / `majorBidi` map to `majorFont.latin`,
   `minor*` to `minorFont.latin`. East-Asian variants project to
   `*.ea` with a Latin fallback.
3. **Word-default fallback**: when the package ships no theme part
   (synthetic fixtures, hand-written demo docs), `WORD_DEFAULT_THEME_FONTS`
   in `style-resolver.ts` returns the Word 2024+ defaults
   ("Aptos Display" / "Aptos") so the editor still agrees with what
   Word would render on export.

The byte-level round-trip is unaffected: the original `<w:rFonts>`
element is captured via `opaqueProps` at parse time and the
serializer re-emits it verbatim unless an explicit `fontFamily`
mutation overrides it, in which case a fresh `<w:rFonts w:ascii=…
w:hAnsi=…/>` replaces both the literal and theme attributes (which
matches Word's authoring behavior).

## 4. Opaque-blob preservation with a typed display classifier

Both products preserve unknown OOXML byte-identically through
opaque containers — `OpaquePart` for whole zip entries,
`OpaqueXml` / `OpaqueBlock` / `OpaqueInline` for inline subtrees the
parser elected not to type. What is DOCX-specific is what happens
_at render time_. A naive renderer that surfaces every preserved
inline tag as `[opaque]` clutters real-world documents to the point
of being unusable: bookmarks, proof markers, field characters,
paragraph-level revision IDs, content-control wrappers and `<mc:…>`
fallback markers are everywhere.

`model/opaque-classification.ts` triages every preserved tag:

| Display class           | Carrier examples                                                                                                                                           | Render behaviour                                         |
| ----------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------- | -------------------------------------------------------- |
| `metadata`              | `w:bookmarkStart`, `w:bookmarkEnd`, `w:proofErr`, `w:lastRenderedPageBreak`, `w:fldChar`, `w:instrText`, `w:permStart`, `w:annotationRef`, `w:rsidRoot`, … | emit nothing (invisible markup)                          |
| `content-wrapper`       | `w:sdt`, `w:sdtContent`, `w:fldSimple`, `mc:AlternateContent`, `mc:Choice`, `mc:Fallback`, `w:smartTag`, `w:customXml`                                     | flatten the carrier's inner text in place of the wrapper |
| `placeholder` (default) | anything we have not classified                                                                                                                            | legacy `[<tag>]` chip                                    |

XLSX has no comparable layer — its opaque parts (charts, drawings,
VBA, custom XML, theme) live _outside_ the editing surface, so they
have no display question. DOCX needs the classifier because opaque
fragments are interleaved with editable content in the body.

**Known caveat (P3, deferred to P4)**: editing _inside_ a
content-wrapper (e.g. typing into the body of a `<w:sdt>`) is
intentionally blocked. The wrapper is round-tripped verbatim;
mutations route through the parent paragraph. The
`subtreeDirty` plumbing for typed mutations on opaque carriers is
listed as a P4 follow-up in `session-summary.md`.

## 5. Comments lifecycle: one mutation, four parts

XLSX comments anchor to a single cell and live in one part
(`xl/commentsN.xml`). DOCX comments are a four-part choreography
that has to land atomically:

1. Mint `word/comments.xml` (or append to it) with the comment
   body, author, date.
2. Register a `<Override>` for `word/comments.xml` in
   `[Content_Types].xml`.
3. Add a `comments` relationship in
   `word/_rels/document.xml.rels`.
4. Insert three inline marks into the target paragraph in the
   right order: `<w:commentRangeStart w:id="N"/>` …
   `<w:commentReference w:id="N"/>` … `<w:commentRangeEnd w:id="N"/>`.

`commands/add-comment.ts` does all four in one handler. The
W15-extended companion (`word/commentsExtended.xml`, carrying
thread + resolved metadata) is its own dirty flag
(`DocxDirtyFlags.commentsExtended`) so resolving a comment can
re-emit only that part without touching `comments.xml`.

This was flagged as harder-than-expected in the P0 session notes
and is the canonical example of why DOCX needed a _multi-part_
dirty taxonomy from day one.

## 6. Multi-part dirty taxonomy

Both products use part-level dirty flags so the serializer can
re-emit only what the agent touched. DOCX's surface area is wider
because the format is more federated:

```ts
interface DocxDirtyFlags {
  body: boolean; // word/document.xml
  comments: boolean; // word/comments.xml
  commentsExtended: boolean; // word/commentsExtended.xml
  rels: boolean; // word/_rels/document.xml.rels
  contentTypes: boolean; // [Content_Types].xml
  headersAndFooters: ReadonlySet<string>; // word/header*.xml, word/footer*.xml (per path)
  media: ReadonlySet<string>; // word/media/image*.png|jpg|… (per path)
  // styles, numbering, theme tracked alongside on the document model itself
}
```

A freshly-parsed snapshot has every flag clear and every set
empty — that is the load-bearing invariant behind the
byte-equality oracle. The serializer reads the flags once per
save:

- Untouched parts are re-emitted from the cached `Uint8Array`
  inside `OoxmlContainer`. Byte-identical, no re-serialization
  cost.
- Touched parts go through the typed serializer, which emits
  schema-valid OOXML against the ECMA-376 XSDs (gated in CI;
  see §13).

## 7. Tracked changes as a typed inline node

DOCX has first-class support for tracked changes; XLSX has none.
The model represents revisions as a `RevisionWrapper` inline node:

```ts
interface RevisionWrapper {
  readonly kind: "revision";
  readonly id: NodeId;
  readonly revisionType: "ins" | "del";
  readonly author: string;
  readonly date: string;
  readonly children: ReadonlyArray<RunChild>;
  // …
}
```

`<w:delText>` is preserved as a `Run` text leaf with `isDelText: true`
so the serializer can emit it back into `<w:del>` carriers
verbatim. `commands/accept-change.ts` and `commands/reject-change.ts`
are typed handlers that flatten the wrapper (accept) or remove its
contents (reject for `ins`, restore for `del`), with the right
inverse-mutation diff entries so the side-panel UI
(`apps/web/app/editor/TrackedChangesUI.tsx`) can stay in sync.

## 8. Headers and footers as full sub-documents

XLSX print-layout headers/footers are a string property on the
sheet. DOCX header/footer content is a _full body of typed blocks_
living in its own zip entry (`word/headerN.xml` / `word/footerN.xml`),
referenced by relationship from `document.xml.rels`, and
attached to a section through the section's `<w:headerReference>`
and `<w:footerReference>` elements with a `w:type` slot:

| Slot      | When used                                                                 |
| --------- | ------------------------------------------------------------------------- |
| `default` | Every page in the section, fallback for missing slots                     |
| `first`   | Page 1 of the section when `<w:titlePg/>` is set on `<w:sectPr>`          |
| `even`    | Even-numbered pages when `<w:settings><w:evenAndOddHeaders/>` is set (P4) |

`agent/header-footer-graph.ts` resolves the right `HeaderFooterPart`
for each `(sectionIndex, slot)` pair so the page chunker (§9) can
hand the renderer the correct chrome at the correct boundary, and
so the side-panel UI (`apps/web/app/editor/HeaderFooterPanel.tsx`)
can target a specific part with `docx:set-header-text` /
`docx:set-footer-text` / `docx:insert-page-number`.

**Known caveat**: toggling "Different first page" flips the
`titlePg` flag but does not auto-mint the missing
`headerN.xml` part. Listed in `session-summary.md` as a P4
follow-up.

## 9. Page chunker: pure, three-tier signal

`renderer/page-chunker.ts` is a pure function (no DOM access, lives
in `@officeai/docx`) that splits the body into a sequence of
`PageChunk`s. It ranks three break signals:

```
1. Hard breaks    : <w:br w:type="page"/>             (always honoured)
2. Hint breaks    : <w:lastRenderedPageBreak/>        (honoured only when no measure())
3. Measured breaks: measure(blockIndex) cumulative    (honoured when measure is provided)
```

`SectionBreak` blocks are section terminators — the section's
`<w:sectPr>` carries the geometry of the section it ends, every
`SectionBreak` is the last block on its page, and the next block
starts a fresh page implicitly because section geometry can
differ.

**Why pure**: the same chunker drives three callers from a single
source of truth:

- The page-decorations PM plugin in `apps/web/app/lib/page-decorations.ts`
  (browser-side, wires `measure(blockIndex)` to
  `getBoundingClientRect()` on the rendered PM blocks).
- The markdown export's `withPageSections` option in
  `agent/markdown.ts` (Node-side, no `measure`, falls back to hard
  - hint breaks).
- The MCP `docx_get_pages` / `docx_get_page_text` tools in
  `@officeai/agent` (LLM-facing, same fallback).

If the chunker lived in the host app, the LLM would see different
page boundaries than the user. Keeping it in `@officeai/docx`
collapses the discrepancy.

## 10. Page chrome via PM decorations + an `extraPlugins` mount seam

XLSX paints chrome (frozen panes, header bands, drag-resize handles)
inside its own grid component. DOCX paints page chrome (page caps,
page edges, header/footer zones, section markers) inside PM via
`Decoration.node` (per-block styling) and `Decoration.widget`
(chrome that is not part of the document content).

The plugin lives in the host (`apps/web/app/lib/page-decorations.ts`,
not in `@officeai/docx`) because it needs DOM measurement and a
React event bus for double-click → header-edit popovers. The
chunker lives in `@officeai/docx` because it has to be reusable
from Node. The seam between the two is `MountOptions.extraPlugins`:

```ts
mountDocxEditor(target, {
  agent,
  extraPlugins: [
    pageDecorationsPlugin({
      getChunks: () => chunkIntoPages(agent.getSnapshot(), measureBlock),
      onZoneEdit: (detail) => openHeaderFooterPopover(detail),
      // …
    }),
  ],
});
```

**Read-only contract**: any plugin passed via `extraPlugins`
_must not_ dispatch transactions that mutate document content. If
a future plugin needs to mutate, it dispatches a typed command on
the bus; the bus echoes the change back through the
`agent.subscribe` re-projection path (§2). This keeps the "every
mutation flows through the bus" invariant intact even when the
visual layer grows new affordances.

## 11. Field rendering as typed leaves

`<w:fldSimple>` and the `PAGE` / `NUMPAGES` family are promoted out
of opaque-blob land into typed inline leaves:

```ts
interface FieldLeaf {
  readonly kind: "field";
  readonly fieldType: "PAGE" | "NUMPAGES" | "DATE" | …;
  /** Verbatim w:instr attribute, e.g. ` PAGE \\* MERGEFORMAT `. */
  readonly instr: string;
  // …
}
```

The verbatim `w:instr` capture is the reason an untouched field
round-trips byte-identical even when we did not implement its
semantics — the serializer re-emits the literal bytes. The paged
renderer (§9) reads the `fieldType` and substitutes the live page
index for `PAGE` / total page count for `NUMPAGES` at render
time. `commands/insert-page-number.ts` produces the leaf for the
toolbar's "insert page number" affordance.

## 12. Numbering as a separate part with abstract + concrete IDs

Lists are not a runtime concern — they are an OOXML part
(`word/numbering.xml`) with two layers:

- `<w:abstractNum>` — defines list shape (bullets vs decimal,
  level indents, level text format strings).
- `<w:num>` — a concrete instance bound to one `<w:abstractNum>`,
  identified by `numId`.

Paragraphs carry `numbering: { numId, ilvl }` in their typed
properties. The parser builds a typed `NumberingPart` from
`numbering.xml`; the serializer re-emits when the part is dirty.
Toolbar list buttons in `apps/web/app/editor/Toolbar.tsx` set
`numId` / `ilvl` via `docx:set-paragraph-list` /
`docx:remove-paragraph-list`.

**Known caveat**: when a document has no `<w:abstractNum>`,
clicking the bullet / numbered button currently fails loudly via a
toast rather than minting one. The auto-mint path is queued as a
P4 follow-up.

## 13. Markdown export doubles as the LLM ingestion path

XLSX's LLM surface is per-cell / per-range — there is nothing to
project. DOCX needs a lossy-but-readable serialization for LLM
prompts and for the CLI `docx read --format markdown` flag.

`agent/markdown.ts` projects `DocxSnapshot → GFM` with these
mappings:

- Paragraphs whose `styleId` is `Title` / `Heading[1-6]` →
  `#` … `######` headings.
- Paragraphs with `numbering.numId` → numbered (`1.`) or bulleted
  (`-`) list items, indented by `ilvl`.
- Tables → GFM pipe tables (best-effort: first row becomes
  header), falling back to `> [table preserved]` on extraction
  failure.
- `SectionBreak` → `---`.
- Comments → appended `## Comments` section listing thread heads
  with the parent paragraph's plain-text snippet.
- `withPageSections: true` (P3.6 / W22) → segment by page chunk
  with `<!-- page N -->` anchors and `## Page N` headings so an
  LLM can cite "page 3" without having to count.

The MCP tool family (`docx_read`, `docx_get_pages`,
`docx_get_page_text`) calls the same projector, so the LLM sees
exactly what the CLI shows.

## 14. PM coord ↔ `DocxPosition` translation, with an explicit `unsupported` channel

`renderer/transaction-to-commands.ts` walks each PM `Step` and
emits typed commands. It supports `ReplaceStep`, `AddMarkStep`,
`RemoveMarkStep` today. Anything else (notably `ReplaceAroundStep`,
which PM uses for list / blockquote wrap-unwrap) lands in the
`unsupported` channel:

```ts
interface UnsupportedTx {
  reason: string;
  step?: Step;
}
```

The host displays a toast ("structural replace around (lists /
blockquote) deferred"). The PM transaction is _still_ applied to
the local view (clock A), so the user sees the change; it is the
funnel that refuses to translate. This is intentional — silently
accepting a step we cannot represent in `docx:*` would corrupt the
round trip the next time the snapshot re-projects.

XLSX has no analogue because every mutation in the grid is already
a typed command — there is no "unsupported step" failure mode to
surface.

## 15. Real-world fixture corpus + LibreOffice round-trip CI

XLSX ships only `fixtures/xlsx/synthetic/` so far. DOCX also ships
`fixtures/docx/real-world/`, with seven full documents emitted by
different writers:

| Fixture                          | Source family | Stresses                                     |
| -------------------------------- | ------------- | -------------------------------------------- |
| `01-styled-letter.docx`          | Word          | run formatting, paragraph styles             |
| `02-report-headers-footers.docx` | Word          | section breaks, header/footer parts, fields  |
| `03-numbered-list.docx`          | LibreOffice   | numbering.xml, ilvl nesting                  |
| `04-table-grid.docx`             | Word          | typed tables, cell merge, row props          |
| `05-inline-image.docx`           | Word          | media parts, drawing relationships           |
| `06-comments-and-changes.docx`   | Word          | comments lifecycle, tracked changes          |
| `07-toc-sdt.docx`                | Word          | SDT content controls, TOC field, MC fallback |

Two CI gates ride on this corpus:

- **`docx-libreoffice-roundtrip`** (`.github/workflows/ci.yml`) —
  every real-world fixture is parsed, re-emitted by the agent,
  then opened by headless LibreOffice via `soffice --convert-to`.
  A silent corruption (foreign-reader can't open our re-emit) fails
  the job. Catches a class of bugs our own byte-equality oracle
  cannot — when _we_ produced the input bytes, the oracle is
  circular.
- **`docx-perf`** — `make perf-docx` builds and serializes a
  synthetic 100-page DOCX, asserting parse / dispatch /
  serialize budgets documented inline in `scripts/perf-docx.mjs`.

## 16. OOXML XSD validation as a CI gate

`make schema-validate` (job `docx-schema-validation`) runs
`xmllint --schema` on every fixture, on **both** sides — the input
bytes _and_ the agent re-emit — against the ECMA-376 5th-edition
Transitional XSDs fetched and SHA-256-pinned via
`scripts/fetch-ooxml-xsd.mjs`.

**Current state**: the job is `continue-on-error: true`. The
`docx`-npm generator (which we use for synthetic fixtures) emits a
small set of `mc:Ignorable` / `w15:*` extension attributes that
the strict Transitional XSD rejects, and we round-trip those
verbatim. The build log captures the violations and the planned
fixes; the gate is in place so as soon as the day-1 noise
clears, flipping `continue-on-error` to `false` is a one-line
change.

XLSX has no XSD validation gate — its serializer rides through
SheetJS, which has its own validation discipline.

## 17. Test pyramid is shaped differently

| Layer              | DOCX                                                                                                                                                    | XLSX                                                                                       |
| ------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------ |
| Unit               | model / serializer / handlers + style resolver + opaque classification + page chunker + header-footer graph + markdown export + transaction-to-commands | model / serializer / handlers + **formula engine (475 tests)** + style-table dedup         |
| Integration        | byte-equality oracle on synthetic + real-world fixtures + four-part comments lifecycle + tracked-change accept/reject                                   | byte-equality oracle + per-command property tests for the inverse mutation                 |
| System (Node-only) | `make perf-docx` (parse/dispatch/serialize budgets), `make schema-validate` (XSD), `make roundtrip-libre` (foreign reader)                              | none today                                                                                 |
| E2E (Playwright)   | editor smoke + page-sheets + double-click header / footer edit + tracked-changes side panel                                                             | editor smoke + drag interactions (resize, drag-extend selection) + caret-aware formula bar |

**Why DOCX leans heavier on integration / system**: the byte
preservation invariant is the headline product property. A unit
test on the serializer cannot prove that the bytes Word saved
yesterday will still open in Word tomorrow — only round-tripping
through a foreign reader (LibreOffice) on real-world inputs proves
that. **Why XLSX leans heavier on unit**: the formula engine is a
deterministic pure function — its 475 unit tests catch regressions
the same day they ship, and there is no foreign-reader equivalent
because `xlsx`-npm round-trips its own bytes losslessly.

`renderer/transaction-to-commands.test.ts` is the most
DOCX-specific test file in the repo. It pins the PM step → typed
command translation contract for ReplaceStep / AddMarkStep /
RemoveMarkStep, asserts that ReplaceAroundStep surfaces in the
`unsupported` channel rather than producing nonsense commands, and
exercises the position-mapping math that backs §14.

---

## When this doc should be updated

- A subsystem here is replaced (e.g. ProseMirror swapped for
  Lexical / TipTap, or the page chunker rewritten to do real
  measured layout) — bump the relevant section with the new
  substrate and the migration rationale.
- The deferred items in §4 (subtree-dirty editing inside opaque
  carriers), §8 (auto-mint header/footer parts on `titlePg`
  toggle), §12 (auto-mint `<w:abstractNum>` for list buttons) or
  §16 (flip XSD validation from advisory to blocking) are picked
  up — strike the caveat and link to the closing build-log entry.
- A PPTX product lands and shares non-trivial machinery with
  DOCX — the contrast tables grow a PPTX column, or a row
  collapses if PPTX/DOCX converge. The companion
  `architecture-pptx-deltas.md` should be created in parallel.
