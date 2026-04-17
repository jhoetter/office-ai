# XLSX — Serializer

> `XlsxSnapshot` → bytes (`.xlsx`) with **byte-preservation** for every
> OOXML part the editor did not touch. The serializer is the write half
> of the round-trip contract: untouched parts ship through verbatim from
> `OoxmlContainer.parts`; dirty parts are re-emitted from the typed
> model.

This document mirrors the SHAPE of [`spec/docx/serializer.md`](../docx/serializer.md).

References:

- [`spec/shared/ooxml-utils.md`](../shared/ooxml-utils.md) — container
  guarantees, XML helpers.
- [`spec/xlsx/parser.md`](parser.md) — what the parser modeled and what
  it left opaque.
- [`spec/xlsx/feature-scope.md`](feature-scope.md) — what we model.
- [`spec/xlsx/edge-cases.md`](edge-cases.md) — round-trip rules
  (EC-O1, EC-O2, EC-S6, EC-F6, EC-F7).
- [`spec/xlsx/analysis-sheetjs.md`](analysis-sheetjs.md) §10–§11 — what
  we delegate to SheetJS vs own ourselves.
- [`packages/docx/src/serializer/serialize.ts`](../../packages/docx/src/serializer/serialize.ts)
  — the dirty-part discipline pattern we mirror.

---

## 1. Goal

Produce an `ArrayBuffer` containing a structurally-valid `.xlsx` such
that:

1. Every part **not touched** by any command since the last load is
   **byte-identical** at the part-content level (SHA-256 match against
   `snapshot.partHashes`). The zip-archive bytes themselves may differ
   (compression level, central directory ordering — JSZip handles this);
   the contract is at the **part** level, never the **archive** level.
2. Every part **touched** by commands is re-emitted from the typed model
   in a shape Excel / LibreOffice / Google Sheets accept without a
   repair dialog. Touched parts are **structurally equivalent**, not
   byte-equal (per EC-O2).
3. Opaque blobs round-trip verbatim; the model never deserializes them
   on read and never re-serializes them on write.
4. `serialize(parse(buf))` with no intervening mutations is a no-op at
   the model boundary: `partHashes` matches for every entry.

---

## 2. Algorithm

```
serializeXlsx(snapshot): Promise<ArrayBuffer>
  container = snapshot.container.clone()

  // Each block below is a no-op when its dirty flag is false.
  // The cumulative effect is: untouched parts stay byte-identical to
  // the loaded bytes; dirty parts are re-emitted from the typed model.

  serializeWorkbookXml(container, snapshot)
  serializeSharedStrings(container, snapshot)
  serializeStyles(container, snapshot)
  serializeSheets(container, snapshot)         // per-sheet dirty flags
  serializeComments(container, snapshot)       // per-sheet dirty flags
  serializeThreadedComments(container, snapshot)
  serializePersons(container, snapshot)
  serializeRelationships(container, snapshot)  // per-rels-file dirty
  serializeContentTypes(container, snapshot)   // only if Override list changed
  serializeCalcChain(container, snapshot)      // optional; off by default

  return container.serialize()
```

The skeleton is intentionally identical to
[`packages/docx/src/serializer/serialize.ts`](../../packages/docx/src/serializer/serialize.ts):
clone the container, walk dirty flags, write back, hand to
`OoxmlContainer.serialize()` for re-zipping.

---

## 3. Dirty-flag model

Every mutation flags one or more parts. The serializer trusts these
flags absolutely — if a flag says "clean", the original bytes ship
through unchanged.

```typescript
export interface XlsxDirtyFlags {
  readonly workbook: boolean; // xl/workbook.xml
  readonly sst: boolean; // xl/sharedStrings.xml
  readonly styles: boolean; // xl/styles.xml
  readonly contentTypes: boolean; // [Content_Types].xml
  readonly packageRels: boolean; // _rels/.rels (rare; ~never in P0)
  readonly workbookRels: boolean; // xl/_rels/workbook.xml.rels
  readonly calcChain: boolean; // xl/calcChain.xml (we may force-remove)
  /** Per-sheet flags, keyed by sheetId. */
  readonly sheets: ReadonlyMap<string, SheetDirty>;
}

export interface SheetDirty {
  readonly cells: boolean; // xl/worksheets/sheet{N}.xml
  readonly rels: boolean; // xl/worksheets/_rels/sheet{N}.xml.rels
  readonly comments: boolean; // xl/comments{N}.xml
  readonly threadedComments: boolean; // xl/threadedComments/threadedComment{N}.xml
}
```

The flag → command mapping (excerpt, for the P0 commands):

| Command                      | Dirties                                                                                             |
| ---------------------------- | --------------------------------------------------------------------------------------------------- |
| `xlsx:set-cell-value`        | `sheets[id].cells`, possibly `sst`                                                                  |
| `xlsx:set-cell-formula`      | `sheets[id].cells`, possibly `sst`, possibly `calcChain`                                            |
| `xlsx:set-cell-format`       | `sheets[id].cells`, `styles`                                                                        |
| `xlsx:set-range-values`      | `sheets[id].cells`, possibly `sst`                                                                  |
| `xlsx:apply-range-format`    | `sheets[id].cells`, `styles`                                                                        |
| `xlsx:merge-cells`           | `sheets[id].cells`                                                                                  |
| `xlsx:unmerge-cells`         | `sheets[id].cells`                                                                                  |
| `xlsx:insert-rows` / `:cols` | `sheets[id].cells` (every sheet whose formulas reference this band — typically just the same sheet) |
| `xlsx:delete-rows` / `:cols` | `sheets[id].cells`                                                                                  |
| `xlsx:add-sheet`             | `workbook`, `workbookRels`, `contentTypes`, `sheets[new]`                                           |
| `xlsx:rename-sheet`          | `workbook`, `sheets[any with x-sheet refs]`                                                         |
| `xlsx:set-sheet-tab-color`   | `sheets[id].cells` (tab color lives in `<sheetView>`)                                               |
| `xlsx:set-freeze-panes`      | `sheets[id].cells` (freeze lives in `<sheetView>`)                                                  |
| `xlsx:reorder-sheets`        | `workbook` (sheet `<sheets>` order)                                                                 |
| `xlsx:add-comment`           | `sheets[id].comments`, `sheets[id].threadedComments`,                                               |
|                              | `sheets[id].rels`, `contentTypes` (1st time per workbook),                                          |
|                              | `workbookRels` (1st time only — for `xl/persons/person.xml`)                                        |
| `xlsx:reply-comment`         | `sheets[id].threadedComments`                                                                       |
| `xlsx:resolve-comment`       | `sheets[id].threadedComments`                                                                       |
| `xlsx:delete-comment`        | `sheets[id].comments`, `sheets[id].threadedComments`,                                               |
|                              | possibly `sheets[id].rels` + `contentTypes` if last comment                                         |
| `xlsx:add-hyperlink`         | `sheets[id].cells`, `sheets[id].rels`                                                               |
| `xlsx:remove-hyperlink`      | `sheets[id].cells`, `sheets[id].rels`                                                               |

Default for a freshly-parsed snapshot: every flag is `false`. So
`serialize(parse(buf))` writes back the original bytes for every part.

---

## 4. Workbook (`xl/workbook.xml`)

```
serializeWorkbookXml(container, snapshot):
  if !snapshot.dirty.workbook: return
  tree = buildWorkbookTree(snapshot.root)
  xml = serializeXml(tree, { xmlDeclaration: XML_DECL })
  container.writeText("xl/workbook.xml", xml)

buildWorkbookTree(workbook):
  return [
    { "?xml": [], ":@": { "@_version": "1.0", "@_encoding": "UTF-8", "@_standalone": "yes" } },
    {
      "workbook": [
        ...optional("fileVersion"),
        buildBookViews(workbook.workbookView),
        { "sheets": workbook.sheets.map(s => ({
            "sheet": [],
            ":@": {
              "@_name": s.name,
              "@_sheetId": String(s.id),
              "@_r:id": s.rId,
              ...(s.visibility !== "visible" ? { "@_state": s.visibility } : {}),
            },
          })) },
        ...(workbook.definedNames.length > 0
          ? [{ "definedNames": workbook.definedNames.map(buildDefinedName) }]
          : []),
        // Every other workbook-level child (calcPr, customWorkbookViews,
        // pivotCaches, smartTagPr, webPublishing, fileRecoveryPr, extLst)
        // is preserved as opaque on the workbook model and re-emitted here
        // verbatim. Pivot caches in particular MUST round-trip — losing
        // them would break the pivot tables we treat as opaque (EC-O3).
        ...workbook.workbookOpaqueChildren,
      ],
      ":@": workbook.workbookRootAttrs,
    },
  ]
```

The workbook root attrs (every `xmlns:*` declaration including
`xmlns:r`, `xmlns:mc`, `xmlns:x15`, `mc:Ignorable`, …) are preserved
verbatim from parse.

We re-emit `xl/workbook.xml` only when the typed fields actually change
(sheet add / rename / reorder / visibility / defined-name edit). The
opaque children (`pivotCaches`, `calcPr`, `extLst`, …) are **always**
included; they just sit in the in-memory model untouched between parse
and serialize, so the rebuilt XML's content for those nodes is
byte-equivalent (modulo attribute order — see
[`docx/serializer.md`](../docx/serializer.md) §"Whitespace and the
byte-stability caveat" for the same caveat).

---

## 5. Sheets (`xl/worksheets/sheet{N}.xml`)

This is the single most important file the serializer writes. Strategy:
**SheetJS owns cell-level XML, we layer everything else on top.**

```
serializeSheets(container, snapshot):
  for sheet in snapshot.root.sheets:
    const sd = snapshot.dirty.sheets.get(sheet.id)
    if !sd || !sd.cells: continue
    const path = `xl/worksheets/sheet${sheet.indexInWorkbook + 1}.xml`
    const sjsXml = renderSheetThroughSheetJS(sheet, snapshot)
    const finalXml = injectNativeAnnotations(sjsXml, sheet, snapshot)
    container.writeText(path, finalXml)
```

### 5.1 SheetJS pass

We build a **single-sheet** SheetJS workbook from our typed model, ask
SheetJS to render it, and harvest just `xl/worksheets/sheet1.xml` from
its output. We use SheetJS for sheet XML because it gets a long list of
fiddly things right that we don't want to reinvent:

- Correct cell type tags (`t="b" | "n" | "s" | "str" | "e" | …`).
- Correct value encoding (booleans as `0`/`1`; errors via numeric code).
- Shared-string index references (`<v>N</v>` where `N` is the SST index),
  using `bookSST: true`.
- Number-format application (the SSF library decides display strings).
- Address arithmetic via `XLSX.utils.encode_cell`.
- `dimension` element computation.
- `<row>` blocks with the right `r=` attributes, `spans=`, `ht=`,
  `customHeight`, `hidden`.

```
renderSheetThroughSheetJS(sheet, snapshot) -> string:
  const sjsWb = XLSX.utils.book_new()
  const sjsWs = ourSheetToSheetJS(sheet, snapshot.root.styles, snapshot.root.sst)
  XLSX.utils.book_append_sheet(sjsWb, sjsWs, sheet.name)
  const out = XLSX.writeXLSX(sjsWb, {
    type: "array",
    bookSST: true,        // emit a real SST; preserves rich text
    compression: true,
    cellDates: false,
    WTF: false,
  })
  // out is a complete xlsx zip with one sheet. We only want sheet1.xml.
  const innerContainer = await OoxmlContainer.load(out)
  return innerContainer.readText("xl/worksheets/sheet1.xml")
```

The `out`-then-load round-trip is acceptable because:

1. The single-sheet workbook is tiny.
2. SheetJS's writer is the only correct implementation we have for
   row-block layout, `spans=`, `dimension=`, etc.
3. We have the exact same pattern in DOCX (where some block types
   serialize via `fast-xml-parser` and we extract just the relevant
   subtree).

Style handling: the SheetJS pass needs an xfId per cell. We map our
content-hashed style id back to the cell's original xfId via
`Styles.xfIdForId(...)` (a roundtripped index that the parser populated
from `cellXfs`). New styles introduced by `xlsx:set-cell-format` get a
fresh xfId allocated when we write `xl/styles.xml` (§7) and that id flows
back to the SheetJS pass before this function runs.

### 5.2 Native annotations layer

SheetJS does not emit `<conditionalFormatting>`, `<dataValidations>`,
`<hyperlinks>` resolution against the rels graph, or `<sheetView>` state
beyond zoom + RTL (analysis-sheetjs §1.3). We layer these in by parsing
the SheetJS output back through `parseXml`, splicing our native blocks
into the right positions per the OOXML schema, and re-serializing.

```
injectNativeAnnotations(xml, sheet, snapshot) -> string:
  tree = parseXml(xml)
  worksheetEl = findChild(tree, "worksheet")

  upsertChild(worksheetEl, "sheetPr",      buildSheetPr(sheet))         // tab color
  upsertChild(worksheetEl, "sheetViews",   buildSheetViews(sheet))       // freeze + selection
  upsertChild(worksheetEl, "mergeCells",   buildMergeCells(sheet))       // SheetJS does emit these but we override with the typed model
  upsertChild(worksheetEl, "conditionalFormatting", buildAllCF(sheet))   // multiple blocks
  upsertChild(worksheetEl, "dataValidations", buildDataValidations(sheet))
  upsertChild(worksheetEl, "hyperlinks",   buildHyperlinks(sheet, snapshot))
  upsertChild(worksheetEl, "sheetProtection", buildSheetProtection(sheet))

  return serializeXml(tree, { xmlDeclaration: XML_DECL })
```

OOXML's `CT_Worksheet` schema requires children in a specific order:

```
sheetPr → dimension → sheetViews → sheetFormatPr → cols → sheetData
→ sheetCalcPr → sheetProtection → protectedRanges → scenarios →
autoFilter → sortState → dataConsolidate → customSheetViews →
mergeCells → phoneticPr → conditionalFormatting → dataValidations →
hyperlinks → printOptions → pageMargins → pageSetup → headerFooter →
rowBreaks → colBreaks → customProperties → cellWatches → ignoredErrors
→ smartTags → drawing → legacyDrawing → legacyDrawingHF → drawingHF →
picture → oleObjects → controls → webPublishItems → tableParts → extLst
```

`upsertChild` finds (or creates) the named child and places it at the
right position relative to its siblings according to this order. The
ordering helper is shared utility code in
`packages/xlsx/src/serializer/sheet-schema-order.ts`.

### 5.3 Per-feature serialization strategy

| Feature                     | Strategy                                                                                                                                                       |
| --------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| Cell values                 | SheetJS pass                                                                                                                                                   |
| Cell formulas               | SheetJS pass; we re-share contiguous identical-shape groups _(P1; in P0 we always emit per-cell formulas — round-trips through Excel which re-shares on save)_ |
| Cell number format          | SheetJS pass for cell.z; `xl/styles.xml` is rebuilt by §7                                                                                                      |
| Cell font/fill/border/align | xfId on the cell points at our rebuilt `cellXfs[]`                                                                                                             |
| Cell hyperlink              | We emit `<hyperlinks>` and write the rels in §10                                                                                                               |
| Cell comment                | §6                                                                                                                                                             |
| Range values                | Decomposed to per-cell writes upstream of the serializer                                                                                                       |
| Apply range format          | Same — per-cell xfId edits that flow into SheetJS pass + styles                                                                                                |
| Merge / unmerge             | Native: `<mergeCells>` block from `sheet.merges` (IndexedRanges)                                                                                               |
| Insert/delete row+col       | The mutation handlers update `cells`, `merges`, `conditionalFormats`, `hyperlinks`, formula refs. The serializer just emits the post-mutation state.           |
| Add sheet                   | New cells/comments/rels parts written; workbook + content-types updated                                                                                        |
| Rename sheet                | `<sheet name=…>` updated in workbook; cross-sheet formula refs already rewritten by mutation handler                                                           |
| Set tab color               | `<sheetPr><tabColor/></sheetPr>` in sheet XML                                                                                                                  |
| Set freeze panes            | `<sheetViews><sheetView><pane …/></sheetView></sheetViews>`                                                                                                    |
| Reorder sheets              | `<sheets>` order in workbook XML                                                                                                                               |
| Filters: apply/clear        | `<autoFilter ref=…>` block; existing filterColumn children preserved if range unchanged                                                                        |
| Sort                        | Decomposed to row-shuffle upstream; serializer just emits                                                                                                      |
| Conditional formatting      | Native: `<conditionalFormatting>` per sqref group; `dxfs` updated in `xl/styles.xml` §7                                                                        |
| Defined names               | `<definedNames>` in workbook XML (read/preserve in P0)                                                                                                         |
| Multi-sheet workbook        | Each sheet serialized independently; workbook stitches them                                                                                                    |
| Charts                      | Opaque (drawings, chart parts shipped through verbatim)                                                                                                        |

### 5.4 What we never emit unless dirty

For each sheet, if `dirty.sheets.get(sheet.id)?.cells !== true`, the
sheet's `xl/worksheets/sheet{N}.xml` ships through verbatim. This is the
heart of the byte-preservation contract.

---

## 6. Comments (`xl/comments{N}.xml` + `xl/threadedComments/threadedComment{N}.xml`)

Comments are split across two parts (legacy + threaded — Excel writes
both for back-compat with Excel < 2017). We serialize both whenever
**either** is dirty, because adding a comment dirties both, resolving a
comment dirties only the threaded part, and the two are tightly coupled
(same pattern as DOCX's comments + commentsExtended split — see
[`packages/docx/src/serializer/serialize.ts`](../../packages/docx/src/serializer/serialize.ts)
lines 70-83).

```
serializeComments(container, snapshot):
  for sheet in snapshot.root.sheets:
    const sd = snapshot.dirty.sheets.get(sheet.id)
    if !sd?.comments: continue
    const path = `xl/comments${sheet.indexInWorkbook + 1}.xml`
    if sheet.comments.length === 0:
      container.removePart(path)
      removeRel(container, sheet, type: ".../comments")
      removeContentTypeOverride(container, "/" + path)
      continue
    ensureCommentsPart(container, sheet, path)
    container.writeText(path, buildLegacyCommentsXml(sheet))

serializeThreadedComments(container, snapshot):
  for sheet in snapshot.root.sheets:
    const sd = snapshot.dirty.sheets.get(sheet.id)
    if !sd?.comments && !sd?.threadedComments: continue
    const records = collectThreadedRecords(sheet)
    const path = `xl/threadedComments/threadedComment${sheet.indexInWorkbook + 1}.xml`
    if records.length === 0:
      container.removePart(path)
      removeRel(container, sheet, type: ".../threadedComment")
      removeContentTypeOverride(container, "/" + path)
      continue
    ensureThreadedCommentsPart(container, sheet, path)
    container.writeText(path, buildThreadedCommentsXml(records))
```

`xl/persons/person.xml` is the third coupled part — it carries the
display-name → personId map referenced by every threaded comment. We
write it whenever a new author is introduced; otherwise it stays clean.

---

## 7. Styles (`xl/styles.xml`)

We **do not** delegate to SheetJS for styles. Per analysis-sheetjs §4.1,
SheetJS rebuilds the styles table to a near-fixed scaffold that drops
named styles, dxfs, full borders, gradient fills, and most fonts. We
own this part end-to-end.

```
serializeStyles(container, snapshot):
  if !snapshot.dirty.styles: return
  const styles = snapshot.root.styles
  const tree = [
    XML_DECL_ENTRY,
    {
      "styleSheet": [
        buildNumFmts(styles.numFmts),
        buildFonts(styles.fonts),
        buildFills(styles.fills),
        buildBorders(styles.borders),
        buildCellStyleXfs(styles.cellStyleXfs),
        buildCellXfs(styles.cellXfs),
        buildCellStyles(styles.cellStyles),
        buildDxfs(styles.dxfs),
        buildTableStyles(styles.tableStyles),
        ...(styles.colors ? [buildColors(styles.colors)] : []),
        ...(styles.extLst ? [opaqueToEntry(styles.extLst)] : []),
      ],
      ":@": styles.rootAttrs,
    },
  ]
  container.writeText("xl/styles.xml", serializeXml(tree, { xmlDeclaration: XML_DECL }))
```

The styles model is **append-only** in P0: existing xfIds keep their
indices through round-trip; new xfIds (introduced by
`xlsx:set-cell-format`) get appended. This is what lets the rest of the
sheet XML keep referring to the same xfId numbers between parse and
serialize without having to re-walk every cell.

When the styles part is regenerated we also bump `dirty.styles` so the
next serialize will include it. The mutation handlers responsible for
allocating new xfIds set this flag.

---

## 8. Shared strings (`xl/sharedStrings.xml`)

```
serializeSharedStrings(container, snapshot):
  if !snapshot.dirty.sst: return
  const sst = snapshot.root.sst
  const tree = [
    XML_DECL_ENTRY,
    {
      "sst": sst.entries.map(buildSharedStringItem),
      ":@": {
        "@_xmlns": SPREADSHEETML_2006,
        "@_count": String(sst.count),
        "@_uniqueCount": String(sst.entries.length),
        ...sst.rootAttrs,                       // any non-default xmlns:*
      },
    },
  ]
  container.writeText("xl/sharedStrings.xml", serializeXml(tree, { xmlDeclaration: XML_DECL }))

buildSharedStringItem(entry):
  if entry.runs:
    return { "si": entry.runs.map(buildRichRun) }
  if entry.rawXml:
    // Preserved verbatim for entries we did not author.
    return parseXml(entry.rawXml)[0]
  return { "si": [{ "t": [{ "#text": entry.text }] }] }
```

**Order-of-first-occurrence is preserved.** When a mutation introduces a
new string we **append** to `sst.entries`; we never reorder. This keeps
`<v>N</v>` references in untouched sheet XML still valid.

When SST is dirty we rebuild. We do not regenerate the SST from scratch
unless every sheet has also been re-emitted — instead we produce a
strictly extended SST whose first M entries match the original. The
extension entries are computed by walking every dirty sheet's cells and
collecting strings not already in the SST.

For EC-S6 (inline string vs SST): we prefer SST when the same string
appears more than once in the workbook; inline (`<is><t>…</t></is>`)
otherwise. The SheetJS pass in §5.1 emits `t="s"` cells with `<v>N</v>`
references; we pick the SST/inline split before invoking it by
pre-populating the workbook's `wb.Strings` table.

---

## 9. Content types (`[Content_Types].xml`)

```
serializeContentTypes(container, snapshot):
  if !snapshot.dirty.contentTypes: return
  const ct = ContentTypes.load(container)
  // Defaults for media extensions (jpg, png, gif) added/removed
  // as media parts come and go. P0 has no media additions for XLSX.
  // Overrides for added/removed parts:
  for path, contentType of snapshot.contentTypeAdditions:
    if !ct.hasOverride(path): ct.addOverride(path, contentType)
  for path of snapshot.contentTypeRemovals:
    ct.removeOverride(path)
  ct.writeBack(container)
```

We **only** edit `[Content_Types].xml` when a part has been added or
removed. The common case (cell edits, comment additions to a sheet that
already has a comments part, format changes) does not dirty content
types.

The full list of XLSX content types we know how to add/remove:

| Operation                         | Content type                                                                |
| --------------------------------- | --------------------------------------------------------------------------- |
| First comment on workbook         | `application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml`  |
| First threaded comment            | `application/vnd.ms-excel.threadedcomments+xml`                             |
| First threaded comment            | `application/vnd.ms-excel.person+xml`                                       |
| Add sheet                         | `application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml` |
| Remove last comment from workbook | (remove the comments override)                                              |

Any other content-type edit requires a future work item; in P0 every
mutation handler sets `dirty.contentTypes` only when its operation falls
into this table.

---

## 10. Relationship graphs (`*.rels`)

```
serializeRelationships(container, snapshot):
  if snapshot.dirty.workbookRels:
    rels = RelationshipGraph.loadFor(container, "xl/workbook.xml")
    rebuildWorkbookRels(rels, snapshot)
    rels.writeBack(container)

  for sheet in snapshot.root.sheets:
    const sd = snapshot.dirty.sheets.get(sheet.id)
    if !sd?.rels: continue
    const sheetPath = `xl/worksheets/sheet${sheet.indexInWorkbook + 1}.xml`
    rels = RelationshipGraph.loadFor(container, sheetPath)
    rebuildSheetRels(rels, sheet)             // hyperlink targets, comments, drawings
    rels.writeBack(container)
```

We **only** re-emit `*.rels` files when their relationship list has
actually changed: a new hyperlink target, a new comments part, a new
drawing reference. Re-numbering of `rId`s is avoided where possible —
existing rels keep their original ids; new ones are minted via
`RelationshipGraph.mintId()`.

The package-level `_rels/.rels` is **never** edited in P0 (it points at
docProps + workbook, all of which are stable parts).

---

## 11. CalcChain (`xl/calcChain.xml`)

`xl/calcChain.xml` is Excel's hint for recalculation order. Excel
rebuilds it on first open if missing — so the safe move is to **delete
it** whenever a formula has changed and let Excel rebuild. We do this
because:

1. Maintaining a correct calcChain across all our mutations is
   error-prone.
2. The cost (a few ms of recompute on first open in Excel) is invisible
   to the user.
3. Every other tool (LibreOffice, Google Sheets) ignores it.

```
serializeCalcChain(container, snapshot):
  if !snapshot.dirty.calcChain: return
  if container.has("xl/calcChain.xml"):
    container.removePart("xl/calcChain.xml")
    removeContentTypeOverride(container, "/xl/calcChain.xml")
    removeRel(container, "xl/_rels/workbook.xml.rels", type: ".../calcChain")
```

Mutation handlers that touch formulas (`xlsx:set-cell-formula`,
`xlsx:set-cell-value` when it overwrites a formula, `xlsx:insert-rows`,
`xlsx:delete-rows`, `xlsx:rename-sheet`) all set `dirty.calcChain`.

---

## 12. What we never re-emit unless dirty

The following parts are **always** preserved as-is from
`OoxmlContainer.parts`. They are never re-serialized in P0; the
serializer ignores them entirely.

- `xl/theme/theme{N}.xml` (any N) — Custom corporate themes round-trip
  bit-identical. SheetJS would overwrite `theme1.xml` with a default; we
  bypass that path entirely.
- `xl/drawings/*` — drawing definitions for charts and images.
- `xl/charts/*` — chart parts (definitions, colors, styles).
- `xl/media/*` — embedded images (PNG, JPEG, EMF, WMF).
- `xl/embeddings/*` — OLE objects.
- `xl/pivotTables/*`, `xl/pivotCache/*` — pivot tables and their caches.
  Per EC-O3 we don't invalidate the cache when the agent edits a
  referenced range; Excel refreshes on open.
- `xl/slicers/*`, `xl/slicerCaches/*` — slicers.
- `xl/timelines/*`, `xl/timelineCaches/*` — timeline filters.
- `xl/queryTables/*`, `xl/connections.xml`, `xl/model/*` — Power Query /
  Power Pivot.
- `xl/customXml/*`, `customXml/*`, `xl/_rels/*.rels` for customXml —
  custom XML data binding (EC-O4).
- `xl/externalLinks/*`, `xl/externalLinks/_rels/*` — external workbook
  references (EC-O5).
- `xl/vbaProject.bin`, `xl/vbaProjectSignature.bin` — VBA macros and the
  signature on the macro project. Per `feature-scope.md` we preserve
  but never execute.
- `xl/printerSettings/*`, `xl/activeX/*`, `xl/ctrlProps/*` — printer
  settings, ActiveX, control properties.
- `xl/tables/*` — Excel Tables (ListObjects). Documented as a known
  fidelity loss in P0 (analysis-sheetjs §8.2): the table block in
  `<tableParts>` of the host sheet round-trips because we preserve the
  full sheet XML when not dirty, and the table parts themselves stay
  opaque. **A mutation that touches a sheet hosting a table is rejected
  in P0** with `precheck → unsupported(table-host-sheet)` until we ship
  table-aware editing.
- `docProps/thumbnail.*` — thumbnail images.
- Any `xl/*` part with an unrecognised content type.

The list mirrors the opaque-blob set in the parser's Stage C table
(see [`parser.md`](parser.md) §2 Stage C). Single source of truth: the
mapping table in [`ooxml-mapping.md`](ooxml-mapping.md) (companion doc).

---

## 13. Errors

```typescript
export type XlsxSerializeErrorKind =
  | "sheet-failed"
  | "comments-failed"
  | "threaded-comments-failed"
  | "styles-failed"
  | "shared-strings-failed"
  | "workbook-failed"
  | "rels-failed"
  | "content-types-failed"
  | "unknown-block"
  | "unknown-cell-kind"
  | "missing-original-bytes"
  | "model-invariant-violation";

export class XlsxSerializeError extends Error {
  readonly name = "XlsxSerializeError";
  readonly kind: XlsxSerializeErrorKind;
  readonly partPath?: string;
  readonly sheetId?: string;
  readonly cause?: unknown;
  constructor(kind: XlsxSerializeErrorKind, message: string,
              opts?: { partPath?: string; sheetId?: string; cause?: unknown }) { ... }
}
```

Failure modes:

| Trigger                                                             | Error                                     |
| ------------------------------------------------------------------- | ----------------------------------------- |
| SheetJS write throws                                                | `sheet-failed`, sheetId, cause = err      |
| Native annotation injection produces invalid XML                    | `sheet-failed`, sheetId                   |
| `parseXml` fails on a part we just emitted (sanity-check assertion) | `sheet-failed` / `comments-failed` / etc. |
| A dirty sheet's `cells` is empty but `cells.size > 0` mismatch      | `model-invariant-violation`               |
| `RelationshipGraph.writeBack` throws                                | `rels-failed`, partPath                   |
| `ContentTypes.writeBack` throws                                     | `content-types-failed`                    |
| Missing original bytes for a clean part (should never happen)       | `missing-original-bytes`, partPath        |
| Unknown cell-value kind in switch                                   | `unknown-cell-kind`                       |
| Unknown block kind in switch                                        | `unknown-block`                           |

The same exhaustive-switch discipline DOCX uses — every switch on a
discriminated union throws `model-invariant-violation` in the default
arm so adding a new kind is a compile-time + runtime error.

---

## 14. Test invariants

Tests live in `packages/xlsx/src/serializer/*.test.ts` and the cross-cutting
roundtrip suite at `packages/xlsx/src/parser/parse-roundtrip.test.ts`.

The invariant test (driven by
`scripts/run-xlsx-roundtrip.mjs`):

```
for each fixture in fixtures/xlsx/synthetic/*.xlsx:
  inputBuf = readFile(fixture)
  snapshot = await parseXlsx(inputBuf)
  outputBuf = await serializeXlsx(snapshot)
  output = await OoxmlContainer.load(outputBuf)
  for path in snapshot.container.parts.keys():
    const original = snapshot.partHashes[path]
    const after = sha256Hex(output.readBytes(path))
    assert original === after,
      `Untouched part ${path} changed: ${original} → ${after}`
```

This is the load-bearing invariant. It catches:

- Any code path that accidentally re-emits an opaque part.
- Forgotten dirty-flag updates that cause a real edit to slip past the
  serializer (the test doesn't apply mutations, so anything that comes
  out different is a serializer bug).
- Encoding drift (BOMs added/removed, line endings normalized).
- Whitespace normalization (we should never re-serialize untouched
  parts; if the original had two-space indent in `xl/styles.xml`, the
  output must too).

Additionally, for **modified** sheets we run a structural-equivalence
test: parse the output, parse the input, compare the typed model
field-by-field. Per EC-O2 we accept attribute reordering and minor
whitespace differences in the modified parts; the assertion is "model
equality after reparse", not "byte equality".

The fixtures cover:

| Fixture                      | What it stresses                            |
| ---------------------------- | ------------------------------------------- |
| `empty.xlsx`                 | Smoke test                                  |
| `single-sheet-text.xlsx`     | SST, basic cells                            |
| `single-sheet-formula.xlsx`  | Formulas, cached values, calcChain          |
| `multi-sheet-cross-ref.xlsx` | Cross-sheet formulas, defined names         |
| `with-comments.xlsx`         | Legacy + threaded comments + persons        |
| `with-pivot.xlsx`            | Pivot table preservation (opaque)           |
| `with-chart.xlsx`            | Chart + drawing parts (opaque)              |
| `with-images.xlsx`           | Media + drawings (opaque)                   |
| `with-vba.xlsx` (`.xlsm`)    | VBA project + signature (opaque)            |
| `with-cf.xlsx`               | Conditional formatting full round-trip      |
| `with-data-validations.xlsx` | Data validations full round-trip            |
| `with-tables.xlsx`           | Excel tables (opaque) + tableParts in sheet |
| `with-external-links.xlsx`   | External link parts (opaque)                |
| `large-50k.xlsx`             | 50k rows × 10 cols; performance budget      |
| `windows1252-comments.xlsx`  | EC-I4: legacy encoding in comments part     |
| `prefixed-namespace.xlsx`    | EC-I3: synthetic non-default xmlns prefix   |

Adding a new fixture: drop it in `fixtures/xlsx/synthetic/` and the
roundtrip test picks it up automatically.

We additionally run `scripts/run-libreoffice-roundtrip.mjs` against
every synthetic fixture (per [`analysis.md`](analysis.md) §3 risk #4)
to catch fidelity issues that the reparse-equality test misses (e.g.
LibreOffice rejecting an attribute order Excel accepts).

---

## 15. The "no edits" contract

The strongest property the serializer holds:

```
∀ buf where parseXlsx(buf) succeeds:
  serializeXlsx(parseXlsx(buf)) produces, for every part path P:
    sha256(output.readBytes(P)) === sha256(buf.unzipped.readBytes(P))
```

This is verified by `parse-roundtrip.test.ts` for every fixture above.
Any change to the parser/serializer that breaks this for a part outside
the dirty set is a bug. PRs that touch the serializer must run
`pnpm test:xlsx-roundtrip` and post the diff to the build log if any
fixture changes.

---

## 16. File layout

```
packages/xlsx/src/serializer/
  serialize.ts                # serializeXlsx entry point
  errors.ts                   # XlsxSerializeError + kind union
  workbook-xml.ts             # xl/workbook.xml
  sheet-xml.ts                # SheetJS pass + native injection
  sheet-schema-order.ts       # CT_Worksheet child ordering helper
  shared-strings.ts           # xl/sharedStrings.xml
  styles.ts                   # xl/styles.xml
  comments.ts                 # legacy + threaded
  conditional-format.ts       # native CF block builders
  data-validation.ts          # native DV block builders
  hyperlinks.ts               # native <hyperlinks> block + sheet rels
  rels.ts                     # rels graph upserts
  content-types.ts            # [Content_Types].xml upserts
  calc-chain.ts               # xl/calcChain.xml removal helper
  sjs-adapter.ts              # XLSX.writeXLSX wrapper
  serialize.test.ts ... etc.
```

Every file mirrors the parser's split (§14 of [`parser.md`](parser.md))
so a feature change touches the same-named file on both sides.
