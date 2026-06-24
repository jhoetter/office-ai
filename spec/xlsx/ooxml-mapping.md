# XLSX — OOXML ↔ Model Mapping

> The most important spec doc for `@officeai/xlsx`. Every typed model
> field declared in [`document-model.md`](document-model.md) maps to a
> row here, in both directions: how the parser populates the field
> from OOXML, and how the serializer emits OOXML from the field.
> Anything not listed becomes an `OpaquePart` (whole part) or
> `OpaqueXml` (subtree) and is round-tripped byte-identical.
>
> Companion files: [`document-model.md`](document-model.md) for the
> shapes referenced below, [`feature-scope.md`](feature-scope.md) for
> what the agent may mutate vs preserve, and
> [`analysis-sheetjs.md`](analysis-sheetjs.md) for which OOXML parts
> SheetJS reads/writes vs drops (drives the opaque list in §15).

## 1. Namespaces

We **match by URI, not by prefix**. The prefix in the source workbook
is preserved in the model's root-attribute capture
(`Sheet.sheetRootAttrs`, `XlsxWorkbook.workbookRootAttrs`) and re-emitted
verbatim. The canonical URIs:

| Preferred prefix | Namespace URI                                                                  | Used for                                                             |
| ---------------- | ------------------------------------------------------------------------------ | -------------------------------------------------------------------- |
| `x:` (default)   | `http://schemas.openxmlformats.org/spreadsheetml/2006/main`                    | All `xl/workbook.xml`, `xl/worksheets/sheet*.xml`, styles, sst, etc. |
| `r:`             | `http://schemas.openxmlformats.org/officeDocument/2006/relationships`          | `r:id` cross-references (hyperlinks, comments, drawings, theme, …)   |
| `xr:`            | `http://schemas.microsoft.com/office/spreadsheetml/2009/9/main`                | `xr:uid` revision ids, `xr:revIDLastSave`                            |
| `xr2:` `xr3:`    | `http://schemas.microsoft.com/office/spreadsheetml/2014/revision` (and `…/15`) | Newer revision metadata; preserved opaquely                          |
| `xdr:`           | `http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing`          | `xl/drawings/drawing*.xml` shape anchors (opaque P0)                 |
| `a:`             | `http://schemas.openxmlformats.org/drawingml/2006/main`                        | DrawingML content (opaque P0)                                        |
| `c:`             | `http://schemas.openxmlformats.org/drawingml/2006/chart`                       | `xl/charts/chart*.xml` (opaque P0)                                   |
| `mc:`            | `http://schemas.openxmlformats.org/markup-compatibility/2006`                  | `mc:Ignorable`, `mc:AlternateContent` (preserved verbatim)           |
| `xml:`           | `http://www.w3.org/XML/1998/namespace`                                         | `xml:space="preserve"` on `<t>` and `<is>`                           |
| `tc:` / `tcXml:` | `http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments`      | `xl/threadedComments/threadedComment*.xml`                           |

> NOTE: A workbook produced by Google Sheets often binds the main
> SpreadsheetML namespace to a non-default prefix (e.g. `s:`). Our
> parser detects this by URI and re-binds; the serializer reuses
> whatever prefix was on disk. `mc:Ignorable` lists are preserved
> verbatim so future Excel features survive round-trip.

## 2. Container parts

`OoxmlContainer` (from `@officeai/core`) loads every zip entry on read
and exposes them as `parts: Map<path, OoxmlPart>`. The serializer
re-emits each part either from its original bytes (when not dirty) or
from the typed model (when dirty). Per-part dirtiness is tracked in
`XlsxSnapshot.dirty`; see [`document-model.md`](document-model.md) §1.

| Part path                                           | Loaded into model                                                  |      Editable      | Notes                                                                                                                                                        |
| --------------------------------------------------- | ------------------------------------------------------------------ | :----------------: | ------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `[Content_Types].xml`                               | container-only                                                     | only on add/remove | Managed by `OoxmlContainer`; we add overrides for new parts (charts when authored, threaded comments part on first reply, etc.). Re-emit verbatim otherwise. |
| `_rels/.rels`                                       | container-only                                                     |         no         | Top-level rels (workbook, core props, app props, custom props, thumbnail). Never re-emitted unless we add a top-level part.                                  |
| `xl/workbook.xml`                                   | `XlsxWorkbook.{sheets, definedNames, date1904, workbookRootAttrs}` |      **yes**       | Sheet list, defined names, workbookPr, bookViews. See §3.                                                                                                    |
| `xl/_rels/workbook.xml.rels`                        | implicit (drives sheet→file mapping)                               |      **yes**       | Re-built on workbook dirty. New rels minted on add-sheet / link to a new shared part.                                                                        |
| `xl/worksheets/sheet{N}.xml`                        | one `Sheet` each                                                   |      **yes**       | Cells, merges, hyperlinks, conditional formatting, autoFilter, sheetViews. See §4.                                                                           |
| `xl/worksheets/_rels/sheet{N}.xml.rels`             | drives `Hyperlink.target` URL resolution + comments anchor         |      **yes**       | Per-sheet rels. Re-emitted when the sheet's hyperlinks or comments change.                                                                                   |
| `xl/sharedStrings.xml`                              | `XlsxWorkbook.sharedStrings`                                       |      **yes**       | We always write SST (i.e. SheetJS option `bookSST: true`) so rich-text round-trips. See §6.                                                                  |
| `xl/styles.xml`                                     | `XlsxWorkbook.styleTable`                                          |      **yes**       | numFmts, fonts, fills, borders, cellXfs, cellStyleXfs, dxfs, namedStyles, tableStyles. See §7.                                                               |
| `xl/comments{N}.xml`                                | `Sheet.comments` (legacy half)                                     |      **yes**       | Per-sheet legacy comments (text + author + cell ref).                                                                                                        |
| `xl/threadedComments/threadedComment{N}.xml`        | `Sheet.comments` (modern half)                                     |      **yes**       | Threading metadata: id, parentId, createdAt, resolved, person ref. New on first threaded write.                                                              |
| `xl/persons/person.xml`                             | implicit (resolves threadedComment author)                         |      **yes**       | Author registry; new entry per first-time author.                                                                                                            |
| `xl/drawings/vmlDrawing{N}.vml`                     | container-only                                                     |    only on add     | Legacy comment anchors. Preserved verbatim; new ones written when adding the first comment to a sheet that has none.                                         |
| `xl/calcChain.xml`                                  | container-only                                                     |         no         | Excel rebuilds on first recalc. We drop on write (we don't re-emit calcChain in P0).                                                                         |
| `xl/theme/theme1.xml`                               | `XlsxWorkbook.theme` (color scheme only)                           |         no         | Color palette parsed for color-resolution; rest preserved opaquely. Other theme parts (theme2.xml, theme3.xml) fully opaque.                                 |
| `xl/charts/chart{N}.xml`                            | `XlsxWorkbook.opaqueParts`                                         |         no         | Charts opaque P0; renderer uses image fallback (see §13).                                                                                                    |
| `xl/charts/colors{N}.xml`, `xl/charts/style{N}.xml` | `XlsxWorkbook.opaqueParts`                                         |         no         | Chart styles. Opaque.                                                                                                                                        |
| `xl/charts/_rels/chart{N}.xml.rels`                 | container-only                                                     |         no         | Chart's own rels (e.g. embedded image references). Opaque.                                                                                                   |
| `xl/drawings/drawing{N}.xml`                        | `XlsxWorkbook.opaqueParts`                                         |         no         | Anchors charts and images to cells. Opaque P0.                                                                                                               |
| `xl/drawings/_rels/drawing{N}.xml.rels`             | container-only                                                     |         no         | Drawing rels (chart, image refs). Opaque.                                                                                                                    |
| `xl/media/image{N}.{png,jpg,gif,…}`                 | `XlsxWorkbook.opaqueParts`                                         |         no         | Embedded image bytes. Opaque P0; agent does not insert images in P0.                                                                                         |
| `xl/embeddings/oleObject{N}.bin`                    | `XlsxWorkbook.opaqueParts`                                         |         no         | Embedded OLE objects. Opaque.                                                                                                                                |
| `xl/tables/table{N}.xml`                            | `XlsxWorkbook.opaqueParts`                                         |         no         | Excel ListObjects. Opaque P0.                                                                                                                                |
| `xl/pivotTables/pivotTable{N}.xml`                  | `XlsxWorkbook.opaqueParts`                                         |         no         | Pivot table definitions. Opaque (out-of-scope per `feature-scope.md`).                                                                                       |
| `xl/pivotCache/pivotCacheDefinition{N}.xml`         | `XlsxWorkbook.opaqueParts`                                         |         no         | Pivot cache definitions. Opaque.                                                                                                                             |
| `xl/pivotCache/pivotCacheRecords{N}.xml`            | `XlsxWorkbook.opaqueParts`                                         |         no         | Pivot cache records. Opaque.                                                                                                                                 |
| `xl/queryTables/queryTable{N}.xml`                  | `XlsxWorkbook.opaqueParts`                                         |         no         | Power Query connections. Opaque.                                                                                                                             |
| `xl/connections.xml`                                | `XlsxWorkbook.opaqueParts`                                         |         no         | External data connections. Opaque.                                                                                                                           |
| `xl/model/*.xml`                                    | `XlsxWorkbook.opaqueParts`                                         |         no         | Power Pivot data model. Opaque.                                                                                                                              |
| `xl/slicers/slicer{N}.xml`                          | `XlsxWorkbook.opaqueParts`                                         |         no         | Slicers. Opaque.                                                                                                                                             |
| `xl/slicerCaches/slicerCache{N}.xml`                | `XlsxWorkbook.opaqueParts`                                         |         no         | Slicer caches. Opaque.                                                                                                                                       |
| `xl/timelines/timeline{N}.xml`                      | `XlsxWorkbook.opaqueParts`                                         |         no         | Timelines. Opaque.                                                                                                                                           |
| `xl/timelineCaches/timelineCache{N}.xml`            | `XlsxWorkbook.opaqueParts`                                         |         no         | Timeline caches. Opaque.                                                                                                                                     |
| `xl/externalLinks/externalLink{N}.xml`              | `XlsxWorkbook.opaqueParts`                                         |         no         | External workbook references. Opaque.                                                                                                                        |
| `xl/externalLinks/_rels/externalLink{N}.xml.rels`   | container-only                                                     |         no         | External-link rels. Opaque.                                                                                                                                  |
| `xl/customXml/item{N}.xml`                          | `XlsxWorkbook.opaqueParts`                                         |         no         | Custom XML data binding. Opaque.                                                                                                                             |
| `xl/customXml/itemProps{N}.xml`                     | `XlsxWorkbook.opaqueParts`                                         |         no         | Custom XML props. Opaque.                                                                                                                                    |
| `customXml/*` (top-level)                           | `XlsxWorkbook.opaqueParts`                                         |         no         | Top-level custom XML (document-host metadata, content type bindings). Opaque.                                                                                |
| `xl/vbaProject.bin`                                 | `XlsxWorkbook.opaqueParts`                                         |         no         | VBA macros. Opaque, NEVER executed.                                                                                                                          |
| `xl/vbaProjectSignature.bin`                        | `XlsxWorkbook.opaqueParts`                                         |         no         | VBA digital signature. Opaque (we capture it ourselves; SheetJS drops it).                                                                                   |
| `xl/printerSettings/printerSettings{N}.bin`         | `XlsxWorkbook.opaqueParts`                                         |         no         | Printer driver settings. Opaque.                                                                                                                             |
| `xl/activeX/activeX{N}.xml` + `.bin`                | `XlsxWorkbook.opaqueParts`                                         |         no         | ActiveX controls. Opaque.                                                                                                                                    |
| `xl/ctrlProps/ctrlProp{N}.xml`                      | `XlsxWorkbook.opaqueParts`                                         |         no         | Form-control properties. Opaque.                                                                                                                             |
| `xl/metadata.xml`                                   | container-only                                                     |      on edit       | XLDAPR cell-metadata. Re-emitted from a fixed template when any cell carries `dynamicArray: true`. Otherwise round-trips verbatim.                           |
| `docProps/core.xml`, `docProps/app.xml`             | container-only                                                     |         no         | Workbook properties. Re-emitted verbatim. Updating these is P1.                                                                                              |
| `docProps/custom.xml`                               | container-only                                                     |         no         | Custom workbook properties. Verbatim.                                                                                                                        |

For every "opaque" entry above, the part is stored as a `Uint8Array`
in `OoxmlContainer.parts` keyed by full zip path, with its
`contentType` resolved from `[Content_Types].xml`. The model exposes
the same set via `XlsxWorkbook.opaqueParts: ReadonlyMap<string, OpaquePart>`
for ergonomic access (each `OpaquePart` carries `{ path, bytes,
contentType, hash }`).

## 3. Workbook part — `xl/workbook.xml`

| OOXML element / attribute                                              | Model field                                        | Notes                                                                         |
| ---------------------------------------------------------------------- | -------------------------------------------------- | ----------------------------------------------------------------------------- | --------------------------- |
| `<workbook>` root attrs (xmlns:\*, mc:Ignorable)                       | `XlsxWorkbook.workbookRootAttrs`                   | Captured verbatim and re-emitted on serialize.                                |
| `<workbookPr date1904="1">`                                            | `XlsxWorkbook.date1904`                            | Boolean-coerced (`"1"` / `"true"` / true; everything else false).             |
| `<workbookPr codeName="ThisWorkbook">`                                 | preserved opaquely on the part                     | Returns from `XlsxWorkbook.opaqueWorkbookPr`. Round-tripped.                  |
| `<bookViews>/<workbookView>`                                           | preserved opaquely on the part                     | Active sheet, window position. Not exposed on the typed model in P0.          |
| `<sheets>/<sheet>` `name`                                              | `Sheet.name`                                       | UTF-8, ≤ 31 chars.                                                            |
| `<sheets>/<sheet>` `sheetId`                                           | `Sheet.sheetId`                                    | Stored as string.                                                             |
| `<sheets>/<sheet>` `r:id`                                              | implicit; resolves to `xl/worksheets/sheet{N}.xml` | Drives the per-sheet part path. Resolved via `xl/_rels/workbook.xml.rels`.    |
| `<sheets>/<sheet>` `state="hidden                                      | veryHidden"`                                       | `Sheet.hidden`                                                                | Absent attribute = visible. |
| Sheet position in `<sheets>` document order                            | `Sheet.index`                                      | 0-based positional index. Stable until a reorder.                             |
| `<definedNames>/<definedName name="…">…</definedName>`                 | `XlsxWorkbook.definedNames[i].{name, ref}`         | `<definedName>` text content → `ref`. See §10.                                |
| `<definedNames>/<definedName localSheetId="N">`                        | `DefinedName.scope = N` (number)                   | When absent, scope is `"workbook"`.                                           |
| `<definedNames>/<definedName comment="…">`                             | `DefinedName.comment`                              |                                                                               |
| `<definedNames>/<definedName hidden="1">`                              | `DefinedName.hidden`                               |                                                                               |
| name `_xlnm._FilterDatabase`, `_xlnm.Print_Area`, `_xlnm.Print_Titles` | `DefinedName.builtin = true`                       | The agent surface refuses to mutate built-in names.                           |
| `<calcPr>`                                                             | preserved opaquely on the part                     | Re-emitted on dirty-rewrite of the part.                                      |
| `<pivotCaches>/<pivotCache>`                                           | preserved opaquely on the part                     | We never edit pivot caches.                                                   |
| `<extLst>` and any other unknown child                                 | preserved opaquely on the part                     | Survives round-trip via the part-bytes fallback when nothing inside is dirty. |

> NOTE: The serializer regenerates `xl/workbook.xml` only when
> `dirty.workbook === true`. The regeneration emits typed fields
> (`<sheets>`, `<definedNames>`) and splices in the captured opaque
> children (`<workbookPr opaque attrs>`, `<bookViews>`, `<calcPr>`,
> `<pivotCaches>`, `<extLst>`) at their original positions to preserve
> Excel's element-order tolerance.

## 4. Worksheet part — `xl/worksheets/sheet{N}.xml`

The worksheet is the largest single XML part. We model the parts the
agent edits and preserve the rest opaquely.

### 4.1 Worksheet root and sheet-level metadata

| OOXML element / attribute                                                                                                                                                                                                                                                                                                                  | Model field                                              | Notes                                                                                                                                                                        |
| ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | -------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `<worksheet>` root attrs                                                                                                                                                                                                                                                                                                                   | `Sheet.sheetRootAttrs`                                   | Namespace declarations preserved verbatim.                                                                                                                                   |
| `<sheetPr>/<tabColor rgb="FF112233"/>`                                                                                                                                                                                                                                                                                                     | `Sheet.tabColor = "112233"`                              | We strip the leading 2-hex alpha byte; restore as `"FF" + rgb` on serialize.                                                                                                 |
| `<sheetPr>/<tabColor theme="4" tint="-0.25"/>`                                                                                                                                                                                                                                                                                             | `Sheet.tabColor` resolved against theme                  | We resolve theme refs at parse time using `XlsxWorkbook.theme.colors[index]` and the OOXML tint formula. Original `theme/tint` is preserved opaquely so re-emit is lossless. |
| `<sheetPr>` other attrs (`codeName`, `enableFormatConditionsCalculation`)                                                                                                                                                                                                                                                                  | `Sheet.opaqueWorksheetChildren` (whole `<sheetPr>` node) | Round-tripped verbatim.                                                                                                                                                      |
| `<dimension ref="A1:E20"/>`                                                                                                                                                                                                                                                                                                                | `Sheet.dimensions`                                       | Recomputed on serialize from `cells` keys; ignored on parse if present (we trust the cells map).                                                                             |
| `<sheetViews>/<sheetView>`                                                                                                                                                                                                                                                                                                                 | partly modeled, rest opaque                              | See §4.2.                                                                                                                                                                    |
| `<sheetFormatPr>`                                                                                                                                                                                                                                                                                                                          | preserved opaquely on `Sheet.opaqueWorksheetChildren`    | Default row height, customHeight flag, baseColWidth.                                                                                                                         |
| `<cols>/<col>`                                                                                                                                                                                                                                                                                                                             | `Sheet.cols`                                             | See §4.3.                                                                                                                                                                    |
| `<sheetData>/<row>/<c>`                                                                                                                                                                                                                                                                                                                    | `Sheet.cells`                                            | The bulk of the part. See §5.                                                                                                                                                |
| `<mergeCells>/<mergeCell ref="A1:B2"/>`                                                                                                                                                                                                                                                                                                    | `Sheet.merges`                                           | One `MergedCell` per `<mergeCell>`. The `count` attribute on `<mergeCells>` is recomputed on serialize.                                                                      |
| `<conditionalFormatting sqref="…"><cfRule …/>`                                                                                                                                                                                                                                                                                             | `Sheet.conditionalFormats`                               | See §11.                                                                                                                                                                     |
| `<dataValidations>/<dataValidation>`                                                                                                                                                                                                                                                                                                       | `Sheet.opaqueWorksheetChildren`                          | Preserved opaquely in P0; agent does not author validations.                                                                                                                 |
| `<hyperlinks>/<hyperlink ref="…" r:id="…"/>`                                                                                                                                                                                                                                                                                               | `Sheet.hyperlinks`                                       | See §9.                                                                                                                                                                      |
| `<printOptions>`, `<pageMargins>`, `<pageSetup>`, `<headerFooter>`, `<rowBreaks>`, `<colBreaks>`, `<customSheetViews>`, `<sheetProtection>`, `<protectedRanges>`, `<scenarios>`, `<dataConsolidate>`, `<phoneticPr>`, `<smartTags>`, `<webPublishItems>`, `<cellWatches>`, `<ignoredErrors>`, `<oleObjects>`, `<controls>`, `<tableParts>` | `Sheet.opaqueWorksheetChildren`                          | All preserved opaquely.                                                                                                                                                      |
| `<autoFilter ref="…">`                                                                                                                                                                                                                                                                                                                     | `Sheet.autoFilter.ref` + `autoFilter.opaqueConditions`   | See §12.                                                                                                                                                                     |
| `<drawing r:id="…"/>`                                                                                                                                                                                                                                                                                                                      | implicit (drives drawing-rel resolution)                 | Drawings are opaque parts; the rel is preserved on the sheet's rels file so the link survives.                                                                               |
| `<legacyDrawing r:id="…"/>`                                                                                                                                                                                                                                                                                                                | implicit (links to vmlDrawing for legacy comments)       | Re-emitted when the sheet's comments part exists.                                                                                                                            |
| `<extLst>`                                                                                                                                                                                                                                                                                                                                 | `Sheet.opaqueWorksheetChildren`                          | Modern extensions (sparklines, slicer cell refs, threaded-comment cell anchors). Preserved verbatim.                                                                         |

### 4.2 SheetViews and frozen panes

| OOXML element / attribute                                                                               | Model field                                     | Notes                                                                                                                     |
| ------------------------------------------------------------------------------------------------------- | ----------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------- |
| `<sheetViews>/<sheetView>` other attrs (`zoomScale`, `tabSelected`, `workbookViewId`, `rightToLeft`, …) | `Sheet.opaqueWorksheetChildren` (whole subtree) | We round-trip the entire `<sheetViews>` block opaquely except for the freeze-pane child, which we splice in/out.          |
| `<sheetView>/<pane state="frozen" xSplit="2" ySplit="1" topLeftCell="C2" activePane="…"/>`              | `Sheet.frozenCols = 2`, `Sheet.frozenRows = 1`  | `xSplit` ↔ `frozenCols`, `ySplit` ↔ `frozenRows`. Both default to 0 / undefined. The `topLeftCell` is recomputed on emit. |
| `<sheetView>/<pane state="split" …/>`                                                                   | preserved opaquely; not promoted to frozen      | Split (non-frozen) panes survive round-trip but are not exposed on the typed model in P0.                                 |
| `<sheetView>/<selection>`                                                                               | preserved opaquely                              | Selection is editor-session state, not persisted model state; we never overwrite it on save.                              |

> NOTE: The serializer's freeze-pane handler updates `<pane>` in-place
> within the captured `<sheetView>` opaque subtree when `frozenRows` /
> `frozenCols` change. When both are zero, the `<pane>` element is
> removed.

### 4.3 Columns and rows

| OOXML element / attribute                                                                                                    | Model field                                            | Notes                                                                                                                                                                                                       |
| ---------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------ | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `<cols>/<col min="3" max="5" width="14.5" customWidth="1" hidden="1" outlineLevel="2" style="3" bestFit="1" collapsed="1"/>` | `Sheet.cols.set(col, …)` for `col` in `[min-1, max-1]` | The OOXML form is range-coalesced; we expand to per-column entries on parse and re-coalesce contiguous identical entries on serialize. `bestFit`, `collapsed`, etc. preserved on the row entry's catch-all. |
| `<sheetData>/<row r="3" ht="22" customHeight="1" hidden="1" outlineLevel="1" s="4" customFormat="1">…</row>`                 | `Sheet.rows.set(row, …)` (key `r-1`)                   | Row attributes parsed into `RowProperties`. `r` attribute is 1-based on the wire; we store 0-based.                                                                                                         |
| `<sheetData>/<row spans="1:5">`                                                                                              | recomputed on emit                                     | `spans` is informational; we re-derive from cell range on serialize.                                                                                                                                        |

## 5. Cells — `<sheetData>/<row>/<c>`

The cell element is the workhorse of OOXML. The full grammar is:

```xml
<c r="A1"
   s="3"
   t="n|s|b|e|str|inlineStr|d"
   cm="1"
   ph="1"
   vm="0">
  <f t="normal|shared|array|dataTable" si="0" ref="A1:A10" aca="1" ca="1" bx="1" del1="1" del2="1" r1="…" r2="…" dt2D="1" dtr="1">…formula text…</f>
  <v>…cached value…</v>
  <is><t xml:space="preserve">…inline string…</t></is>
</c>
```

### 5.1 Cell attributes

| OOXML attribute                        | Model field                                     | Notes                                                                                                         |
| -------------------------------------- | ----------------------------------------------- | ------------------------------------------------------------------------------------------------------------- |
| `r="A1"`                               | `Cell.row`, `Cell.col`                          | Parsed via `a1ToRowCol`; both indices stored 0-based. Required by the OOXML spec.                             |
| `s="3"`                                | `Cell.styleId`                                  | Index into `XlsxWorkbook.styleTable.cellXfs`. Absent or `"0"` = default style; we store undefined for `"0"`.  |
| `t="n"` / absent                       | `Cell.value: number`                            | `<v>` parsed as `Number(text)`. Default when `t` absent is `"n"`.                                             |
| `t="s"`                                | `Cell.value: string \| RichText`                | `<v>` is the SST index; resolved to `sharedStrings.entries[index].plain` (string) or `.rich` (RichText).      |
| `t="str"`                              | `Cell.value: string`                            | "String-typed formula result". `<v>` text used directly.                                                      |
| `t="inlineStr"`                        | `Cell.value: string \| RichText`                | `<is>` is the inline string subtree. Plain `<is><t>…</t></is>` → `string`; `<is><r>…</r>…</is>` → `RichText`. |
| `t="b"`                                | `Cell.value: boolean`                           | `<v>` is `"0"` or `"1"`.                                                                                      |
| `t="e"`                                | `Cell.value: CellError`                         | `<v>` is the literal error string (e.g. `"#REF!"`); we map to the `CellError` enum.                           |
| `t="d"`                                | `Cell.value: DateCell { iso }`                  | `<v>` is an ISO 8601 date. Rare in real workbooks (Excel almost always uses `t="n"` + date numFmt).           |
| no `t` attr but `numFmt` is date-typed | `Cell.value: DateCell { serial }`               | We detect date format codes via `model/numfmt-builtins.ts::isDateFormat(numFmtId)` and promote the value.     |
| `cm="1"`                               | `Cell.dynamicArray = true`                      | Cell-metadata pointer to an XLDAPR block in `xl/metadata.xml`. Treated as a flag.                             |
| `ph="1"`                               | preserved on the cell's pass-through projection | Phonetic-text indicator (Japanese workbooks). Not surfaced on the typed model in P0.                          |
| `vm="N"`                               | preserved on the cell's pass-through projection | Value-metadata index. Not surfaced on the typed model in P0.                                                  |

### 5.2 Per cell-type details

#### Number — `t="n"` (default)

```xml
<c r="A1" s="2"><v>42</v></c>
```

- Parse: `Cell.value = Number("42") === 42`.
- Emit: `<c r="A1" s="2"><v>42</v></c>` (no `t` attribute; `t="n"` is the default).
- Edge: `Number("Infinity")` and `Number("NaN")` are coerced to
  `CellError.Num` on parse; the OOXML wire form should never carry
  these but we've seen them from third-party generators.

#### Shared string — `t="s"`

```xml
<c r="B2" t="s"><v>17</v></c>
```

- Parse: `Cell.value = sharedStrings.entries[17].rich ?? sharedStrings.entries[17].plain`.
- Emit: append the value to `XlsxWorkbook.sharedStrings` (or reuse via
  `indexByPlain` for plain strings); emit `<v>{newIndex}</v>`.
- Rich text vs plain: rich-text SST entries always serialize as
  `RichText`; plain SST entries serialize as `string`.

#### Inline string — `t="inlineStr"`

```xml
<c r="C3" t="inlineStr"><is><r><rPr><b/></rPr><t>Bold</t></r><t>plain</t></is></c>
```

- Parse: `<is>` contents converted to `RichText` (or `string` when no
  runs).
- Emit: we **promote inline strings to SST on serialize** (set the
  shared-string index and emit `t="s"`). This matches Excel's behavior
  when re-saving.
- Round-trip: original `t="inlineStr"` cells survive the round-trip as
  SST entries, not as `<is>`. This is intentional and documented as a
  benign deviation.

#### Formula-result string — `t="str"`

```xml
<c r="D4" t="str"><f>UPPER("hi")</f><v>HI</v></c>
```

- Parse: `Cell.value = "HI"`, `Cell.formula = { text: 'UPPER("hi")' }`.
- Emit: when the cached value is a string and a formula is present,
  re-emit as `t="str"`.

#### Boolean — `t="b"`

```xml
<c r="E5" t="b"><v>1</v></c>
```

- Parse: `Cell.value = (text === "1")`.
- Emit: `<v>1</v>` or `<v>0</v>`.

#### Error — `t="e"`

```xml
<c r="F6" t="e"><v>#DIV/0!</v></c>
```

- Parse: `Cell.value = CellError.Div0` (mapped from the literal string).
- Emit: literal error string in `<v>`.

#### Empty cell

```xml
<c r="G7"/>
<c r="G8" s="3"/>
```

- Parse: cell omitted from `Sheet.cells` UNLESS it carries a `s` (style)
  attribute or the SheetJS adapter ran with `sheetStubs: true` and
  the cell appeared in the original. Style-only cells produce
  `Cell { value: null, styleId }`.
- Emit: cells with `value: null` and no `formula` and no `styleId`
  and no `hyperlinkId` and no `commentId` are dropped from the
  output (sparse OOXML).

### 5.3 Formulas

OOXML carries the formula text in the `<f>` child (without the leading
`=`) and the cached evaluation result in `<v>`. The model preserves
both.

| `<f>` attribute                               | Model field                                                                                                        | Notes                                                                                                                                                                                                 |
| --------------------------------------------- | ------------------------------------------------------------------------------------------------------------------ | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| no `t` (default `normal`)                     | `Cell.formula = { text }`                                                                                          | Standalone formula. Cached value in sibling `<v>`.                                                                                                                                                    |
| `t="shared"` `si="0"` `ref="A1:A10"` (corner) | `Cell.formula = { text, shared: { si, x: 0, y: 0 } }` and `arrayRef` unset                                         | Master formula for the shared group. The `ref` attribute defines the group's range; we store it implicitly via the shared `si` and follower offsets. The `x/y` are 0 on the corner.                   |
| `t="shared"` `si="0"` (follower)              | `Cell.formula = { text: '<masterText>', shared: { si, x, y } }` where `x = col - cornerCol`, `y = row - cornerRow` | Follower cells **inherit** the corner's text. We expand the shared formula on parse so every cell carries its own `text`; the `shared` metadata records the original group membership for round-trip. |
| `t="array"` `ref="A1:B5"`                     | `Cell.formula = { text, arrayRef }` on every cell in the range; `text` is non-empty only on the corner             | Array (CSE) formula. The non-corner cells get `formula = { text: "", arrayRef }` so their group membership survives.                                                                                  |
| `t="dataTable"`                               | preserved opaquely as `formula.text`                                                                               | `=TABLE(…)` rows/columns of a what-if table. We do not evaluate; the original text round-trips.                                                                                                       |
| `aca="1"` (always-calculate-array)            | preserved opaquely on `formula.text`                                                                               | The formula engine treats this as a hint; the flag round-trips via the cell's pass-through projection.                                                                                                |
| `ca="1"` (calculation-applied)                | preserved opaquely                                                                                                 | Indicates the cached value is up-to-date. We re-set on every recalc.                                                                                                                                  |

> NOTE: We **un-share shared formulas on parse** (every cell carries
> the expanded `text`) but **re-share on serialize** (group cells
> with the same `Formula.shared.si`, emit one corner with `t="shared"
ref="…" si="…"`, and the followers with `t="shared" si="…"` and no
> text). This matches what Excel would write and minimizes diff noise.
> Single-member shared groups become standalone formulas on serialize.

### 5.4 Cached value handling

- When `Cell.formula` is present and `Cell.value !== null`, `<v>` carries
  the cached result encoded per `Cell.value`'s type (`t="str"` for
  string, `t="b"` for boolean, `t="e"` for error, no `t` for number).
- When `Cell.formula` is present and `Cell.value === null`, no `<v>` is
  emitted (Excel will recalculate on open).
- When the formula engine runs and produces a fresh value, the
  mutation that wraps the recalc updates `Cell.value`; the serializer
  then emits `<v>` from that updated value.

## 6. Shared strings — `xl/sharedStrings.xml`

| OOXML element                                                                                                               | Model field                                                     | Notes                                                                                         |
| --------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------- | --------------------------------------------------------------------------------------------- |
| `<sst count="…" uniqueCount="…">`                                                                                           | recomputed on emit                                              | Both counts are derived; we trust our `entries.length` over whatever the file claims.         |
| `<si>/<t xml:space="preserve">…</t>`                                                                                        | `SharedStringEntry { plain: text }`                             | Plain SST entry. `xml:space="preserve"` is preserved when leading/trailing whitespace exists. |
| `<si>/<r>/<rPr>/{<b/>,<i/>,<u/>,<strike/>,<color rgb="…"/>,<sz val="11"/>,<rFont val="…"/>,<scheme val="minor"/>}/<t>…</t>` | `SharedStringEntry { plain, rich: { plain, runs: RichRun[] } }` | Rich-text SST entry. Each `<r>` becomes a `RichRun { text, font }`.                           |
| `<si>/<phoneticPr>` and `<rPh>`                                                                                             | preserved on the SST entry as opaque                            | Japanese phonetic guides. Round-tripped verbatim.                                             |

> NOTE: We always emit the SST (`bookSST: true` in SheetJS terms). This
> means even cells whose source was `t="inlineStr"` or `t="str"`
> reach the SST on serialize, ensuring rich-text round-trip and
> matching Excel's own re-save behavior.

## 7. Styles — `xl/styles.xml`

The styles part has many sub-tables. The mapping is direct: each
sub-table populates the corresponding `StyleTable` array.

### 7.1 Number formats — `<numFmts>`

| OOXML element                                 | Model field                            | Notes                                                                                                                                                                             |
| --------------------------------------------- | -------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `<numFmt numFmtId="164" formatCode="0.00%"/>` | `StyleTable.numFmts.set(164, "0.00%")` | Built-in IDs (0–163) are NOT emitted by Excel — we maintain a static map (`model/numfmt-builtins.ts`) of OOXML's built-in formats and merge with the file's `<numFmts>` on parse. |

### 7.2 Fonts — `<fonts>`

| OOXML element                                                    | Model field                        | Notes                                                   |
| ---------------------------------------------------------------- | ---------------------------------- | ------------------------------------------------------- | -------------------- | ---------------- | ------------------------------------ |
| `<font>/<name val="Calibri"/>`                                   | `Font.name`                        |                                                         |
| `<font>/<sz val="11"/>`                                          | `Font.size`                        |                                                         |
| `<font>/<b/>` (presence) or `<b val="0"/>` (false)               | `Font.bold`                        |                                                         |
| `<font>/<i/>`, `<font>/<strike/>`                                | `Font.italic`, `Font.strike`       | Same presence-vs-`val="0"` semantics.                   |
| `<font>/<u val="single                                           | double                             | singleAccounting                                        | doubleAccounting"/>` | `Font.underline` | Bare `<u/>` (no `val`) = `"single"`. |
| `<font>/<color rgb="FF112233"/>`                                 | `Font.color = { rgb: "FF112233" }` | ARGB form preserved.                                    |
| `<font>/<color theme="1" tint="-0.349"/>`                        | `Font.color = { theme, tint }`     |                                                         |
| `<font>/<color indexed="64"/>`                                   | `Font.color = { indexed }`         |                                                         |
| `<font>/<color auto="1"/>`                                       | `Font.color = { auto: true }`      |                                                         |
| `<font>/<vertAlign val="superscript                              | subscript                          | baseline"/>`                                            | `Font.vertAlign`     |                  |
| `<font>/<family val="2"/>`                                       | `Font.family`                      |                                                         |
| `<font>/<scheme val="major                                       | minor"/>`                          | `Font.scheme`                                           |                      |
| `<font>/{<charset/>,<outline/>,<shadow/>,<condense/>,<extend/>}` | preserved opaquely on the entry    | Round-trips verbatim; not surfaced on the typed `Font`. |

### 7.3 Fills — `<fills>`

| OOXML element                                                                           | Model field                                             | Notes                                                                     |
| --------------------------------------------------------------------------------------- | ------------------------------------------------------- | ------------------------------------------------------------------------- |
| `<fill>/<patternFill patternType="none"/>`                                              | `Fill { type: "none" }`                                 | First entry is always `none` (Excel convention).                          |
| `<fill>/<patternFill patternType="gray125"/>`                                           | `Fill { type: "pattern", patternType: "gray125" }`      | Second entry is conventionally `gray125`.                                 |
| `<fill>/<patternFill patternType="solid"><fgColor …/></patternFill>`                    | `Fill { type: "solid", color }`                         | The "fill color" is the `fgColor`, not the `bgColor`.                     |
| `<fill>/<patternFill patternType="darkVertical"><fgColor …/><bgColor …/></patternFill>` | `Fill { type: "pattern", patternType, color, bgColor }` |                                                                           |
| `<fill>/<gradientFill>…<stop>…</gradientFill>`                                          | `Fill { type: "gradient", gradient: OpaqueXml }`        | Gradient stops, type, degree, top/left/right/bottom — preserved opaquely. |

### 7.4 Borders — `<borders>`

| OOXML element                                                                                                                         | Model field                                                             | Notes                                                                                                                                        |
| ------------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------- |
| `<border diagonalUp="1" diagonalDown="1">/<top style="thin"><color rgb="FF000000"/></top><right …/><bottom …/><left …/><diagonal …/>` | `Border.{top, right, bottom, left, diagonal, diagonalUp, diagonalDown}` | Each side maps to a `BorderEdge { style, color }`. Empty `<top/>` (no `style` attr) = no border; absent `<color/>` = `Color { auto: true }`. |
| `<border>/{<vertical/>,<horizontal/>}`                                                                                                | preserved opaquely on the entry                                         | These appear inside `<dxf>` border records (used by table styles to draw inner gridlines). Not surfaced on `Border` in P0.                   |

### 7.5 Cell xfs — `<cellXfs>` and `<cellStyleXfs>`

```xml
<cellXfs count="…">
  <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
  <xf numFmtId="2" fontId="1" fillId="2" borderId="1" xfId="0"
      applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="0"
      quotePrefix="0" pivotButton="0">
    <alignment horizontal="center" vertical="center" wrapText="1" indent="0" textRotation="0" shrinkToFit="0"/>
    <protection locked="1" hidden="0"/>
  </xf>
</cellXfs>
```

| OOXML attribute / child                                                                           | Model field (`CellFormat`)           | Notes                                                                                  |
| ------------------------------------------------------------------------------------------------- | ------------------------------------ | -------------------------------------------------------------------------------------- |
| `numFmtId`                                                                                        | `CellFormat.numFmtId`                | Index into `StyleTable.numFmts`.                                                       |
| `fontId`                                                                                          | `CellFormat.fontId`                  | Index into `StyleTable.fonts`.                                                         |
| `fillId`                                                                                          | `CellFormat.fillId`                  |                                                                                        |
| `borderId`                                                                                        | `CellFormat.borderId`                |                                                                                        |
| `xfId`                                                                                            | `CellFormat.xfId`                    | Parent named-style (`cellStyleXfs`) index.                                             |
| `applyNumberFormat`, `applyFont`, `applyFill`, `applyBorder`, `applyAlignment`, `applyProtection` | corresponding `applyX` boolean field | When `false`, the xf inherits the corresponding aspect from its `cellStyleXfs` parent. |
| `quotePrefix`, `pivotButton`                                                                      | corresponding boolean field          |                                                                                        |
| `<alignment>`                                                                                     | `CellFormat.alignment`               | See `Alignment` mapping (§7.6).                                                        |
| `<protection>`                                                                                    | `CellFormat.protection`              |                                                                                        |
| `<extLst>` and unknown children                                                                   | preserved opaquely on the entry      | Future extensions survive.                                                             |

`<cellStyleXfs>` is parsed identically to `<cellXfs>` and populates
`StyleTable.cellStyleXfs`.

### 7.6 Alignment — `<alignment>` inside an `<xf>`

| OOXML attribute    | Model field              | Notes                                                                                                                        |
| ------------------ | ------------------------ | ---------------------------------------------------------------------------------------------------------------------------- | ------------------------ | ------------------------------ |
| `horizontal`       | `Alignment.horizontal`   | Enum values per the OOXML spec (`general`, `left`, `center`, `right`, `fill`, `justify`, `centerContinuous`, `distributed`). |
| `vertical`         | `Alignment.vertical`     | `top`, `center`, `bottom`, `justify`, `distributed`.                                                                         |
| `wrapText="1"`     | `Alignment.wrapText`     |                                                                                                                              |
| `textRotation="N"` | `Alignment.textRotation` | Integer; `255` = vertical text per the OOXML spec.                                                                           |
| `indent="N"`       | `Alignment.indent`       |                                                                                                                              |
| `shrinkToFit="1"`  | `Alignment.shrinkToFit`  |                                                                                                                              |
| `readingOrder="0   | 1                        | 2"`                                                                                                                          | `Alignment.readingOrder` | 0 = context, 1 = LTR, 2 = RTL. |
| `relativeIndent`   | preserved opaquely       | Excel-specific; not surfaced.                                                                                                |
| `justifyLastLine`  | preserved opaquely       |                                                                                                                              |

### 7.7 Differential formats — `<dxfs>` and named styles

| OOXML element                                                    | Model field                                             | Notes                                                                                                                                                 |
| ---------------------------------------------------------------- | ------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------- |
| `<dxfs>/<dxf>`                                                   | `StyleTable.dxfs[i]: Partial<CellFormat>`               | `<dxf>` carries any subset of `<font>`, `<fill>`, `<border>`, `<numFmt>`, `<alignment>`, `<protection>`. Referenced by `ConditionalFormatRule.dxfId`. |
| `<cellStyles>/<cellStyle name="Normal" xfId="0" builtinId="0"/>` | `StyleTable.namedStyles[i]: { name, xfId, builtinId? }` | Index into `cellStyleXfs` via `xfId`. `builtinId` is set for Excel's predefined styles ("Normal" = 0, "Heading 1" = 16, etc.).                        |
| `<tableStyles>/<tableStyle …>`                                   | `StyleTable.tableStyles[i]: OpaqueXml`                  | Preserved opaquely; pivot/list-object styling is out of scope for P0 authoring.                                                                       |
| `<colors>/<indexedColors>`, `<colors>/<mruColors>`               | preserved on the styles part as opaque                  | Custom palette + most-recently-used colors. Round-tripped.                                                                                            |
| `<extLst>`                                                       | preserved opaquely on the styles part                   |                                                                                                                                                       |

## 8. Comments — `xl/comments{N}.xml` + `xl/threadedComments/threadedComment{N}.xml`

OOXML stores comments in two parallel parts when threading is in use.
The model collapses both into `Sheet.comments`. The serializer
reconstitutes both parts from the same source data.

### 8.1 Legacy comments part

```xml
<comments xmlns="…/spreadsheetml/2006/main">
  <authors><author>Alice</author></authors>
  <commentList>
    <comment ref="B2" authorId="0" shapeId="0">
      <text><r><t>Look at this!</t></r></text>
    </comment>
  </commentList>
</comments>
```

| OOXML element / attribute      | Model field                                             | Notes                                             |
| ------------------------------ | ------------------------------------------------------- | ------------------------------------------------- |
| `<authors>/<author>`           | resolves `authorId` → `Comment.author`                  | Author registry per part. We re-emit identically. |
| `<comment ref="B2">`           | `Comment.ref` (parsed via `parseRangeRef`, single cell) |                                                   |
| `<comment authorId="N">`       | `Comment.author = authors[N]`                           |                                                   |
| `<comment>/<text>/<r>…</r>`    | `Comment.text: RichRun[]`                               | Same `<r>` shape as SST rich text (§6).           |
| `<comment>/<text>/<t>` (plain) | `Comment.text = [{ text }]`                             |                                                   |

### 8.2 Threaded comments part

```xml
<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments">
  <threadedComment ref="B2" dT="2026-04-18T10:00:00Z" personId="{GUID-1}" id="{TC-GUID-1}">
    <text>Top-level comment</text>
  </threadedComment>
  <threadedComment ref="B2" dT="2026-04-18T10:05:00Z" personId="{GUID-2}" id="{TC-GUID-2}" parentId="{TC-GUID-1}" done="1">
    <text>Reply that resolves the thread</text>
  </threadedComment>
</ThreadedComments>
```

| OOXML attribute / child | Model field                                              | Notes                                                                                                                              |
| ----------------------- | -------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------- |
| `id="{GUID}"`           | minted-side correlation                                  | We mint our own integer `Comment.id`; the OOXML GUID is preserved on the cell's pass-through projection so the thread round-trips. |
| `ref="B2"`              | `Comment.ref`                                            |                                                                                                                                    |
| `dT="ISO 8601"`         | `Comment.createdAt`                                      |                                                                                                                                    |
| `personId="{GUID}"`     | resolves to `Comment.author` via `xl/persons/person.xml` | The persons part holds `<person id="{GUID}" displayName="Alice"/>`.                                                                |
| `parentId="{GUID}"`     | `Comment.threadParentId`                                 | Resolved to the parent `Comment.id`.                                                                                               |
| `done="1"`              | `Comment.resolved = true`                                |                                                                                                                                    |
| `<text>` (plain text)   | `Comment.text = [{ text }]`                              | Threaded comments are plain-text only (no `<r>` runs in this namespace).                                                           |

> NOTE: When a workbook has both legacy and threaded comments for the
> same cell, the legacy entry is the "fallback view" Excel shows in
> non-threaded clients. Our model records the threaded version (with
> `createdAt`, `threadParentId`, `resolved`) and re-emits both parts
> on serialize so older clients still see something.

## 9. Hyperlinks — `<hyperlinks>` + `xl/worksheets/_rels/sheet{N}.xml.rels`

```xml
<!-- inside the worksheet -->
<hyperlinks>
  <hyperlink ref="A1:A3" r:id="rId4" tooltip="Click me" display="Visit"/>
  <hyperlink ref="B5"   location="Sheet2!A1" tooltip="Jump"/>
  <hyperlink ref="C7"   r:id="rId5"/>
</hyperlinks>

<!-- inside xl/worksheets/_rels/sheet1.xml.rels -->
<Relationships xmlns="…">
  <Relationship Id="rId4" Type="…/relationships/hyperlink" Target="https://example.com" TargetMode="External"/>
  <Relationship Id="rId5" Type="…/relationships/hyperlink" Target="mailto:foo@example.com" TargetMode="External"/>
</Relationships>
```

| OOXML element / attribute                             | Model field                                                             | Notes                                                                            |
| ----------------------------------------------------- | ----------------------------------------------------------------------- | -------------------------------------------------------------------------------- |
| `<hyperlink ref="A1:A3"/>`                            | `Hyperlink.ref` (via `parseRangeRef`)                                   |                                                                                  |
| `<hyperlink r:id="rId4"/>` resolved to a URL `Target` | `Hyperlink.target = { kind: "url", url }`                               | The rels file's `Target` attribute carries the URL when `TargetMode="External"`. |
| `Target="mailto:user@example.com?subject=Hi"`         | `Hyperlink.target = { kind: "mailto", email, subject }`                 | We split the `mailto:` scheme on `?` and lift `subject=` into the typed field.   |
| `<hyperlink location="Sheet2!A1"/>`                   | `Hyperlink.target = { kind: "sheet", ref: parseRangeRef("Sheet2!A1") }` | `location` is internal; no rels entry exists for this kind.                      |
| `<hyperlink location="MyDefinedName"/>`               | `Hyperlink.target = { kind: "definedName", name }`                      | When `location` matches a workbook defined name and is not an A1 ref.            |
| anything we couldn't classify                         | `Hyperlink.target = { kind: "raw", raw }`                               | Catch-all preserves the original `Target` string.                                |
| `tooltip`, `display` attrs                            | `Hyperlink.tooltip`, `Hyperlink.display`                                |                                                                                  |

The agent commands `xlsx:add-hyperlink` and `xlsx:remove-hyperlink`
rewrite both halves: `<hyperlink>` element in the sheet AND the rels
entry. Removing the last hyperlink in a sheet does not delete the
rels file (Excel tolerates an empty rels graph; doing so would dirty
unrelated rels).

## 10. Defined names — `xl/workbook.xml` `<definedNames>`

```xml
<definedNames>
  <definedName name="taxRate">Sheet1!$B$1</definedName>
  <definedName name="region" localSheetId="2" comment="EU only">'Inputs'!$A$1:$A$10</definedName>
  <definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">Sheet1!$A$1:$D$50</definedName>
</definedNames>
```

| OOXML attribute / element                      | Model field                      | Notes                                                                                                    |
| ---------------------------------------------- | -------------------------------- | -------------------------------------------------------------------------------------------------------- |
| `name="…"`                                     | `DefinedName.name`               | UTF-8. Names starting with `_xlnm.` are reserved built-ins.                                              |
| `localSheetId="N"`                             | `DefinedName.scope = N` (number) | Sheet-scoped name. Absent → `scope: "workbook"`.                                                         |
| (text content)                                 | `DefinedName.ref`                | The reference body, kept as the original string (may be a single ref, a multi-range list, or a formula). |
| `comment="…"`                                  | `DefinedName.comment`            |                                                                                                          |
| `hidden="1"`                                   | `DefinedName.hidden`             |                                                                                                          |
| name `_xlnm._FilterDatabase`                   | `DefinedName.builtin = true`     | Auto-managed by `xlsx:apply-auto-filter`. The agent surface refuses direct edits.                        |
| name `_xlnm.Print_Area` / `_xlnm.Print_Titles` | `DefinedName.builtin = true`     | Preserved verbatim in P0; the agent does not author print settings.                                      |

## 11. Conditional formatting — `<conditionalFormatting>` + `<cfRule>`

```xml
<conditionalFormatting sqref="A1:A100 C1:C100">
  <cfRule type="cellIs" priority="1" operator="greaterThan" dxfId="3" stopIfTrue="0">
    <formula>100</formula>
  </cfRule>
  <cfRule type="containsText" priority="2" operator="containsText" dxfId="5" text="ERROR">
    <formula>NOT(ISERROR(SEARCH("ERROR",A1)))</formula>
  </cfRule>
</conditionalFormatting>
<conditionalFormatting sqref="B1:B100">
  <cfRule type="colorScale" priority="3">
    <colorScale>
      <cfvo type="min"/>
      <cfvo type="percentile" val="50"/>
      <cfvo type="max"/>
      <color rgb="FFFF0000"/><color rgb="FFFFFF00"/><color rgb="FF00FF00"/>
    </colorScale>
  </cfRule>
</conditionalFormatting>
```

| `<cfRule type=…>`                                                                                                                                                        | Model rule kind                       | Notes                                                                                                                                                                                                           |
| ------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | ------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `type="cellIs"`                                                                                                                                                          | `CellValueRule { kind: "cellValue" }` | `operator` attribute → `CellValueRule.operator`. `<formula>` children → `CellValueRule.formulas` (one for unary operators, two for `between`/`notBetween`).                                                     |
| `type="containsText"`, `notContainsText`, `beginsWith`, `endsWith`                                                                                                       | `TextRule`                            | The `text` attribute carries the literal string; the `<formula>` child is the redundant Excel-style equivalent and is not exposed on the model.                                                                 |
| `type="timePeriod"`                                                                                                                                                      | `DateOccurringRule`                   | The `timePeriod` attribute (`today`, `yesterday`, `last7Days`, `thisWeek`, etc.) → `DateOccurringRule.timePeriod`.                                                                                              |
| `type="colorScale"`                                                                                                                                                      | `ColorScaleRule`                      | `<colorScale>/<cfvo>` children → `ColorScaleRule.stops[i].{type, value}`. `<colorScale>/<color>` children → `.color`. Two-color scales have two `cfvo`+`color` pairs; three-color scales have three.            |
| `type="dataBar"`                                                                                                                                                         | `DataBarRule`                         | `<dataBar>/<cfvo>` for min/max; `<dataBar>/<color>` for the bar color. `<dataBar minLength="…" maxLength="…">` attrs → `DataBarRule.minLength`/`maxLength`. `<extLst>` (axis customization) preserved opaquely. |
| `type="iconSet"`, `top10`, `aboveAverage`, `duplicateValues`, `uniqueValues`, `expression`, `notContainsBlanks`, `containsBlanks`, `containsErrors`, `notContainsErrors` | `OpaqueRule { raw }`                  | Preserved verbatim. The agent does not author these in P0; the renderer falls back to the dxf overlay only.                                                                                                     |

| `<cfRule>` attribute | Model field                 | Notes                                                                                                                        |
| -------------------- | --------------------------- | ---------------------------------------------------------------------------------------------------------------------------- |
| `priority="N"`       | `RuleBase.priority`         | 1-based; lower wins.                                                                                                         |
| `dxfId="N"`          | `RuleBase.dxfId`            | Index into `StyleTable.dxfs`.                                                                                                |
| `stopIfTrue="1"`     | `RuleBase.stopIfTrue`       |                                                                                                                              |
| (parent) `sqref="…"` | `RuleBase.sqref` (per rule) | We flatten: every rule carries its parent's `sqref`. On serialize, we re-group rules with identical `sqref` into one parent. |

## 12. Auto-filters — `<autoFilter ref="…">`

```xml
<autoFilter ref="A1:D100">
  <filterColumn colId="2">
    <filters><filter val="Apple"/><filter val="Banana"/></filters>
  </filterColumn>
</autoFilter>
```

| OOXML element / attribute                     | Model field                                            | Notes                                                                                                                                                  |
| --------------------------------------------- | ------------------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `<autoFilter ref="A1:D100"/>`                 | `Sheet.autoFilter.ref`                                 | When the sheet has no auto-filter, the field is undefined.                                                                                             |
| `<autoFilter>/{<filterColumn>,<sortState>}`   | `Sheet.autoFilter.opaqueConditions`                    | Preserved verbatim in P0. The agent applies/clears the filter range; conditions round-trip but are not authored.                                       |
| Implicit `_xlnm._FilterDatabase` defined name | `XlsxWorkbook.definedNames` entry with `builtin: true` | Excel auto-creates this name when `<autoFilter>` is set; we mirror the same behavior on `xlsx:apply-auto-filter`. Removed on `xlsx:clear-auto-filter`. |

## 13. Charts and images (opaque P0)

Charts and images consist of multiple linked parts:

```
xl/worksheets/sheet1.xml         → has <drawing r:id="rIdN"/>
xl/worksheets/_rels/sheet1.xml.rels → rIdN → xl/drawings/drawing1.xml
xl/drawings/drawing1.xml          → has <xdr:twoCellAnchor> with <xdr:graphicFrame> referencing a chart
xl/drawings/_rels/drawing1.xml.rels → rels to xl/charts/chart1.xml AND/OR xl/media/image1.png
xl/charts/chart1.xml              → the actual chart definition
xl/charts/colors1.xml, xl/charts/style1.xml → optional chart formatting
xl/charts/_rels/chart1.xml.rels   → may rel to embedded image fallback
xl/media/image1.png               → embedded raster (often the chart's image fallback)
```

All of these parts round-trip as `OpaquePart` in
`XlsxWorkbook.opaqueParts`. The renderer:

1. Resolves the `<drawing r:id="…"/>` on the worksheet.
2. Walks the drawing's `<xdr:twoCellAnchor>` for cell-anchor coordinates
   (parsed inline from the opaque XML — we don't surface a typed model).
3. Looks for an embedded image fallback in the chart's rels graph
   (Excel embeds a PNG of the rendered chart for non-Excel viewers).
4. If found, displays the PNG. If not, draws a placeholder
   "[Chart: name]" tile.

Authoring new charts is **out of scope** in P0. Modifying chart data
ranges is also out of scope. The agent surface refuses any command
that targets a chart anchor with a structured "chart-authoring out of
scope" error.

## 14. Themes — `xl/theme/theme{N}.xml`

| OOXML element                                                                                                                                   | Model field                                    | Notes                                                                                                                                                                          |
| ----------------------------------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `<a:theme>/<a:themeElements>/<a:clrScheme>/<a:srgbClr val="112233"/>` (per slot: `lt1`, `dk1`, `lt2`, `dk2`, `accent1..6`, `hlink`, `folHlink`) | `XlsxWorkbook.theme.colors[i] = { name, rgb }` | We parse only `theme1.xml`'s color scheme (12 indexable slots). Other theme parts (`theme2.xml`, …) and other theme elements (font scheme, format scheme) round-trip opaquely. |
| `<a:srgbClr val="…"/>` vs `<a:sysClr val="windowText" lastClr="…"/>`                                                                            | the latter falls back to `lastClr` for `rgb`   | `sysClr` resolves to its `lastClr` attribute (Excel's last-known system color).                                                                                                |
| Anything else in the theme                                                                                                                      | preserved on the part as opaque                | Round-tripped verbatim.                                                                                                                                                        |

> NOTE: Only `theme1.xml` is parsed for the color palette. Workbooks
> with a custom theme override (`xl/theme/theme2.xml`) preserve those
> bytes as `OpaquePart` but use `theme1.xml`'s palette for rendering.
> This is consistent with Excel's Office 365 behavior, which always
> reads the active theme from `theme1.xml`.

## 15. Opaque parts — full list

For every part below, the bytes are stored as a `Uint8Array` in
`OoxmlContainer.parts` keyed by full zip path, with `contentType`
resolved from `[Content_Types].xml`. Each is exposed via
`XlsxWorkbook.opaqueParts: ReadonlyMap<string, OpaquePart>` for
ergonomic access. None of these are mutated by the agent in P0.

| Part path glob                                                              | Reason                                                                                           |
| --------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------ |
| `xl/charts/*`                                                               | Chart authoring out of scope (P0 renderer uses image fallback).                                  |
| `xl/drawings/*`                                                             | Drawing anchors for charts and images. Opaque.                                                   |
| `xl/media/*`                                                                | Embedded images. Opaque.                                                                         |
| `xl/embeddings/*`                                                           | OLE embeddings. Opaque.                                                                          |
| `xl/tables/*`                                                               | Excel ListObjects. Opaque P0 (P1 promotes to typed `XlsxTable`).                                 |
| `xl/pivotTables/*`, `xl/pivotCache/*`                                       | Pivot tables and caches. Out of scope per `feature-scope.md`.                                    |
| `xl/queryTables/*`, `xl/connections.xml`                                    | Power Query connections. Out of scope.                                                           |
| `xl/model/*`                                                                | Power Pivot data model. Out of scope.                                                            |
| `xl/slicers/*`, `xl/slicerCaches/*`                                         | Slicers. Out of scope.                                                                           |
| `xl/timelines/*`, `xl/timelineCaches/*`                                     | Timelines. Out of scope.                                                                         |
| `xl/externalLinks/*`                                                        | External workbook references. Out of scope.                                                      |
| `xl/customXml/*`, `customXml/*`                                             | Custom XML data binding (document-host metadata, content-type metadata). Out of scope.           |
| `xl/vbaProject.bin`                                                         | VBA macros. Preserved, never executed.                                                           |
| `xl/vbaProjectSignature.bin`                                                | VBA digital signature. Preserved (we capture this ourselves; SheetJS drops it).                  |
| `xl/printerSettings/*`                                                      | Printer driver settings. Opaque.                                                                 |
| `xl/activeX/*`, `xl/ctrlProps/*`                                            | ActiveX controls and form-control properties. Opaque.                                            |
| `xl/persons/person.xml`                                                     | Threaded-comment author registry. We re-emit when modifying threaded comments; otherwise opaque. |
| `xl/theme/theme{2..N}.xml`                                                  | Non-default themes. Opaque (we parse only `theme1.xml`).                                         |
| `xl/calcChain.xml`                                                          | Calculation chain. Dropped on write; Excel rebuilds on first recalc.                             |
| any unrecognised part with content-type registered in `[Content_Types].xml` | Future-proofing. Catch-all opaque storage.                                                       |

> NOTE: The drop-on-write of `xl/calcChain.xml` is the **only** part
> we intentionally do not preserve on round-trip. All other untouched
> parts round-trip byte-identical (modulo zip-level recompression; see
> `OoxmlContainer.serialize` doc-comment).

## 16. Content types — `[Content_Types].xml`

`[Content_Types].xml` is managed end-to-end by `OoxmlContainer`.
The XLSX serializer never re-emits it from scratch; instead, it
calls `container.addContentTypeOverride(path, contentType)` whenever
a new part appears (e.g. on first threaded-comment write to a sheet
that previously had none). The serializer's invariant:

> The `[Content_Types].xml` part is dirty IFF a part was added or
> removed since load, OR a part's content type changed.

The default extensions block (`<Default Extension="rels" ContentType="…"/>`,
etc.) is preserved verbatim.

## 17. Top-level rels — `_rels/.rels`

The package-level rels carry the workbook (`rId1`), core props
(`rId2`), extended props (`rId3`), and optionally custom props,
thumbnail, and digital-signature origin. **We never re-emit
`_rels/.rels` unless we add a top-level part.** In P0 we do not add
top-level parts, so `_rels/.rels` always round-trips byte-identical.

Workbook-level rels (`xl/_rels/workbook.xml.rels`) are different —
they are re-emitted when the workbook part is dirty (sheet
add/remove, defined-name changes, theme refs).

## 18. Where we differ from SheetJS

For implementers familiar with SheetJS: we do **not** delegate the
parts listed in this section to `XLSX.read` / `XLSX.writeXLSX` even
though those API calls exist. We use SheetJS as a focused cell-value /
formula / SSF adapter only. The OOXML parts SheetJS would either
silently drop or rebuild from a minimal template — and which we
own ourselves — are:

1. `[Content_Types].xml` (SheetJS rebuilds; we manage via `OoxmlContainer`).
2. `_rels/.rels` (SheetJS rebuilds from a fixed set; we round-trip verbatim).
3. `xl/_rels/workbook.xml.rels` (SheetJS rebuilds; we re-emit only on dirty).
4. `xl/styles.xml` (SheetJS emits a minimal scaffold; we own full fidelity).
5. `xl/theme/*.xml` (SheetJS regenerates `theme1.xml`, drops the rest; we round-trip all).
6. `xl/sharedStrings.xml` (SheetJS optional; we always emit).
7. Conditional formatting, data validations, sheet protection details, page setup, header/footer, breaks, scenarios, controls, custom views, ignored errors, table parts (SheetJS drops; we capture in `Sheet.opaqueWorksheetChildren` or `Sheet.conditionalFormats`).
8. Tables, pivots, charts, drawings, images, slicers, timelines, custom XML, external links, VBA signature (SheetJS marks "TODO"; we capture as `OpaquePart`).
9. The `xl/metadata.xml` XLDAPR block (SheetJS hard-codes a template; we round-trip the original unless a cell's `dynamicArray` flag changed).
10. The full `xl/calcChain.xml` (SheetJS reads optional, never writes; we drop on write).

The serializer's rule of thumb: if a part is in `dirty.*`, regenerate
from typed model; otherwise re-emit from `OoxmlContainer.parts.get(path).bytes`.
