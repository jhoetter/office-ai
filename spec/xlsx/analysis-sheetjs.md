# SheetJS — Analysis for OfficeAI (clean-room)

> Reference commit: **`515d1c6f2e1d3ca422ee9198b177cfd926434936`** ("mini refresh
> [ci skip]") on the `master` branch of the SheetJS Community Edition mirror at
> <https://github.com/SheetJS/sheetjs>. This is the last source-bearing commit
> before the SheetJS team archived the GitHub mirror; subsequent development
> moved to <https://git.sheetjs.com/sheetjs/sheetjs>. The Community Edition is
> Apache 2.0 licensed (see `LICENSE`).
>
> Method: read-only inspection of `bits/*.js` and `types/index.d.ts` at the
> commit above, with `xlsx.js` (the pre-built bundle) cross-referenced for
> helper functions hoisted out of `bits/`. **No code from SheetJS has been
> copied into this document or into `@officeai/xlsx`. All snippets are
> paraphrased shape descriptions, not transcribed source.**
>
> Target audience: implementors of `@officeai/xlsx` who need to decide
> precisely what we can delegate to SheetJS at runtime versus what we must
> model ourselves and round-trip via opaque OOXML blobs.

---

## 0. Orientation: how the SheetJS source is laid out

SheetJS CE is assembled from many `bits/*.js` files concatenated in numeric
order by the `Makefile` to produce `xlsx.js`, `xlsx.mjs`, the mini variants,
and the ExtendScript build. TS ambient declarations live in
`types/index.d.ts`. The XLSX-relevant bits:

| File                    | Purpose                                                   |
| ----------------------- | --------------------------------------------------------- |
| `bits/10_ssf.js`        | "Spreadsheet format" — number-format codes (built-ins +). |
| `bits/30_ctype.js`      | `[Content_Types].xml` part-type registry.                 |
| `bits/31_rels.js`       | `_rels/*.rels`.                                           |
| `bits/42_sstxml.js`     | `xl/sharedStrings.xml`.                                   |
| `bits/47_styxml.js`     | `xl/styles.xml`.                                          |
| `bits/52_calcchain.js`  | `xl/calcChain.xml`.                                       |
| `bits/53_externlink.js` | `xl/externalLinks/*`.                                     |
| `bits/54_drawing.js`    | Drawings (chart resolution only).                         |
| `bits/55_vml.js`        | Legacy VML (only used for comment anchors).               |
| `bits/56-58_*cmnt*.js`  | Comments (xml + bin + threaded).                          |
| `bits/59_vba.js`        | VBA project passthrough.                                  |
| `bits/61-64_f*.js`      | Formula parsing helpers (no evaluation).                  |
| `bits/67_wsxml.js`      | Worksheet body (`xl/worksheets/sheet*.xml`).              |
| `bits/69_chartxml.js`   | Chart `numCache` extraction (cached values only).         |
| `bits/72_wbxml.js`      | `xl/workbook.xml`.                                        |
| `bits/85_parsezip.js`   | XLSX/XLSB top-level zip walker.                           |
| `bits/86_writezip.js`   | XLSX/XLSB writer.                                         |
| `bits/87_read.js`       | `XLSX.read`/`readFile`.                                   |
| `bits/88_write.js`      | `XLSX.write`/`writeXLSX`/`writeFile`.                     |
| `bits/95_api.js`        | `XLSX.utils` namespace.                                   |

`WTF` in options means "throw on unrecognised tag"; default is "swallow
silently" — SheetJS's house style for un-modelled parts is to drop them.

---

## 1. Inventory of OOXML parts SheetJS reads

The authoritative list of recognised `xl/*` parts is the `ct2type` map at
`bits/30_ctype.js:5`. Anything mapped to a non-`"TODO"` string is a part
SheetJS has a parser for; anything mapped to `"TODO"` is recognised by content
type but has no parsing path (it is silently dropped on read and not emitted
on write).

### 1.1 Parts SheetJS parses into the workbook object

Parsed from `xl/`:

| Part                                       | Driven by                                  | Notes                                                                                                                     |
| ------------------------------------------ | ------------------------------------------ | ------------------------------------------------------------------------------------------------------------------------- |
| `xl/workbook.xml`                          | `parse_wb_xml` (`bits/72_wbxml.js`)        | sheet list, defined names, `WBProps`, `bookViews`, codeName. Many subelements (pivotCaches, smartTagPr, etc.) are no-ops. |
| `xl/_rels/workbook.xml.rels`               | `parse_rels` (`bits/31_rels.js`)           | Used to resolve sheet → file mapping.                                                                                     |
| `xl/worksheets/sheet*.xml`                 | `parse_ws_xml` (`bits/67_wsxml.js`)        | See §1.3 for what _inside_ the worksheet is captured.                                                                     |
| `xl/worksheets/_rels/sheet*.xml.rels`      | `parse_rels`                               | Used for hyperlink targets, comments, drawings.                                                                           |
| `xl/sharedStrings.xml`                     | `parse_sst_xml` (`bits/42_sstxml.js`)      | Always parsed when referenced.                                                                                            |
| `xl/styles.xml`                            | `parse_sty_xml` (`bits/47_styxml.js`)      | Partial — see §4.                                                                                                         |
| `xl/theme/theme1.xml`                      | `parse_theme_xml` (`bits/49_theme.js`)     | Only when `opts.cellStyles` is set; only the colour scheme is consumed.                                                   |
| `xl/comments*.xml`                         | `parse_cmnt_xml` (`bits/57_cmntxml.js`)    | Hooked in via the worksheet rels graph.                                                                                   |
| `xl/threadedComments/threadedComment*.xml` | `parse_tcmnt_xml`                          | Modern threaded comments.                                                                                                 |
| `xl/persons/person.xml`                    | `parse_people_xml`                         | Used to resolve threaded-comment authors.                                                                                 |
| `xl/calcChain.xml`                         | `parse_cc_xml` (`bits/52_calcchain.js`)    | Only if `opts.bookDeps` is set.                                                                                           |
| `xl/externalLinks/externalLink*.xml`       | `parse_xlink` (`bits/53_externlink.js`)    | Read into `Workbook.externbooks`; **dropped on write**.                                                                   |
| `xl/metadata.xml`                          | `parse_xlmeta_xml` (`bits/51_xlsxmeta.js`) | Only the `XLDAPR` cell-metadata index, used to flag dynamic-array cells (`cell.D`).                                       |
| `xl/vbaProject.bin`                        | Raw passthrough                            | Only when `opts.bookVBA`. Stored on `wb.vbaraw` as a binary blob.                                                         |
| `xl/chartsheets/sheet*.xml`                | `parse_cs` + `parse_chart`                 | A chartsheet becomes a worksheet whose cells are populated from the chart's `numCache` only — see `bits/69_chartxml.js`.  |
| `xl/dialogsheets/sheet*.xml`               | `parse_ds`                                 | Treated as worksheets.                                                                                                    |
| `xl/macrosheets/sheet*.xml`                | `parse_ms`                                 | Excel-4.0-era XLM macrosheets.                                                                                            |

Parsed from outside `xl/`:

- `[Content_Types].xml` (`parse_ct`, `bits/30_ctype.js:230`).
- `_rels/.rels`.
- `docProps/core.xml`, `docProps/app.xml`, `docProps/custom.xml`
  (`bits/33_coreprops.js`, `bits/34_extprops.js`, `bits/35_custprops.js`).

What's parsed _inside_ the worksheet body (`bits/67_wsxml.js`):

- `dimension` → `ws['!ref']`.
- `sheetPr` → `wb.Workbook.Sheets[idx].CodeName` and an outline-pr writer
  hook (only `summaryBelow` and `summaryRight` survive).
- `sheetViews` → workbook-level zoom and RTL flag only. Per-view freeze
  panes, split panes, gridline visibility, header visibility, selection,
  and active-cell are **not** preserved.
- `cols` → column widths/levels/hidden, but only when `opts.cellStyles`.
- `sheetData` → cells (the bulk of the work).
- `mergeCells` → `ws['!merges']`.
- `hyperlinks` → resolved against the rels graph and pasted onto cells as
  `cell.l`.
- `autoFilter` → `ws['!autofilter'] = { ref }` (only the range is kept; no
  filter-column criteria).
- `pageMargins` → `ws['!margins']`.

That's it for the worksheet body. The list of standard worksheet children
that are **silently dropped** is much longer:

- `sheetFormatPr`, `pageSetup`, `headerFooter`, `printOptions`,
  `colBreaks`, `rowBreaks`, `customSheetViews`, `sheetProtection`
  (algorithm/hash/salt — `protectionRow` is a TODO comment in the source),
  `protectedRanges`, `scenarios`, `dataConsolidate`, `phoneticPr`,
  `conditionalFormatting`, `dataValidations`, `controls`, `oleObjects`,
  `customProperties`, `cellWatches`, `ignoredErrors`, `smartTags`,
  `webPublishItems`, `tableParts`, `extLst` (and all of the
  Excel-2010+/2016+/365 `extLst` extensions: sparklines, slicer refs,
  protected-range modern, threaded-comment cell anchor, etc.).

### 1.2 Recognised content types with no parser ("TODO" parts)

From `bits/30_ctype.js`, all parts mapped to `"TODO"` — i.e. recognised in
`[Content_Types].xml` but with no implementation:

- Pivot: `pivotTable+xml`, `pivotCacheDefinition+xml`,
  `pivotCacheRecords+xml`, plus the binary equivalents.
- Charts: `drawingml.chart+xml`, `chartcolorstyle+xml`, `chartstyle+xml`,
  `chartex+xml`, `drawingml.chartshapes+xml`.
- Diagrams: `diagramColors+xml`, `diagramData+xml`, `diagramLayout+xml`,
  `diagramStyle+xml`.
- Tables (ListObjects): `spreadsheetml.table+xml`, `ms-excel.table`,
  `tableSingleCells*`.
- Slicers / Timelines: `slicer*`, `slicerCache*`, `Timeline*`,
  `TimelineCache*`.
- Connections / Query Tables: `connections*`, `queryTable*`.
- Sort map: `wsSortMap`.
- Shared workbook revisions: `userNames*`, `revisionHeaders*`,
  `revisionLog*`.
- Drawings of any kind: `drawingml` parts above plus
  `vmlDrawing` (read of VML drawings is not in `ct2type`; the writer only
  emits VML for legacy comments).
- Theme override: `themeOverride+xml`.
- Embedded objects: `oleObject`, `image/png` (and other media types).
- Printer settings, ActiveX, attached toolbars, control properties.
- Custom XML: `customXmlProperties+xml`, `spreadsheetml.customProperty`.
- VBA signature: `vbaProjectSignature` (the project itself is preserved as
  a blob; the signature is dropped, which invalidates it).
- Volatile dependencies, data model, survey, web survey.

This is the canonical "what we have to layer ourselves" list (§8).

### 1.3 Parts preserved as-is on write

**None.** `bits/86_writezip.js` (function `write_zip_xlsx`) constructs the
output ZIP from scratch using only the SheetJS in-memory workbook object.
There is no "keep original blobs" path. The closest thing is `opts.bookFiles`
on the **read** side (`bits/85_parsezip.js:234`) which exposes the raw zip
entries for _inspection_, but nothing in the writer reads them back.

`opts.bookVBA`/`wb.vbaraw` is the only true round-trip blob: the binary
`xl/vbaProject.bin` is read in unmodified and written out unmodified.

This is the foundation of our opaque-blob layer (§11): SheetJS will never
preserve any of the TODO parts above for us; we have to keep the bytes
ourselves and re-zip them on write.

---

## 2. Inventory of OOXML parts SheetJS writes

`write_zip_xlsx` (`bits/86_writezip.js:136`) is the canonical reference. It
emits, in order:

1. `docProps/core.xml` (always).
2. `docProps/app.xml` (always; sheet names always overwritten from
   `wb.SheetNames`).
3. `docProps/custom.xml` — only when `wb.Custprops` differs from `wb.Props`
   _and_ has at least one key.
4. For each sheet `i` in `wb.SheetNames`:
   - `xl/worksheets/sheet{i+1}.xml` — every sheet gets a fresh worksheet
     XML, regardless of whether it was originally a chartsheet, dialogsheet,
     or macrosheet. The `_type` switch in `write_zip_xlsx` falls through to
     "default", and `case 'chart'` is explicitly fall-through.
   - `xl/comments{i+1}.xml` — only if `ws['!comments']`.
   - `xl/threadedComments/threadedComment{i+1}.xml` — only if any comment
     has `T == true` (threaded).
   - `xl/drawings/vmlDrawing{i+1}.vml` — only if `ws['!legacy']` and there
     are comments needing a VML anchor.
   - `xl/worksheets/_rels/sheet{i+1}.xml.rels` — only if any worksheet rels
     were added.
5. `xl/sharedStrings.xml` — only if `opts.bookSST` is true (default false).
   When false, strings are inlined as `<c t="str"><v>…</v></c>`. **Important
   nuance**: SheetJS's "inline" strings are _not_ the OOXML `inlineStr`
   form (`<c t="inlineStr"><is><t>…</t></is></c>`); they use `t="str"`,
   which is the "string-typed formula result" form. Excel reads it
   correctly, but text-search tools and validators looking for `inlineStr`
   won't find it.
6. `xl/workbook.xml` (always).
7. `xl/theme/theme1.xml` (always — and always a _fresh_ theme, not the
   original; `write_theme(wb.Themes, opts)` regenerates from a default
   template if `wb.Themes` is empty).
8. `xl/styles.xml` (always — see §4 for what is in it).
9. `xl/vbaProject.bin` (only if `wb.vbaraw` and `bookType` is one of
   xlsm/xlsb/xlam/xls).
10. `xl/metadata.xml` (**always** — the writer emits a hard-coded XLDAPR
    template regardless of what was read; see `bits/51_xlsxmeta.js:81`).
11. `xl/persons/person.xml` (only if there are >1 people, which the writer
    seeds with `"SheetJ5"`).
12. `[Content_Types].xml`.
13. `_rels/.rels`.
14. `xl/_rels/workbook.xml.rels`.

### 2.1 What the writer does with relationships

`opts.rels` (the package-level rels) and `opts.wbrels` (the workbook-level
rels) are built up by `add_rels` calls scattered across the writer. Per-sheet
`wsrels` are local. The relationship IDs are auto-allocated — original
`rId` values from the input are **not** preserved.

### 2.2 What the writer does with custom XML, the `_rels/` graph, etc.

- Custom XML (`customXml/item*.xml`, `customXml/itemProps*.xml`,
  `customXml/_rels/itemN.xml.rels`) is **not** read and **not** written.
- The `_rels/.rels` package rels graph is rebuilt from a small fixed set:
  workbook (`rId1`), core props (`rId2`), extended props (`rId3`),
  custom props (`rId4` if present). Any other top-level relationships
  (e.g. to a `customXml/` part, to a thumbnail, to a digital signature
  origin) are discarded.
- The workbook rels file is rebuilt from theme + styles + sharedStrings
  - per-sheet rels + metadata + people. Pivot caches, connections,
    external link refs, and any "extension" rels are dropped.

Net: SheetJS owns the ZIP as a whole. There is no notion of an "untouched
part". Our implementation must layer that ourselves.

---

## 3. Cell value model

Defined formally in `types/index.d.ts:597` and `:632`, and implemented by
`parse_ws_xml_data` (`bits/67_wsxml.js:316`) on read and `write_ws_xml_cell`
(`bits/67_wsxml.js:259`) on write.

### 3.1 Cell shape

A cell is a plain JS object with optional fields:

- `t` (required): the type tag — one of:
  - `'b'` boolean
  - `'n'` number (the default, including dates-as-serials)
  - `'e'` error
  - `'s'` string (resolved either from the SST or inline)
  - `'d'` date (only when `opts.cellDates` was set on read)
  - `'z'` stub (a present-but-empty cell; only emitted when
    `opts.sheetStubs`)
- `v`: the raw value. JS type depends on `t`:
  - `b` → boolean
  - `n` → number (a JS double)
  - `e` → number (the Excel error code; see §3.3)
  - `s` → string (the resolved text; rich-text markup is preserved
    separately on `r`)
  - `d` → JS `Date`
  - `z` → undefined
- `f` (optional): formula text without the leading `=`. SheetJS does _not_
  store `=`. When read, `_xlfn.` prefixes are stripped unless `opts.xlfn` is
  set.
- `F` (optional): the enclosing-range string for an array formula. Set on
  every cell that is part of the array, but the formula text itself is
  only on the corner cell.
- `w` (optional): the formatted text representation of the value (as Excel
  would display it given the number format). Computed on read when
  `opts.cellText` (default true).
- `z` (optional): the number-format code (the _string_ code, e.g.
  `"m/d/yy"`). Populated on read when `opts.cellNF`.
- `s` (optional): the _style_ — actually the resolved cell-xf object from
  `styles.CellXf[xfId]`, not the raw xfId. Populated when
  `opts.cellStyles`. Note that this is a copy-by-reference; if you mutate
  it, you mutate the shared style.
- `r` (optional): the rich-text _XML markup_ preserved verbatim. Available
  on `s` cells whose source was an SST entry with rich runs.
- `h` (optional): an HTML rendering of `r`, for browser display. Computed
  when `opts.cellHTML` (default true).
- `c` (optional): an array of comments anchored on this cell.
- `l` (optional): a hyperlink: `{ Target, Tooltip?, Rel?, location?,
display? }`. Populated from the `<hyperlink>` records resolved against
  the worksheet rels.
- `D` (optional): boolean, true iff the cell carries the `cm`
  cell-metadata attribute pointing at an `XLDAPR` metadata block — i.e.
  it's an Excel-365 dynamic-array spill anchor.

### 3.2 Numbers, dates, blanks

- **Numbers**: `parseFloat(p.v)` — an IEEE-754 double. Anything beyond
  ~15 significant digits silently loses precision on read. Excel itself
  has the same limit, but it _also_ has a separate "number stored as text"
  cell type that some workflows depend on; SheetJS conflates the two.
- **Dates**: default `t='n'`, `v` = OLE date serial (days since
  1900-01-00; `WBProps.date1904` flips to the 1904 system). With
  `opts.cellDates`, becomes `t='d'`, `v` = JS `Date` via `parseDate`.
  `parseDate` reads ISO strings as local time, so cross-TZ round-trips
  with `cellDates:true` can shift by a day around midnight.
- **Errors**: `cell.v` holds the numeric Excel error code, `cell.w` the
  string form. Codes (`bits/29_xlsenum.js`): `0x00 #NULL!`,
  `0x07 #DIV/0!`, `0x0F #VALUE!`, `0x17 #REF!`, `0x1D #NAME?`,
  `0x24 #NUM!`, `0x2A #N/A`, `0x2B #GETTING_DATA`.
- **Blanks**: `<c r="A1"/>` dropped unless `opts.sheetStubs`; then
  `{t:'z', v:undefined}`. The empty-SST string is `{t:'s', v:''}`, not
  a stub.
- **Booleans**: `parsexmlbool` (`"1"`, `"true"`, `"TRUE"`).

### 3.3 Sparse vs dense

Default: flat string-keyed object (`ws["A1"] = {...}`, plus `ws["!ref"]`,
`ws["!merges"]`, etc.). With `opts.dense = true`, a 2D array
(`ws[r][c]`). Dense form is more memory-efficient and easier to map onto
a tabular internal model; cell shape is identical between modes.

---

## 4. Style table

`parse_sty_xml` (`bits/47_styxml.js:379`) reads, in order, **only**:

1. `numFmts` → `styles.NumberFmt[id] = formatCode` (re-mapping ids > 392
   to recover Excel's 60+392 builtin range).
2. `fonts` → `styles.Fonts[]` with `{name, sz, bold, italic, underline,
strike, outline, shadow, condense, extend, vertAlign, family, scheme,
color}`.
3. `fills` → `styles.Fills[]` with `{patternType, fgColor, bgColor}`.
   Gradient-fill `stop` records become `<gradientFill>` close-ish wrappers,
   but the gradient stops themselves are dropped.
4. `borders` → `styles.Borders[]` — but inspect the source: the per-side
   parsing (`<left>`, `<right>`, `<top>`, `<bottom>`, `<diagonal>`,
   `<horizontal>`, `<vertical>`, `<start>`, `<end>`) is implemented as
   no-op cases. Only `diagonalUp`/`diagonalDown` flags survive. **Border
   styles and colours are silently dropped on read.**
5. `cellXfs` → `styles.CellXf[]` with `{numFmtId, fillId, fontId,
borderId, xfId, applyAlignment, applyBorder, applyFill, applyFont,
applyNumberFormat, applyProtection, pivotButton, quotePrefix,
alignment}`.

What's **not** parsed (the comment in `parse_sty_xml` flags these
explicitly):

- `cellStyleXfs` (the named-style table) — completely ignored.
- `cellStyles` (the human-readable style names: "Normal", "Heading 1",
  "Comma" …) — ignored.
- `dxfs` (differential formatting records used by conditional formatting,
  table styles, pivot styles) — ignored.
- `tableStyles` — ignored.
- `colors` (custom `indexedColors` / `mruColors` palette) — ignored.
- `extLst` extensions (used by Excel 2010+ for things like advanced fill
  effects).

### 4.1 Styles on write

`write_sty_xml` (`bits/47_styxml.js:413`) does _not_ round-trip the style
table. It emits a near-fixed scaffold:

- `numFmts`: rebuilt from `wb.SSF` via `write_numFmts`, restricted to
  specific ID ranges (5–8, 23–26, 41–44, 50+).
- `fonts`: hardcoded single Calibri-12.
- `fills`: hardcoded `none` + `gray125`.
- `borders`: hardcoded single empty border.
- `cellStyleXfs`: hardcoded single default xf.
- `cellXfs`: dynamic, populated by `get_cell_style` during write,
  deduped on the `numFmtId/fillId/fontId/borderId` + alignment quintuple.
- `cellStyles`: hardcoded single `Normal`.
- `dxfs count="0"`, `tableStyles count="0"`.

Practical consequence: even if the input had ten fonts, a dozen named
styles, and dxfs, the output has one font, no named styles, no dxfs.
Any reference to `fontId>0`, `fillId>1`, `borderId>0`, or `xfId>0`
points at non-existent indices. Original `xfId 7` may emerge as output
`xfId 2`. **We cannot preserve styling through SheetJS's writer.**

---

## 5. Shared strings

### 5.1 Read

`parse_sst_xml` (`bits/42_sstxml.js:200`) is straightforward: every `<si>`
becomes one entry in the `strs` array. Per `<si>`:

- Plain `<t>` text → `{ t: <text>, r: <raw xml>, h: <html> }`.
- Rich runs `<r>` → `{ t: <concatenated plaintext>, r: <raw xml>, h:
<rich-text HTML> }` via `parse_rs` and `rs_to_html`.

So rich text _is_ preserved — but only as raw XML markup in `cell.r`. The
parser does not give us a structured representation of the runs.

### 5.2 Write — when SST is used vs inline strings

`write_zip_xlsx` calls `write_sst_xml` only if `opts.Strings.length > 0`,
which is populated only when `opts.bookSST` is true. With the default
(`bookSST: false`), `write_ws_xml_cell` (`bits/67_wsxml.js:295`) emits
cells as `t="str"` with `<v>…</v>` — **not** proper
`<is><t>…</t></is>` inlineStr blocks.

Consequences:

- Default writes have no `xl/sharedStrings.xml` and no `<is>` blocks.
- Excel reads this fine, but diff tools and validators expecting SST or
  `inlineStr` will trip.
- **Rich text (`cell.r`) is dropped** on this path — only `cell.v` is
  written.
- We almost certainly want `bookSST: true` on every write.

### 5.3 Inline strings on read

`p.t === "inlineStr"` is unwrapped via `parse_si` and stored as
`t='s'`, with rich-text content preserved on `cell.r`. Read
losslessly, but they become SST or `t="str"` on write — they don't
round-trip as `<is>`.

---

## 6. Format strings

`bits/10_ssf.js` is the embedded "SSF" (SpreadSheet Format) library.

### 6.1 Built-in format IDs

`table_fmt` (`bits/10_ssf.js:70`) hard-codes the 25-or-so OOXML built-in
formats: `0='General'`, `1='0'`, `2='0.00'`, `3='#,##0'`, `9='0%'`,
`14='m/d/yy'`, `22='m/d/yy h:mm'`, `49='@'` (text), and the rest. The
`SSF_default_map` table aliases historical IDs (e.g. format 27 → 14,
format 67 → 9) so workbooks authored in non-en-US locales still resolve
correctly.

### 6.2 Custom format codes

Every `<numFmt formatCode="…" numFmtId="N"/>` becomes
`styles.NumberFmt[N]`. IDs above 0x188 (392) are remapped down to recover
Excel's "reusable upper range" pattern (see `bits/47_styxml.js:321`).

### 6.3 Date format detection

`fmt_is_date(table_fmt[fmtid])` looks for unescaped `m`, `d`, `y`, `h`,
`s`, etc. in the format code. This is the gate for the
`opts.cellDates` cell-type promotion at `bits/67_wsxml.js:482`.

### 6.4 Locale handling

There is essentially **none**. `[$-409]` and similar locale prefixes are
treated as opaque: SSF passes them through but does not interpret them.
Date and number formatting is always in English. `dateNF` lets you
override the default date format on a workbook-wide basis but does not
provide locale support.

This matters for our agent layer: any "format this number as
German-style currency" feature has to be implemented above SheetJS, not
delegated to it.

### 6.5 Format strings on write

`write_numFmts` (`bits/47_styxml.js:290`) only emits formats whose IDs
fall in specific ranges (5-8, 23-26, 41-44, 50+). IDs in the "built-in"
range are assumed to be present in every Excel reader and are not
re-emitted. This is correct per the spec but worth noting if we ever
want to inspect raw output.

---

## 7. Formulas

### 7.1 Evaluation

**SheetJS does not evaluate formulas.** This is explicit throughout the
codebase and in the `README` (the formula-evaluation feature is part of
SheetJS Pro, not the CE). Cached values are preserved verbatim; if the
cached value is missing, the cell is forced to `{t:'n', v:0}` at
`bits/67_wsxml.js:419`.

### 7.2 Formula representation

- `cell.f` is the formula text without the leading `=`.
- On read, `_xlfn.` and `_xlws.` prefixes (used by Excel for "future"
  functions like `UNIQUE`, `XLOOKUP`, etc.) are stripped via the `_xlfn`
  helper unless `opts.xlfn` is true.
- On write, the prefixes need to be reapplied. There is a `fuzzyfmla`
  detector at `bits/61_fcommon.js:53` that classifies formula strings,
  but the reapplication logic is incomplete: simple references survive,
  but complex expressions can lose the `_xlfn.` prefix on round-trip,
  which can break Excel < 2019.
- For our purposes, set `opts.xlfn: true` on read so the prefixes are
  preserved verbatim, and set them ourselves on write.

### 7.3 Array formulas

Read at `bits/67_wsxml.js:396`. The corner cell carries `cell.f` plus
`cell.F = "<range>"`. Every other cell in the range carries `cell.F`
only (no `cell.f`). Written back symmetrically — `write_ws_xml_cell` at
`bits/67_wsxml.js:298` adds `t="array"` and `ref="<range>"` to the `<f>`
tag of the corner cell.

### 7.4 Shared formulas

Read at `bits/67_wsxml.js:399`. SheetJS **un-shares** them: it parses the
master formula, then for every cell that references the shared formula
by `si`, it calls `shift_formula_xlsx` (`bits/61_fcommon.js:46`) to
compute the cell-relative shift and stores the _expanded_ formula on
`cell.f`.

Consequence: round-trip eliminates `<f t="shared" si="…">` constructs.
The output has fully-expanded formulas in every cell. Excel re-shares
on save, but git diffs and downstream tools will see noisy diffs.

### 7.5 Dynamic arrays (Excel 365)

The `cm` attribute on a cell points at an `XLDAPR` block in
`xl/metadata.xml`. SheetJS reads this and sets `cell.D = true`
(`bits/67_wsxml.js:483`). On write, `cell.D` triggers `cm="1"` on the
cell tag, and `xl/metadata.xml` is emitted from a hard-coded XLDAPR
template (`bits/51_xlsxmeta.js:81`). This means dynamic-array semantics
are preserved as a flag, but the concrete metadata isn't round-tripped —
it's regenerated.

For our model: treat `cell.D` as a "this is a spill anchor" boolean.

### 7.6 Defined names

Read into `wb.Workbook.Names` (`bits/72_wbxml.js:84`). Each entry is
`{Name, Ref, Sheet?, Comment?, Hidden?}`. Sheet-scoped names use the
`localSheetId` field. Names round-trip cleanly.

Note: `_xlnm._FilterDatabase` (the implicit name behind autofilter
ranges) is auto-managed by `write_ws_xml_autofilter` at
`bits/67_wsxml.js:228`. We should not touch it directly.

---

## 8. What SheetJS DOES NOT preserve well

This section is the prioritised opaque-blob target list for our
`OoxmlContainer`. For each item, the answer to "should we capture the raw
bytes and pass them through?" is yes.

### 8.1 PivotTables and PivotCaches

`"TODO"` in `bits/30_ctype.js`. No parser, no writer. Affected:
`xl/pivotTables/pivotTable*.xml`, `xl/pivotCache/pivotCacheDefinition*.xml`,
`xl/pivotCache/pivotCacheRecords*.xml`, the workbook `<pivotCaches>`
references, and modern extLst-based enhancements. A workbook round-trips
with the pivot completely removed.

### 8.2 Excel Tables (ListObjects)

`"TODO"`. `xl/tables/table*.xml` and the worksheet `<tableParts>` are
unhandled. A table on write becomes an unframed range — auto-filter,
totals row, table style, sort state all lost. High-priority opaque
preservation target.

### 8.3 Charts

`"TODO"` for chart definitions; only `numCache` cached values are
extracted by `bits/69_chartxml.js`. Series/axes/title/legend/formatting
are all dropped. Embedded charts inside worksheets (via
`xl/drawings/drawing*.xml` → `xl/charts/chart*.xml`) are not resolved.
Affected: `xl/charts/chart*.xml`, `xl/charts/colors*.xml`,
`xl/charts/style*.xml`, `xl/charts/_rels/*.rels`,
`xl/drawings/drawing*.xml` and their `_rels`.

### 8.4 Drawings, images, OLE

- `xl/drawings/*` (other than the chart-resolution path) — dropped.
- `xl/media/*` (PNG/JPEG/EMF/WMF embedded images) — dropped.
- `xl/embeddings/*` (OLE) — dropped.
- `xl/activeX/*`, `xl/ctrlProps/*` — dropped.

### 8.5 / 8.6 Conditional formatting & data validations

`<conditionalFormatting>` and `<dataValidations>` are **silently dropped
on read**: `parse_ws_xml`'s regex extraction only covers `sheetPr`,
`dimension`, `sheetViews`, `cols`, `sheetData`, `autoFilter`,
`mergeCells`, `hyperlinks`, `pageMargins`. Everything else sits in the
post-`sheetData` blob and is never inspected. Writes regenerate the
worksheet entirely, so neither can survive.

### 8.7 Custom XML parts

`customXml/item*.xml`, `customXml/itemProps*.xml`, their `_rels`, and the
workbook → custom-xml rels are all unhandled. Breaks
SharePoint/InfoPath/content-type metadata.

### 8.8 VBA

Preserved as a binary blob via `wb.vbaraw` when `opts.bookVBA` is set.
**The signature (`vbaProjectSignature`) is dropped**, so signed macros
become unsigned. We have to capture the signature ourselves.

### 8.9 Sparklines, slicers, timelines

All `"TODO"`. `xl/slicers/*`, `xl/slicerCaches/*`, `xl/timelines/*`,
`xl/timelineCaches/*`, and the extLst-based sparkline groups in
worksheets are dropped.

### 8.10 External links

`parse_xlink` runs but the result is discarded (the `var externbooks`
line in `bits/85_parsezip.js:121` is a commented assignment — purely for
side effects on the rels graph). `xl/externalLinks/*` is never emitted.
Cached values stay intact but recalculation in Excel will fail.

### 8.11 Calculation chain

Read only when `opts.bookDeps`. Never written. Mostly fine — Excel
rebuilds it on first recalc — but some validators warn on its absence.

### 8.12 Connections, query tables, data model

All `"TODO"`. PowerQuery tables and connections (`xl/queryTables/*`,
`xl/connections.xml`, `xl/model/*`) are completely unhandled.

### 8.13 Theme

`theme1.xml` parsed only when `opts.cellStyles` is set, and only the
colour scheme is consumed. Font scheme, format scheme, and `theme2+.xml`
are dropped. On write a default theme is emitted — custom corporate
themes are overwritten.

### 8.14 Sheet protection details

`<sheetProtection>` algorithm/hashValue/saltValue/spinCount are TODO
(comment at `bits/67_wsxml.js:146`). Only the legacy 16-bit
`crypto_CreatePasswordVerifier_Method1` hash is computed on write.

### 8.15 Sheet-level details dropped by `parse_ws_xml`

Only `sheetPr/codeName` + `outlinePr`, `dimension`,
`sheetViews/zoomScale` + `RTL`, `cols/width|hidden|outlineLevel` (with
`cellStyles`), `sheetData`, `autoFilter/ref`, `mergeCells`,
`hyperlinks`, and `pageMargins` are preserved. Page setup,
headers/footers, breaks, custom views, scenarios, ignored errors,
controls, OLE objects, smart tags, extLst — all silently dropped.

---

## 9. Round-trip fidelity

There is no formal round-trip guarantee in SheetJS CE. The README and
docs explicitly position the library as a _data extraction and
generation_ tool, not a fidelity-preserving editor. The following is
what's known empirically.

### 9.1 Bytes are never identical

Even for a freshly Excel-authored XLSX with two cells, SheetJS's output
differs:

- Different element ordering in worksheet/workbook XML.
- Theme regenerated to a default template.
- Styles table rebuilt to the minimal scaffold (§4.1).
- Inline `t="str"` instead of SST entries (§5.2) by default.
- Shared formulas expanded (§7.4).
- Hardcoded XLDAPR block in metadata regardless of original.
- `[Content_Types].xml` re-ordered with a fixed default-extensions list
  (`bits/30_ctype.js:264`).
- Relationship IDs renumbered.
- `docProps/app.xml` fully regenerated; SheetNames overwritten;
  Manager/Company dropped if not in `wb.Props`.
- ZIP central-directory ordering and DEFLATE level differ from Excel's.

### 9.2 Where SheetJS documents deviations

It mostly doesn't, in source. Signals: `// TODO:` comments in `bits/`,
`"TODO"` strings in `bits/30_ctype.js`, and `opts.WTF` (throws on
unrecognised tags — the developer's introspection tool).

For OfficeAI: treat **no part of `xl/` as round-trippable through SheetJS**.
Cells, formulas, and basic styles are usable as data; byte-level structure
is not.

---

## 10. API surface we will use

### 10.1 Read

```ts
XLSX.read(buf, {
  type: "array", // we'll always be passing Uint8Array
  cellFormula: true, // populate cell.f
  cellStyles: true, // populate cell.s; implies cellNF + sheetStubs
  cellNF: true, // populate cell.z (number-format string)
  cellDates: false, // keep dates as serial numbers (stable internal repr)
  cellHTML: false, // skip HTML synthesis (we'll render ourselves)
  cellText: true, // populate cell.w (formatted display text)
  sheetStubs: true, // emit empty cells; needed for round-trip integrity
  dense: true, // 2D-array storage; less GC pressure, easier mapping
  bookFiles: true, // expose raw zip file map (workbook.files) for the opaque-blob layer
  bookVBA: true, // preserve xl/vbaProject.bin
  bookDeps: false, // we'll regenerate calc chain ourselves
  bookSheets: false,
  bookProps: false,
  xlfn: true, // preserve _xlfn. prefixes; we will manage them
  WTF: false, // turn on only in dev/test for surfacing dropped parts
});
```

Why each flag:

- `cellFormula: true` — we need the formula text on `cell.f` for our
  formula engine and for opaque preservation.
- `cellStyles: true` — we need `cell.s` populated even though SheetJS's
  style coverage is partial; combined with our own raw `xl/styles.xml`
  capture we can rebuild the table.
- `cellNF: true` — we need the format code on `cell.z` to resolve type
  promotion and rendering.
- `cellDates: false` — we want a stable internal representation. Date
  serials never have timezone drift; JS `Date` does.
- `sheetStubs: true` — without this, empty cells with style or
  formatting are dropped, breaking layout.
- `dense: true` — SheetJS wraps `ws[r][c]` arrays only when the option is
  set (`bits/67_wsxml.js:23`). Dense form maps cleanly to our row-major
  internal storage.
- `bookFiles: true` — the only way to extract the raw zip entries from
  SheetJS's perspective. Combined with our own ZIP read pass, this gives
  us a snapshot of every part for the opaque-blob layer.
- `bookVBA: true` — preserves the binary VBA project. We still need to
  capture `vbaProjectSignature` ourselves.
- `xlfn: true` — keeps `_xlfn.` prefixes verbatim so we don't have to
  guess where to reapply them on write.
- `WTF: false` in production, `true` in our test suite — SheetJS will
  throw on every unrecognised tag, which is exactly the signal we need
  to drive our opaque-blob coverage tests.

### 10.2 Write

We will _not_ call `XLSX.write` for the primary write path because we
need to preserve our opaque blobs (which SheetJS will overwrite). Use
SheetJS's writer only for the data-bearing parts that we then merge into
our own ZIP (see §11).

When we do call it (e.g. for a "create from scratch" path):

```ts
XLSX.writeXLSX(wb, {
  type: "array", // returns Uint8Array
  bookSST: true, // emit a real xl/sharedStrings.xml; preserves rich text on round-trip
  compression: true, // DEFLATE; closer to Excel's defaults
  cellDates: false, // emit dates as serials (matches what most tools expect)
  WTF: false,
});
```

`writeXLSX` (vs `write`) is the lighter entry point that bypasses the
`bookType` switch and goes straight to `write_zip_xlsx`. Use it
unconditionally for our XLSX target — never write XLSB through SheetJS
because the dedupe and metadata behaviours diverge.

### 10.3 Utilities we will use

From `bits/95_api.js`:

- `XLSX.utils.encode_cell` / `decode_cell` / `encode_range` / `decode_range`
  — addressing helpers, very stable.
- `XLSX.utils.encode_col` / `decode_col` / `encode_row` / `decode_row`
  — same.
- `XLSX.utils.format_cell(cell, v?, opts?)` — runs the SSF formatter on a
  cell for display. Useful for our renderer.
- `XLSX.utils.aoa_to_sheet`, `XLSX.utils.json_to_sheet`,
  `XLSX.utils.sheet_to_json` — for our agent commands that ingest or
  emit tabular data, these are the ergonomic boundary.
- `XLSX.SSF.format(fmt, value)` — apply a number format outside of a
  cell context.

Avoid:

- `XLSX.utils.book_new` / `book_append_sheet` — they're fine for
  "create from scratch" but don't give us hooks for opaque preservation.
  We'll wrap them in our own `XlsxWorkbook` factory.
- `XLSX.utils.sheet_to_html` — uses SheetJS's own table model, not ours.
  We'll write our own renderer.
- `cell_set_hyperlink` / `cell_add_comment` — these mutate cells in
  place. Use them only via our command bus so the model stays
  authoritative.

---

## 11. Where we layer our own logic

The high-level pattern for `packages/xlsx/src/parser/`:

1. **Open the zip ourselves** (with `fflate` or `jszip` — the same
   dependency the rest of OfficeAI uses for DOCX). Walk every entry,
   compute SHA-256 per part, and stash the raw bytes in
   `OoxmlContainer.parts` keyed by full path.

2. **Hand the same buffer to `XLSX.read`** with the option set in
   §10.1. Now we have two views of the same workbook:
   - `OoxmlContainer.parts`: ground-truth bytes, every part.
   - SheetJS `WorkBook`: parsed cells, formulas, partial styles.

3. **Translate SheetJS's `WorkBook` → our `XlsxWorkbook` model.** Per
   sheet: walk the dense 2D array; map `{t,v,f,F,z,s,l,c,r,D,w}` → our
   `XlsxCell`; preserve `cell.r` rich-text XML verbatim; lift sheet
   metadata (ref, merges, autofilter, hyperlinks, margins, cols).
   Workbook level: copy `wb.Workbook.Names`, `wb.Props`, `wb.Custprops`.

4. **Augment with our own parsers for things SheetJS drops:**
   - `<conditionalFormatting>`, `<dataValidations>`, `<sheetProtection>`
     parsed from the raw worksheet XML in `OoxmlContainer.parts`.
   - `xl/tables/table*.xml` → first-class `XlsxTable`.
   - `xl/styles.xml`: parse `cellStyleXfs`, `cellStyles`, `dxfs`, full
     borders, gradient fills.
   - `xl/theme/*.xml`: full theme, including non-default variants.
   - `xl/calcChain.xml`: optional; we can rebuild.

5. **Mark the rest as opaque.** Anything in `OoxmlContainer.parts` not
   claimed above stays as a `Uint8Array` with content-type and
   rels-graph neighbours intact. Known opaque set:
   `xl/pivotTables/*`, `xl/pivotCache/*`, `xl/charts/*`,
   `xl/drawings/*`, `xl/media/*`, `xl/embeddings/*`, `xl/customXml/*`,
   `customXml/*`, `docProps/customXml*`, `xl/queryTables/*`,
   `xl/connections.xml`, `xl/model/*`, `xl/slicers/*`,
   `xl/slicerCaches/*`, `xl/timelines/*`, `xl/timelineCaches/*`,
   `xl/printerSettings/*`, `xl/activeX/*`, `xl/ctrlProps/*`,
   `xl/externalLinks/*`, `xl/vbaProjectSignature.bin`, themes other
   than `theme1.xml`, and any `xl/*` part with an unrecognised content
   type.

### 11.1 Write path

The serializer is `OoxmlContainer.serialize()`-driven, not
`XLSX.write`-driven:

1. For each modified sheet: build a SheetJS-shaped `WorkBook` with only
   that sheet, call `XLSX.writeXLSX` (or our own adapter over
   `write_ws_xml`), replace the entry in `OoxmlContainer.parts`.
2. For unmodified sheets: reuse original bytes verbatim.
3. For our first-class parts (styles, tables, conditional formats):
   serialize ourselves.
4. Rebuild `[Content_Types].xml` and rels graphs from
   `OoxmlContainer`'s tracked content types and rels.
5. Re-zip with deterministic ordering and DEFLATE level matching
   Excel's defaults to minimise diffs.

### 11.2 What we explicitly delegate to SheetJS

- Cell-value parsing (SST resolution, inline strings, error codes,
  date serials).
- Formula text extraction and shared-formula expansion.
- Number-format formatting via SSF (via `XLSX.utils.format_cell`).
- Address arithmetic via `XLSX.utils.encode_cell` etc.
- Hyperlink resolution against the worksheet rels graph.

### 11.3 What we explicitly do NOT delegate

- The ZIP layer (we own it).
- `[Content_Types].xml` and `_rels/*.rels` (we own them; SheetJS would
  overwrite).
- `xl/styles.xml` (we own it; SheetJS would emit a minimal scaffold).
- `xl/theme/*.xml` (we own it; SheetJS would emit a default).
- Any of the §1.2 / §8 "TODO" parts (we own them as opaque blobs).
- Conditional formatting, data validations, sheet protection details,
  sheet view state, page setup, header/footer, breaks, scenarios,
  controls, custom views, ignored errors (we parse and own them; SheetJS
  drops them).
- Tables / ListObjects (we own them as a first-class model type).
- Defined-name management for `_xlnm._FilterDatabase` and friends — we
  have to keep both SheetJS and our model in sync because SheetJS
  auto-edits them.

---

## 12. Five-sentence summary

SheetJS CE is a competent **data layer** for XLSX: it reads cell values,
formulas, hyperlinks, comments, and a partial slice of styles into a
plain-object workbook model, and writes them back as a structurally
correct (if minimal) `xl/*` tree. It does _not_ preserve, on either
read or write, the OOXML parts that make up the bulk of a real Excel
workbook's surface area: tables, charts, drawings, images, pivot
tables, conditional formatting, data validations, slicers, timelines,
custom XML, themes beyond `theme1`, full styles (cellStyleXfs / dxfs /
named cell styles / borders / gradient fills), connections, query
tables, external links, sheet-protection algorithms, and most
worksheet-level layout state — all of which it either silently drops
or treats as `"TODO"` content types. For OfficeAI we will use SheetJS
behind a thin adapter as the cell/formula/SSF engine via
`XLSX.read(buf, {dense: true, cellFormula: true, cellStyles: true,
cellNF: true, sheetStubs: true, bookFiles: true, bookVBA: true,
xlfn: true, cellDates: false})`, but **we own the ZIP** and we own the
serializer. Every OOXML part we don't first-class in our model will be
captured as a `Uint8Array` in `OoxmlContainer` on read and re-emitted
verbatim on write, and the only sheets we re-serialize through SheetJS
are the ones we've actually modified — giving us byte-stable round-trips
for untouched sheets and intentional, scoped diffs for the ones the
agent edits.
