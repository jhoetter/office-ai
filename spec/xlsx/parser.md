# XLSX — Parser

> Bytes (`.xlsx`) → `XlsxSnapshot` with **opaque-blob preservation** for every
> OOXML part the model does not first-class. The parser is the read half of
> the byte-preservation contract — every part it does not deserialize gets
> stashed verbatim in `OoxmlContainer.parts` so the serializer can re-emit
> it bit-identical.

This document mirrors the SHAPE of [`spec/docx/parser.md`](../docx/parser.md):
high-level goal, two-stage pipeline, per-feature parsing strategies,
pseudocode, error model, and budget. It is the implementation contract for
`packages/xlsx/src/parser/`.

References:

- [`spec/shared/ooxml-utils.md`](../shared/ooxml-utils.md) — `OoxmlContainer`,
  `parseXml`/`serializeXml`, namespace handling.
- [`spec/xlsx/feature-scope.md`](feature-scope.md) — what we model vs what we
  preserve as opaque.
- [`spec/xlsx/edge-cases.md`](edge-cases.md) — encoding, namespaces,
  shared/array formulas, sparse cells, sheet rename.
- [`spec/xlsx/analysis-sheetjs.md`](analysis-sheetjs.md) §10 — the read-side
  options for `XLSX.read`.
- [`spec/xlsx/analysis.md`](analysis.md) §1.1–§1.3 — KEEP / DIFFER / IMPROVE.

---

## 1. Goal

Given an `ArrayBuffer | Uint8Array` containing a real-world `.xlsx`, produce
an `XlsxSnapshot`:

```typescript
interface XlsxSnapshot {
  readonly format: "xlsx";
  readonly revision: 0;
  readonly root: XlsxWorkbook;
  /** SHA-256 hex per zip entry path; the byte-preservation oracle. */
  readonly partHashes: Record<string, string>;
  /** The container is attached to the snapshot for the serializer. */
  readonly container: OoxmlContainer;
  readonly dirty: XlsxDirtyFlags;
}
```

Properties the parser must satisfy:

1. **Total**: any `.xlsx` produced by Excel 2010–365 (desktop and web),
   LibreOffice Calc, or Google Sheets export parses without throwing —
   provided the zip is intact and `xl/workbook.xml` exists. Unrecognized
   sheet sub-elements, sheet types we don't model (`chartsheet`,
   `dialogsheet`), and unknown content types are **not** errors; they
   degrade to opaque blobs.
2. **Faithful**: every part in the input zip ends up in
   `OoxmlContainer.parts` with byte-identical bytes. Hashes are computed at
   parse time and recorded in `snapshot.partHashes`. The serializer's
   "untouched parts byte-equal" invariant is verified against this map.
3. **Idempotent at the model boundary**: parsing the same file twice
   produces structurally equal `XlsxWorkbook`s (modulo `nodeId` minting,
   which uses an injectable `IdMinter` in tests).
4. **Fast**: the 50k-row × 10-col synthetic fixture parses in **< 1.5 s**
   on the CI box (Node 22, 2-core ARM). See §10.

---

## 2. Two-stage pipeline

The parser layers three independent passes over the same input buffer.
Each pass produces a different view; the final reconciliation step merges
them into the model.

```
                    ┌─────────────────────────────────────────────┐
                    │  Stage A: OoxmlContainer.fromBuffer(buf)    │
   ArrayBuffer ────►│  - unzip every entry                        │
                    │  - populate parts: Map<string, Uint8Array>  │
                    │  - hash each part → partHashes              │
                    └────────────────────┬────────────────────────┘
                                         │
                         ┌───────────────┴────────────────┐
                         ▼                                ▼
       ┌──────────────────────────────┐   ┌──────────────────────────────┐
       │ Stage B: SheetJS adapter     │   │ Stage C: Native XML pass     │
       │ XLSX.read(buf, {dense, ...}) │   │ parseXml on subset of parts: │
       │ - cells, formulas, SST       │   │   sharedStrings, styles,     │
       │ - merges, hyperlinks,        │   │   comments*, *.rels,         │
       │   autofilter range,          │   │   conditionalFormatting,     │
       │   number-format codes        │   │   defined names,             │
       │ - dynamic-array flags (D)    │   │   sheet-level CF / DV blocks │
       └──────────────┬───────────────┘   └──────────────┬───────────────┘
                      │                                  │
                      └────────────────┬─────────────────┘
                                       ▼
                ┌────────────────────────────────────────────┐
                │ Stage D: Reconciliation                    │
                │ - merge SheetJS cells with native          │
                │   annotations (CF rules, comments,         │
                │   hyperlinks, defined-name scopes)         │
                │ - build XlsxWorkbook                       │
                │ - record partHashes                        │
                └────────────────────────────────────────────┘
                                       ▼
                                 XlsxSnapshot
```

### Stage A — `OoxmlContainer`

```
container = await OoxmlContainer.load(buf)
// container.parts now has every zip entry as Uint8Array, in zip order.
// container.allHashes() is the partHashes oracle.
```

This is the same code DOCX uses; see
[`packages/core/src/ooxml/container.ts`](../../packages/core/src/ooxml/container.ts).
The container stores bytes in load order so re-emit is stable.

If the zip itself fails to load → `XlsxParseError("zip-corruption", ...)`
(see §7).

### Stage B — SheetJS adapter pass

We hand the **same buffer** (not the container's part bytes) to SheetJS.
Why the original buffer: SheetJS opens the zip itself and walks every
entry to populate its workbook object; passing it the buffer avoids a
round-trip through our container. The cost is one redundant unzip — at our
sizes this is < 50 ms even for a 50k-row sheet, and the benefits (keeping
the SheetJS adapter dependency-free of our container, letting SheetJS use
its own streaming where it has it) are worth it.

```typescript
const sjsWorkbook = XLSX.read(buf, {
  type: "array",
  dense: true, // 2D-array storage; row-major maps to our model
  cellFormula: true, // populate cell.f
  cellStyles: true, // populate cell.s; implies cellNF + sheetStubs
  cellNF: true, // populate cell.z (number-format string)
  sheetStubs: true, // emit empty cells with formatting
  cellDates: false, // keep dates as serial numbers (stable repr)
  cellText: true, // populate cell.w (formatted display text)
  cellHTML: false, // we render ourselves
  bookFiles: true, // expose raw zip entries (we use container instead,
  // but having both is cheap and aids debugging)
  bookVBA: true, // preserve xl/vbaProject.bin metadata flag
  bookDeps: false, // we rebuild calc chain ourselves
  bookSheets: false,
  bookProps: false,
  xlfn: true, // preserve _xlfn. prefixes verbatim
  WTF: false, // production: false. Tests: true (surfaces drops)
});
```

Each option's rationale is documented in
[`spec/xlsx/analysis-sheetjs.md`](analysis-sheetjs.md) §10.1. The TL;DR is
"give us the maximum possible cell/formula/format detail while keeping
internal representations stable; we own everything outside the cell layer."

### Stage C — Native XML pass

SheetJS silently drops large swaths of the worksheet body and ignores
several whole parts (see [analysis-sheetjs §1, §8](analysis-sheetjs.md)).
We re-parse a small targeted subset of the container ourselves with
`parseXml` from `@officeai/core/ooxml` (which is `fast-xml-parser` with
`preserveOrder: true` — see [`spec/shared/ooxml-utils.md`](../shared/ooxml-utils.md)).

The native subset:

| Part path                                    | Why we re-parse                                            |
| -------------------------------------------- | ---------------------------------------------------------- |
| `xl/_rels/workbook.xml.rels`                 | Authoritative sheet → file map; defined-name external refs |
| `xl/worksheets/_rels/sheet{N}.xml.rels`      | Hyperlink targets, comments rels, drawings rels            |
| `xl/sharedStrings.xml`                       | Order-of-first-occurrence preserved for stable diffs       |
| `xl/styles.xml`                              | Full table (`cellStyleXfs`, `dxfs`, named styles, borders) |
| `xl/comments{N}.xml`                         | Author + text + body run formatting + cell anchor          |
| `xl/threadedComments/threadedComment{N}.xml` | Thread structure (parent/child, resolved)                  |
| `xl/persons/person.xml`                      | Author display names for threaded comments                 |
| `xl/worksheets/sheet{N}.xml`                 | `<conditionalFormatting>`, `<dataValidations>`,            |
|                                              | `<hyperlinks>`, `<sheetProtection>`, `<sheetView>`         |

Everything else stays in `container.parts` as opaque bytes:
`xl/pivotTables/*`, `xl/pivotCache/*`, `xl/charts/*`, `xl/drawings/*`,
`xl/media/*`, `xl/embeddings/*`, `xl/customXml/*`, `customXml/*`,
`xl/queryTables/*`, `xl/connections.xml`, `xl/model/*`, `xl/slicers/*`,
`xl/slicerCaches/*`, `xl/timelines/*`, `xl/timelineCaches/*`,
`xl/printerSettings/*`, `xl/activeX/*`, `xl/ctrlProps/*`,
`xl/externalLinks/*`, `xl/vbaProject.bin`,
`xl/vbaProjectSignature.bin`, `xl/theme/theme{N}.xml` for `N>1`,
`xl/calcChain.xml`, `xl/metadata.xml`, and any part whose content type
is not in our recognised set.

### Stage D — Reconciliation

Merge SheetJS's cell layer with our native annotations into the final
`XlsxWorkbook`. The merge is **left-biased on cells** (SheetJS wins for
`{t,v,f,F,z,s.numFmtId,l,c,D}`) and **right-biased on annotations** (our
native pass wins for `{conditionalFormats, dataValidations, comments
body run formatting, defined-name scoping, hyperlinks targets,
sheetView state, sheetProtection details}`).

See `XlsxWorkbook` in
[`spec/xlsx/document-model.md`](document-model.md) (companion doc) for the
full target shape.

---

## 3. Entry point

```
parseXlsx(buf, opts?): Promise<XlsxSnapshot>:
  // Stage A: zip + part-hash oracle
  try:
    container = await OoxmlContainer.load(buf)
  catch err:
    throw new XlsxParseError("zip-corruption", "Failed to read XLSX as zip", { cause: err })

  if !container.has("xl/workbook.xml"):
    throw new XlsxParseError("missing-required-part",
                             "Missing xl/workbook.xml",
                             { path: "xl/workbook.xml" })

  // Detect & decode encoding per part (EC-I4 in edge-cases.md)
  // Most parts are UTF-8; we honour <?xml encoding="..."?> declarations
  // and fall back to UTF-8 with replacement on decode failure.
  // The container does this transparently in readText().

  // Stage B: SheetJS adapter
  let sjsWorkbook
  try:
    sjsWorkbook = XLSX.read(buf, SJS_READ_OPTS)
  catch err:
    throw new XlsxParseError("invalid-xml",
                             "SheetJS failed to parse workbook",
                             { cause: err })

  // Stage C: native XML pass
  workbookMeta   = parseWorkbookXml(container)         // sheets, definedNames, props
  workbookRels   = parseRels(container, "xl/_rels/workbook.xml.rels")
  sst            = parseSharedStrings(container)        // optional
  styles         = parseStyles(container)               // full styles incl. dxfs
  commentParts   = parseAllCommentsParts(container)    // by sheet idx
  threadedComments = parseAllThreadedComments(container)
  people         = parsePeople(container)               // optional
  sheetXmlByIdx  = parseSheetXmls(container, workbookMeta, workbookRels)
                   // for each sheet: { conditionalFormatting, dataValidations,
                   //   hyperlinks, sheetProtection, sheetView, autoFilter, mergeCells }

  // Stage D: reconciliation per sheet
  sheets = []
  for (idx, meta) in workbookMeta.sheets:
    sjsSheet  = sjsWorkbook.Sheets[meta.name]
    nativeSheet = sheetXmlByIdx[idx]
    sheets.push(reconcileSheet(meta, sjsSheet, nativeSheet,
                               styles, sst,
                               commentParts[idx], threadedComments[idx], people))

  workbook = {
    sheets,
    definedNames: workbookMeta.definedNames,
    styles,
    sst,
    workbookView: workbookMeta.workbookView,
    workbookProtection: workbookMeta.workbookProtection,
    props: workbookMeta.props,
    customProps: workbookMeta.customProps,
    workbookRootAttrs: workbookMeta.rootAttrs,
  }

  partHashes = container.allHashes()
  dirty = freshDirtyFlags()

  return { format: "xlsx", revision: 0, root: workbook,
           partHashes, container, dirty }
```

The shape of `XlsxDirtyFlags` is documented in
[`spec/xlsx/document-model.md`](document-model.md), but for the parser
the important property is: **all flags are `false` after a fresh parse**.
This is what makes `serialize(parse(buf))` byte-equivalent for every part.

---

## 4. Sheet parsing

```
reconcileSheet(meta, sjsSheet, nativeSheet, styles, sst, comments,
               threaded, people) -> Sheet:

  cells = parseCells(sjsSheet, styles, sst)
  merges = parseMerges(sjsSheet)              // from ws['!merges']
  hyperlinks = nativeSheet.hyperlinks         // we own resolution
  conditionalFormats = nativeSheet.conditionalFormatting   // typed model
  dataValidations = nativeSheet.dataValidations            // typed model
  sheetProtection = nativeSheet.sheetProtection            // typed model
  sheetView = nativeSheet.sheetView                        // freeze, zoom, RTL
  autoFilter = nativeSheet.autoFilter                      // range only in P0
  cols = nativeSheet.cols                                  // widths, hidden, outlineLevel
  rows = nativeSheet.rows                                  // heights, hidden, outlineLevel

  return Sheet {
    id: meta.sheetId,
    name: meta.name,
    visibility: meta.state,                  // "visible" | "hidden" | "veryHidden"
    tabColor: nativeSheet.sheetView.tabColor,
    cells,                                   // Map<string, Cell> keyed by `${row}:${col}` (0-based)
    merges,                                  // IndexedRanges<Merge>
    conditionalFormats,
    dataValidations,
    hyperlinks: indexHyperlinksByCell(hyperlinks),
    comments: indexCommentsByCell(comments, threaded, people),
    sheetProtection,
    frozenRows: sheetView.frozen?.rows ?? 0,
    frozenCols: sheetView.frozen?.cols ?? 0,
    autoFilter,
    cols,
    rows,
    sheetRootAttrs: nativeSheet.rootAttrs,
  }
```

### 4.1 Cells

```
parseCells(sjsSheet, styles, sst) -> Map<string, Cell>:
  out = new Map()
  if !sjsSheet || !sjsSheet['!ref']: return out
  const range = decode_range(sjsSheet['!ref'])
  for r in [range.s.r .. range.e.r]:
    const row = sjsSheet[r]                  // dense: row is an array
    if !row: continue
    for c in [range.s.c .. range.e.c]:
      const sjs = row[c]
      if !sjs || sjs.t === undefined: continue
      out.set(`${r}:${c}`, sjsCellToCell(sjs, styles, sst))
  return out

sjsCellToCell(sjs, styles, sst) -> Cell:
  return Cell {
    value: sjsValueToValue(sjs),             // {kind:"string"|"number"|"bool"|"error"|"blank", ...}
    formula: sjs.f
              ? { text: sjs.f,
                  arrayRef: sjs.F,           // if part of array formula
                  cachedValue: sjs.v,
                  shared: undefined }        // shared flag rebuilt at serialize
              : undefined,
    styleId: sjs.s ? styles.idForXf(sjs.s) : undefined,
    numFmt: sjs.z,                           // format code string
    richText: sjs.r,                         // raw rich-text XML (preserved verbatim)
    dynamicArrayAnchor: sjs.D === true,
  }
```

`sjsValueToValue` uses an exhaustive switch on `sjs.t`:

| `t`     | `Cell.value`                                               |
| ------- | ---------------------------------------------------------- |
| `'n'`   | `{ kind: "number", value: sjs.v }`                         |
| `'s'`   | `{ kind: "string", value: sjs.v }`                         |
| `'b'`   | `{ kind: "bool", value: sjs.v }`                           |
| `'e'`   | `{ kind: "error", code: errorCodeFromExcel(sjs.v) }`       |
| `'d'`   | `{ kind: "number", value: dateToSerial(sjs.v) }`           |
| `'z'`   | `{ kind: "blank" }`                                        |
| `'str'` | `{ kind: "string", value: sjs.v }` (formula string result) |

Errors are normalised to our textual code set (`"#REF!"`, `"#NAME?"`,
`"#DIV/0!"`, `"#VALUE!"`, `"#NULL!"`, `"#NUM!"`, `"#N/A"`) — we do not
carry SheetJS's numeric Excel error codes through the model.

### 4.2 Merges

```
parseMerges(sjsSheet) -> IndexedRanges<Merge>:
  const out = new IndexedRanges()
  for m in sjsSheet['!merges'] ?? []:
    out.add({ top: m.s.r, left: m.s.c, bottom: m.e.r, right: m.e.c })
  return out
```

Merges are an `IndexedRanges<Merge>` so insert/delete row+column shift in
the same one-place pattern as Univer's `RefRangeService` — see
[`spec/xlsx/analysis.md`](analysis.md) §1.1.

### 4.3 Conditional formatting

SheetJS drops conditional formats entirely (analysis-sheetjs §8.5/§8.6).
We parse them ourselves from the sheet XML.

```
parseConditionalFormatting(sheetTree) -> ConditionalFormat[]:
  out = []
  for cfBlock in findAllChildren(sheetTree, "conditionalFormatting"):
    const sqref = attr(cfBlock, "sqref")           // e.g. "A1:A100 C1:C100"
    const ranges = parseSqref(sqref)
    for cfRule in childrenWithTag(cfBlock, "cfRule"):
      const type = attr(cfRule, "type")
      const priority = parseInt(attr(cfRule, "priority"), 10)
      const dxfId = attr(cfRule, "dxfId")          // resolved against styles.dxfs

      switch type:
        case "cellIs":
          out.push({ kind: "cell-is", op: attr(cfRule, "operator"),
                     formulas: collectFormulas(cfRule),
                     ranges, dxfId, priority })
        case "containsText":
          out.push({ kind: "contains-text",
                     text: attr(cfRule, "text"),
                     ranges, dxfId, priority })
        case "timePeriod":
          out.push({ kind: "date-occurring",
                     period: attr(cfRule, "timePeriod"),
                     ranges, dxfId, priority })
        case "colorScale":
          out.push({ kind: "color-scale",
                     stops: parseColorScale(cfRule),
                     ranges, priority })
        case "dataBar":
          out.push({ kind: "data-bar",
                     stops: parseDataBar(cfRule),
                     ranges, priority })
        default:
          out.push({ kind: "opaque-cf", raw: capture(cfRule), ranges, priority })
  return out
```

Per [`feature-scope.md`](feature-scope.md): the agent only **authors**
`cell-is` rules in P0; everything else round-trips. The opaque-cf branch
preserves rules we don't model (icon sets, top10, expression-based, etc.).

### 4.4 Hyperlinks

```
parseHyperlinks(sheetTree, sheetRels) -> Hyperlink[]:
  out = []
  for h in findAllChildren(sheetTree, "hyperlinks/hyperlink"):
    const ref = attr(h, "ref")               // "A1" or "A1:B2"
    const rId = attr(h, "r:id")
    const location = attr(h, "location")     // internal anchor
    const tooltip = attr(h, "tooltip")
    const display = attr(h, "display")
    const target = rId ? sheetRels.byId(rId)?.target : undefined
    out.push({ ref: parseSqref(ref), target, location, tooltip, display, rId })
  return out
```

Indexed by anchor cell at reconciliation; multi-cell hyperlinks expand to
one entry per cell.

### 4.5 Comments

```
parseComments(commentsTree, threadedTree, people) -> Comment[]:
  authors = parseAuthorsList(commentsTree)         // legacy authors[]
  out = []
  for cmt in findAllChildren(commentsTree, "commentList/comment"):
    const ref = attr(cmt, "ref")
    const authorId = parseInt(attr(cmt, "authorId"), 10)
    const author = authors[authorId] ?? ""
    const body = parseRichText(findChild(cmt, "text"))    // run-formatted
    out.push({
      ref, author, text: body.text, runs: body.runs,
      threadedId: undefined,                  // populated below if matched
      resolved: false,
    })

  // Layer threaded-comments on top: a threaded comment in
  // xl/threadedComments/threadedCommentN.xml carries person refs that we
  // resolve against xl/persons/person.xml. The legacy comment is the
  // "anchor" (Excel writes both for back-compat).
  for t in findAllChildren(threadedTree, "threadedComment"):
    const ref = attr(t, "ref")
    const id = attr(t, "id")
    const personId = attr(t, "personId")
    const parentId = attr(t, "parentId")
    const dT = attr(t, "dT")                  // ISO timestamp
    const done = attr(t, "done") === "1"
    const author = people.get(personId)?.displayName ?? ""
    const text = collectThreadedText(t)

    matchAndAttach(out, ref, { id, parentId, author, text, dT, done })
  return out
```

Threaded comments preserve their thread graph through reconciliation.
Adding a new comment via `xlsx:add-comment` writes both the legacy entry
and the threaded entry — see
[`spec/xlsx/serializer.md`](serializer.md) §6.

---

## 5. Workbook-level parsing

```
parseWorkbookXml(container) -> WorkbookMeta:
  tree = parseXml(container.readText("xl/workbook.xml"))
  root = findChild(tree, "workbook")
  rootAttrs = readAllAttrs(root)             // namespace declarations
  sheets = []
  for s in findAllChildren(root, "sheets/sheet"):
    sheets.push({
      name: attr(s, "name"),
      sheetId: attr(s, "sheetId"),
      rId: attr(s, "r:id"),
      state: attr(s, "state") ?? "visible",
    })
  definedNames = []
  for n in findAllChildren(root, "definedNames/definedName"):
    definedNames.push({
      name: attr(n, "name"),
      ref: getTextContent(n),
      localSheetId: attr(n, "localSheetId"),
      hidden: attr(n, "hidden") === "1",
      comment: attr(n, "comment"),
    })
  workbookView = parseWorkbookView(findChild(root, "bookViews/workbookView"))
  workbookProtection = parseWorkbookProtection(findChild(root, "workbookProtection"))
  return { sheets, definedNames, workbookView, workbookProtection,
           props: parseDocProps(container), customProps: parseCustomProps(container),
           rootAttrs }
```

Defined names are preserved in P0 (we read them; the agent doesn't
edit them yet — see `feature-scope.md` "Defined names").

---

## 6. Styles + shared strings

### 6.1 Styles

`parseStyles(container)` is our own implementation, **not** SheetJS's. We
re-parse `xl/styles.xml` to capture the full table — including the parts
SheetJS drops (analysis-sheetjs §4):

```
parseStyles(container) -> Styles:
  if !container.has("xl/styles.xml"):
    return emptyStyles()
  tree = parseXml(container.readText("xl/styles.xml"))
  root = findChild(tree, "styleSheet")
  return Styles {
    rootAttrs: readAllAttrs(root),
    numFmts:    parseNumFmts(findChild(root, "numFmts")),
    fonts:      parseFonts(findChild(root, "fonts")),
    fills:      parseFills(findChild(root, "fills")),
    borders:    parseBorders(findChild(root, "borders")),
    cellStyleXfs: parseCellXfs(findChild(root, "cellStyleXfs")),
    cellXfs:    parseCellXfs(findChild(root, "cellXfs")),
    cellStyles: parseCellStyles(findChild(root, "cellStyles")),
    dxfs:       parseDxfs(findChild(root, "dxfs")),       // CF + table styles use these
    tableStyles: parseTableStyles(findChild(root, "tableStyles")),
    colors:     parseColors(findChild(root, "colors")),
    extLst:     captureOpaque(findChild(root, "extLst")),
  }
```

`Styles` is the **typed** model on the read side. We expose a
content-hashed style id at the model boundary so collaborative edits to
the same xfId from different agents converge — see
[`spec/xlsx/analysis.md`](analysis.md) §1.2 row "Style storage".

### 6.2 Shared strings

```
parseSharedStrings(container) -> SharedStringTable:
  if !container.has("xl/sharedStrings.xml"):
    return emptyTable()
  tree = parseXml(container.readText("xl/sharedStrings.xml"))
  root = findChild(tree, "sst")
  uniqueCount = parseInt(attr(root, "uniqueCount") ?? "0", 10)
  count = parseInt(attr(root, "count") ?? "0", 10)
  entries = []
  for si in childrenWithTag(root, "si"):
    entries.push(parseSharedStringItem(si))   // { text, runs?, rawXml }
  return { uniqueCount, count, entries, rootAttrs: readAllAttrs(root) }
```

We store the raw `<si>` XML on each entry so the serializer can re-emit
rich-text markup byte-for-byte for entries we did not author. New entries
(from agent edits) get rebuilt from typed runs.

The order-of-first-occurrence is preserved on parse; the serializer
appends new entries to the end so existing references (`<c><v>N</v></c>`
where N is the SST index) stay valid for untouched cells.

---

## 7. Errors

The parser is **strict** in the same way DOCX is — every failure is a
structured `XlsxParseError`, never silent.

```typescript
export type XlsxParseErrorKind =
  | "zip-corruption"        // EC-I1
  | "missing-required-part" // EC-I2 (xl/workbook.xml)
  | "invalid-xml"           // EC-I3 (XML decoding/parsing)
  | "unsupported-format"    // not an OOXML SpreadsheetML zip at all
  | "namespace-collision";  // two parts redefine the same prefix incompatibly

export class XlsxParseError extends Error {
  readonly name = "XlsxParseError";
  readonly kind: XlsxParseErrorKind;
  readonly path?: string;        // zip path of the offending part, if any
  readonly cause?: unknown;
  constructor(kind: XlsxParseErrorKind, message: string,
              opts?: { path?: string; cause?: unknown }) { ... }
}
```

| Trigger                                                    | Error                                             |
| ---------------------------------------------------------- | ------------------------------------------------- |
| `JSZip.loadAsync` throws (truncated zip, not a zip)        | `zip-corruption`                                  |
| Zip OK but `xl/workbook.xml` not present                   | `missing-required-part`, path = `xl/workbook.xml` |
| Zip OK, workbook present, but `<workbook>` root missing    | `unsupported-format`                              |
| `parseXml` throws on a part we tried to deserialize        | `invalid-xml`, path = part path, cause = err      |
| Two prefixes resolve to incompatible namespace URIs (rare) | `namespace-collision`                             |
| `XLSX.read` throws                                         | `invalid-xml`, cause = SheetJS error              |

Unrecognized worksheet sub-elements, unknown content types, and parts not
on our subset list are **never errors** — they degrade to opaque blobs.
This is the "fail loudly" principle from the architecture: we throw on
structural impossibilities, we degrade on un-modelled content.

Symbol export: `XlsxParseError` is exported from
`@officeai/xlsx/parser/errors` and re-exported from the package root.

---

## 8. Opaque-blob discipline

Everything not on the modeled list (§2 Stage C table) is left untouched
in `OoxmlContainer.parts`. The parser does **not** call `parseXml` on
opaque parts — they exist only as `Uint8Array` until either:

1. The serializer ships them through verbatim (the common case), or
2. A future workstream graduates them to first-class status (in which
   case the parser learns to deserialize them and the serializer learns
   to re-emit from the typed model — same pattern as DOCX with images).

The risk register in [`analysis.md`](analysis.md) §3 item 1 calls this
out as our most important invariant. Validation: post-export, our
synthetic-fixture round-trip test asserts that for every part path P
where `dirty[P] === false` the output bytes are SHA-256-equal to the
input (see [`serializer.md`](serializer.md) §10 and
[`acceptance-criteria.md`](acceptance-criteria.md)).

---

## 9. Namespace handling

OOXML namespaces are matched **by URI, not by prefix**. The same workbook
opened twice may have `xmlns:x="…/spreadsheetml/2006/main"` once and
`xmlns:s="…/spreadsheetml/2006/main"` once — both refer to the same
schema. Our `parseXml` retains the original prefixes verbatim
(`preserveOrder: true`), and our element matchers compare by **URI** when
possible:

```
const SPREADSHEETML_2006 = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
const RELATIONSHIPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
const X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
const XR = "http://schemas.microsoft.com/office/spreadsheetml/2014/revision"
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"

function matchByUri(elementEntry, expectedUri) -> boolean:
  // Resolve the prefix on the element's tag against the inherited
  // xmlns map; compare URIs. Fall back to prefix-equality when the
  // element carries no explicit binding (covers the vast majority of
  // real-world workbooks where every tool uses the standard prefixes).
```

For sheet-internal element traversal we use prefix-equality with the
canonical OOXML prefixes (`x:`, `r:`, `mc:`, `xr:`, `x14:`, …) — every
real Excel workbook uses these — and only fall back to URI matching for
the EC-I3 case. When EC-I3 _does_ trigger, namespace resolution happens
once at the root of each part and is cached for the rest of the parse.

The root-element xmlns declarations from each part are preserved verbatim
on the corresponding model node (`workbookRootAttrs`, `sheetRootAttrs`,
…) so the serializer round-trips them.

---

## 10. Encoding handling

Per [edge-cases EC-I4](edge-cases.md):

- `OoxmlContainer.readText(path)` strips a leading UTF-8 BOM and decodes
  with `TextDecoder("utf-8", { fatal: false })`.
- Before passing the raw text to `parseXml`, we inspect the XML
  declaration. If `<?xml version="1.0" encoding="windows-1252"?>` (or
  any other non-UTF-8 encoding) is present, we decode the part bytes
  with the declared encoding before re-encoding to UTF-8 for the
  parser. The container retains the original bytes — the encoding
  conversion only affects the in-memory tree, never `partHashes`.
- On serialize, **modified** parts are always emitted UTF-8 with the
  standard `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>`
  declaration. **Untouched** parts ship through verbatim — encoding and
  BOM included.

Old Excel builds emit Windows-1252 in `xl/comments1.xml` occasionally;
this is the canonical EC-I4 case and is regression-tested with a
synthetic fixture.

---

## 11. Performance budget

Target box: Node 22 LTS, GitHub Actions Linux 2-core ARM runner.

| Workload                                            | Budget                                        |
| --------------------------------------------------- | --------------------------------------------- |
| 50k-row × 10-col synthetic fixture, parse           | **< 1.5 s**                                   |
| 1k-row × 20-col, parse                              | < 50 ms                                       |
| 100k-row × 5-col with shared formulas, parse        | < 3 s                                         |
| Hash all parts (incl. `partHashes` materialization) | < 100 ms (50k-row fixture; SHA-256 dominates) |

Cost model:

- **JSZip unzip**: O(zip size). Bottleneck on the largest sheets.
- **SheetJS parse**: O(used cells) — dense mode is allocation-light.
- **Native XML pass**: O(sheet XML size). Only the sheets we care about
  - small targeted parts; ignores opaque parts entirely.
- **SHA-256 hashing**: O(part bytes). Done once per part at parse time.

We will measure with `scripts/perf-xlsx.mjs` (analogous to
`scripts/perf-docx.mjs`) and gate the CI budget at G5 in
[`acceptance-criteria.md`](acceptance-criteria.md).

If the budget is exceeded the recourse is, in order of cheapness:
(1) skip Stage C parts on a sheet that has no CF / DV / hyperlinks /
comments (detected via the sheet rels graph in O(1)); (2) lazily hash
parts on demand; (3) stream `xl/sharedStrings.xml`. (2) and (3) are
P1-only.

---

## 12. Determinism and tests

The parser is pure modulo `mintNodeId` (`IdMinter` for legacy comment
anchors Excel emits without stable ids — tests inject a deterministic
counter) and object identity. Tests live in
`packages/xlsx/src/parser/*.test.ts`, one per parsed part-family
(`parse-cells`, `parse-styles`, `parse-cf`, `parse-comments`,
`parse-merges`, `parse-defined-names`, `parse-encoding` for EC-I4,
`parse-namespace` for EC-I3, `parse-errors` covering every
`XlsxParseErrorKind`, `parse-roundtrip` for the byte-preservation
invariant — every part path P has `partHashes[P] ===
sha256(serialize(parse(buf)).get(P))`).

---

## 13. What the parser intentionally does NOT do

- **It does not evaluate formulas.** `cell.formula.cachedValue` is taken
  from the input as-is (per EC-S4). The formula engine recomputes only
  on mutation (see `formula-engine.md`).
- **It does not unfold shared formulas eagerly into per-cell formulas.**
  Per EC-F6 we _do_ expand on import (since we re-share at serialize),
  but the expansion is delegated to SheetJS — `cell.f` arrives already
  expanded. Our own code never sees `<f t="shared" si="…">`.
- **It does not validate** semantic constraints (sheet-name uniqueness,
  defined-name shadowing, range bounds). That's the command bus's job at
  mutation time. The parser accepts whatever the file says.
- **It does not normalize** style equality. Two cells with logically
  identical xf records keep their original xfIds. The
  `Styles.idForXf(...)` content-hash is exposed but does not collapse
  duplicates on parse.
- **It does not touch `xl/calcChain.xml`.** Treated as opaque on read;
  re-emitted only when the serializer rebuilds it (typically not in P0;
  see [`serializer.md`](serializer.md) §11).

Documenting these explicitly so a future implementor doesn't add scope
creep.

---

## 14. File layout

```
packages/xlsx/src/parser/
  parse.ts              # parseXlsx entry point + reconciliation
  errors.ts             # XlsxParseError + kind union
  sjs-adapter.ts        # XLSX.read wrapper with our options
  sheet-xml.ts          # native pass over xl/worksheets/sheetN.xml
  workbook-xml.ts       # xl/workbook.xml
  styles.ts             # xl/styles.xml (full)
  shared-strings.ts     # xl/sharedStrings.xml
  comments.ts           # legacy + threaded
  conditional-format.ts # CF rule kinds
  data-validation.ts    # DV rules (P0: read-only)
  hyperlinks.ts         # sheet-level hyperlinks
  defined-names.ts      # workbook + sheet-scope
  rels.ts               # rels graph helpers (uses @officeai/core)
  xml-helpers.ts        # findChild/attrOf/etc., shared with serializer
  parse.test.ts ... etc.
```

The split mirrors `packages/docx/src/parser/` so the patterns are
familiar to anyone who has worked on either format.
