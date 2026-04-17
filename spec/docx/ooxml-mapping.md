# DOCX — OOXML ↔ Model Mapping

> The most important spec doc. Every recognized model node corresponds to
> a row here. Anything not listed becomes an `OpaqueBlock` / `OpaqueXml`.
>
> Namespace prefix `w:` = `http://schemas.openxmlformats.org/wordprocessingml/2006/main`.

## Container parts

| Part path                                   | Loaded into model              |      Editable      | Notes                      |
| ------------------------------------------- | ------------------------------ | :----------------: | -------------------------- |
| `[Content_Types].xml`                       | container only                 | only on add/remove | re-emit verbatim otherwise |
| `_rels/.rels`                               | container only                 |         no         | re-emit verbatim           |
| `word/document.xml`                         | `DocxDocument.body`            |      **yes**       | the main story             |
| `word/_rels/document.xml.rels`              | `RelationshipGraph` (per part) |      partial       | new rels minted on insert  |
| `word/comments.xml`                         | `DocxDocument.comments`        |      **yes**       | new comments appended      |
| `word/commentsExtended.xml`                 | container only                 |         no         | preserved verbatim         |
| `word/commentsIds.xml`                      | container only                 |         no         | preserved verbatim         |
| `word/styles.xml`                           | container only                 |         no         | preserved verbatim         |
| `word/numbering.xml`                        | container only                 |         no         | preserved verbatim         |
| `word/header*.xml`, `word/footer*.xml`      | container only                 |         no         | preserved verbatim         |
| `word/footnotes.xml`, `word/endnotes.xml`   | container only                 |         no         | preserved verbatim         |
| `word/settings.xml`, `word/webSettings.xml` | container only                 |         no         | preserved verbatim         |
| `word/fontTable.xml`, `word/theme/...`      | container only                 |         no         | preserved verbatim         |
| `word/media/*`, `word/embeddings/*`         | container only                 |    only on add     | preserved verbatim         |
| `customXml/*`, `docProps/*`                 | container only                 |         no         | preserved verbatim         |

## Body block elements

| OOXML element                                 | Model node                 | Notes                                          |
| --------------------------------------------- | -------------------------- | ---------------------------------------------- |
| `w:body`                                      | `DocxDocument.body` (root) | the wrapper                                    |
| `w:p`                                         | `Paragraph`                | recognized in full                             |
| `w:tbl`                                       | `Table` (raw)              | parsed as opaque this session; cells preserved |
| `w:sectPr` (in body)                          | `SectionBreak` (raw)       | preserved verbatim                             |
| `w:sdt`                                       | `OpaqueBlock`              | structured document tag (form controls)        |
| `w:bookmarkStart`/`Start`/`End` (block-level) | `OpaqueBlock`              | preserved                                      |
| anything else under `w:body`                  | `OpaqueBlock`              | catch-all                                      |

## Paragraph properties (`w:pPr`)

| OOXML element                        | `ParagraphProperties` field | Notes                                                     |
| ------------------------------------ | --------------------------- | --------------------------------------------------------- |
| `w:pStyle val="X"`                   | `styleId = "X"`             |                                                           |
| `w:jc val="X"`                       | `alignment`                 | values: left/center/right/both → justify                  |
| `w:ind`                              | `indentation`               | attrs: `w:left`, `w:right`, `w:firstLine`, `w:hanging`    |
| `w:spacing`                          | `spacing`                   | attrs: `w:before`, `w:after`, `w:line`, `w:lineRule`      |
| `w:numPr/w:numId` + `w:numPr/w:ilvl` | `numbering`                 | preserved verbatim, not introspected                      |
| `w:rPr` (inside `w:pPr`)             | n/a                         | the paragraph mark's run properties — preserved as opaque |
| anything else                        | `opaqueProps[]`             |                                                           |

## Inline elements (within `w:p` or `w:hyperlink`)

| OOXML element                                    | Model node             | Notes                                            |
| ------------------------------------------------ | ---------------------- | ------------------------------------------------ |
| `w:r`                                            | `Run`                  | recognized in full                               |
| `w:hyperlink`                                    | `Hyperlink`            | `r:id` → `relationshipId`; `w:anchor` → `anchor` |
| `w:commentRangeStart`                            | `CommentRangeStart`    | `w:id` → `commentId`                             |
| `w:commentRangeEnd`                              | `CommentRangeEnd`      | `w:id` → `commentId`                             |
| `w:ins`                                          | `RevisionWrapper{ins}` | `w:author`, `w:date`, `w:id`                     |
| `w:del`                                          | `RevisionWrapper{del}` | as above; child runs typically have `w:delText`  |
| `w:smartTag`, `w:fldSimple`, `w:permStart`, etc. | `OpaqueInline`         | preserved                                        |

## Run properties (`w:rPr`)

| OOXML element | `RunProperties` field | Notes                                                  |
| ------------- | --------------------- | ------------------------------------------------------ |
| `w:b`         | `bold`                | presence = true; `val="0"` = false                     |
| `w:i`         | `italic`              | as above                                               |
| `w:u`         | `underline`           | `val="single"` etc. → string; absent or `none` → false |
| `w:strike`    | `strike`              |                                                        |
| `w:rFonts`    | `fontFamily`          | take `w:ascii` (others preserved as opaque)            |
| `w:sz`        | `fontSize`            | `val` is half-points                                   |
| `w:color`     | `color`               | hex without leading `#`                                |
| `w:highlight` | `highlight`           | named OOXML colors                                     |
| anything else | `opaqueProps[]`       |                                                        |

## Run children (`w:r/*`)

| OOXML element        | `RunChild`                                                                          |
| -------------------- | ----------------------------------------------------------------------------------- |
| `w:t`                | `{kind:"text", text, xmlSpacePreserve}` (sets preserve when `xml:space="preserve"`) |
| `w:delText`          | `{kind:"text", ...}` inside a `RevisionWrapper{del}`                                |
| `w:br`               | `{kind:"break", breakType?}`                                                        |
| `w:tab`              | `{kind:"tab"}`                                                                      |
| `w:drawing`          | `{kind:"drawing", raw}` (images preserved opaquely)                                 |
| `w:object`, `w:pict` | `{kind:"opaque", raw}`                                                              |

## Comments part (`word/comments.xml`)

| OOXML element            | Model             |
| ------------------------ | ----------------- |
| `w:comments`             | `comments[]` root |
| `w:comment` with attrs   | `DocxComment`     |
| `w:p` inside `w:comment` | `body[Paragraph]` |

When adding a comment via `docx:add-comment`:

1. Mint a new `commentId` (max(existing) + 1, as string).
2. Append a `w:comment` to `word/comments.xml`.
3. Insert `w:commentRangeStart`, target text, `w:commentRangeEnd`, then a
   run containing `w:commentReference` into the body at the requested range.
4. If `word/comments.xml` did not previously exist:
   - Create it with the standard root.
   - Add a relationship in `word/_rels/document.xml.rels` of type
     `http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments`.
   - Add the override in `[Content_Types].xml`.

## Relationships

| Type URI suffix                                                                                                                                                  | Used for                                          |
| ---------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------- |
| `/relationships/hyperlink`                                                                                                                                       | external hyperlinks; preserved + minted on insert |
| `/relationships/comments`                                                                                                                                        | the comments part                                 |
| `/relationships/commentsExtended`, `/commentsIds`                                                                                                                | preserved verbatim                                |
| `/relationships/image`                                                                                                                                           | inline images; preserved verbatim                 |
| `/relationships/header`, `/footer`, `/numbering`, `/styles`, `/fontTable`, `/theme`, `/webSettings`, `/settings`, `/footnotes`, `/endnotes`, `/glossaryDocument` | preserved verbatim                                |

## Anything else

If we encounter an element under `w:body` that doesn't match a row above,
it becomes an `OpaqueBlock` whose `raw.subtree` is the parser's verbatim
subtree. Same for inline-context unknowns (`OpaqueInline`).

The serializer re-emits opaque subtrees byte-for-byte (modulo XML
serializer normalization — see `serializer.md` §Whitespace).
