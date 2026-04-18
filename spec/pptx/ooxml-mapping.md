# PPTX — OOXML ↔ Model Mapping

> The most important spec doc. Every recognized model node corresponds to
> a row here. Anything not listed becomes an `OpaqueShape` / `OpaquePart`.
>
> Namespace prefixes:
>
> - `p:` = `http://schemas.openxmlformats.org/presentationml/2006/main`
> - `a:` = `http://schemas.openxmlformats.org/drawingml/2006/main`
> - `r:` = `http://schemas.openxmlformats.org/officeDocument/2006/relationships`
> - `mc:` = `http://schemas.openxmlformats.org/markup-compatibility/2006`

## Container parts

| Part path                                                  | Loaded into model              |      Editable      | Notes                                   |
| ---------------------------------------------------------- | ------------------------------ | :----------------: | --------------------------------------- |
| `[Content_Types].xml`                                      | container only                 | only on add/remove | re-emit verbatim otherwise              |
| `_rels/.rels`                                              | container only                 |         no         | re-emit verbatim                        |
| `ppt/presentation.xml`                                     | `PptxPresentation` head        |      **yes**       | sldIdLst rewrites on slide CRUD/move    |
| `ppt/_rels/presentation.xml.rels`                          | `RelationshipGraph`            |      **yes**       | new rels minted on add-slide            |
| `ppt/slides/slide{N}.xml`                                  | `Slide` typed                  |      **yes**       | the editable canvas                     |
| `ppt/slides/_rels/slide{N}.xml.rels`                       | `RelationshipGraph` (per part) |      **yes**       | new rels minted on insert-image, etc.   |
| `ppt/slideLayouts/slideLayout{N}.xml`                      | `OpaquePart`                   |         no         | preserved verbatim                      |
| `ppt/slideLayouts/_rels/slideLayout{N}.xml.rels`           | container only                 |         no         | preserved verbatim                      |
| `ppt/slideMasters/slideMaster{N}.xml`                      | `OpaquePart`                   |         no         | preserved verbatim                      |
| `ppt/slideMasters/_rels/slideMaster{N}.xml.rels`           | container only                 |         no         | preserved verbatim                      |
| `ppt/theme/theme{N}.xml`                                   | `OpaquePart`                   |         no         | preserved verbatim; read at render time |
| `ppt/notesSlides/notesSlide{N}.xml`                        | `OpaquePart`                   |         no         | dropped on `delete-slide`               |
| `ppt/notesSlides/_rels/notesSlide{N}.xml.rels`             | container only                 |         no         | dropped on `delete-slide`               |
| `ppt/notesMasters/notesMaster1.xml`                        | container only                 |         no         | preserved verbatim                      |
| `ppt/notesMasters/_rels/notesMaster1.xml.rels`             | container only                 |         no         | preserved verbatim                      |
| `ppt/media/*`                                              | `MediaPart`                    |    only on add     | dedup by SHA-256                        |
| `ppt/embeddings/*`, `ppt/charts/*`, `ppt/diagrams/*`       | container only                 |         no         | preserved verbatim                      |
| `ppt/tableStyles.xml`, `ppt/viewProps.xml`, `presProps.xml`| container only                 |         no         | preserved verbatim                      |
| `docProps/*`, `customXml/*`                                | container only                 |         no         | preserved verbatim                      |

## `ppt/presentation.xml` structure

| OOXML element                | Model destination                          | Notes                                            |
| ---------------------------- | ------------------------------------------ | ------------------------------------------------ |
| `<p:presentation>` root attrs| `presentationRootAttrs`                    | re-emitted verbatim                              |
| `<p:sldIdLst>`               | `slides[*].{slideId,relId,partPath}`       | drives slide ordering                            |
| `<p:sldSz>`                  | `slideSize`                                | EMU dimensions + type                            |
| `<p:notesSz>`                | `notesSize`                                |                                                  |
| any other tail child         | `presentationOpaqueTail[]`                 | sldMasterIdLst, notesMasterIdLst, defaultTextStyle, custShowLst, embeddedFontLst, extLst, … |

## `ppt/slides/slide{N}.xml` — slide root

| OOXML path                    | Model field                  | Notes                                              |
| ----------------------------- | ---------------------------- | -------------------------------------------------- |
| `<p:sld>` root attrs          | `slideRootAttrs`             | xmlns declarations + `show="0"` etc.               |
| `<p:cSld>` attrs              | `cSldAttrs`                  | name attribute                                     |
| `<p:cSld><p:bg>` (if present) | `spTreeHead[0]`              | preserved verbatim as the first head item          |
| `<p:cSld><p:spTree>` head     | `spTreeHead[]`               | `<p:nvGrpSpPr>`, `<p:grpSpPr>` for the spTree     |
| children of `<p:spTree>`      | `shapes[]`                   | each child becomes a typed Shape or OpaqueShape    |
| `<p:clrMapOvr>`, `<p:transition>`, `<p:timing>`, `<p:extLst>`, … | `slideOpaqueTail[]` | preserved verbatim |

## Shape elements (children of `<p:spTree>`)

| OOXML element        | Model node                | Notes                                                       |
| -------------------- | ------------------------- | ----------------------------------------------------------- |
| `<p:sp>`             | `TextShape` if has txBody | shape with text frame                                       |
| `<p:sp>` (no txBody) | `OpaqueShape`             | pure-geometry shape — preserved verbatim                    |
| `<p:pic>`            | `Picture`                 | media-backed                                                |
| `<p:grpSp>`          | `GroupShape`              | recursive; children parsed                                  |
| `<p:cxnSp>`          | `OpaqueShape`             | connector                                                   |
| `<p:graphicFrame>` (table) | `TableShape`        | `<a:graphicData @uri="…/drawingml/2006/table">` → typed (F2) |
| `<p:graphicFrame>` (other) | `OpaqueShape`       | host for `c:chart` / `dgm:relIds` (SmartArt)                |
| `<mc:AlternateContent>` | `OpaqueShape`          | mc-wrapped content                                          |
| anything else        | `OpaqueShape`             | catch-all                                                   |

## `<p:sp>` internals (TextShape only)

| OOXML path                                  | Model field                    | Notes                                                |
| ------------------------------------------- | ------------------------------ | ---------------------------------------------------- |
| `<p:nvSpPr><p:cNvPr @id @name>`             | `cNvPrId`, `name`              |                                                      |
| `<p:nvSpPr><p:nvPr><p:ph @type @idx>`       | `placeholder`                  | absent on non-placeholder shapes                     |
| any other `<p:nvSpPr>` child / sibling      | `nvSpPrTail[]`                 | `<p:cNvSpPr>`, extLst inside nvSpPr                  |
| `<p:spPr><a:xfrm><a:off @x @y>`             | `position`                     | EMU                                                  |
| `<p:spPr><a:xfrm><a:ext @cx @cy>`           | `size`                         | EMU                                                  |
| any other `<p:spPr>` child                  | `spPrTail[]`                   | prstGeom, custGeom, fills, lines, effects, …         |
| `<p:style>`                                 | `styleRaw`                     | preserved verbatim                                   |
| `<p:txBody>`                                | `txBody`                       | typed (see below)                                    |

## `<p:pic>` internals

| OOXML path                                     | Model field           | Notes                           |
| ---------------------------------------------- | --------------------- | ------------------------------- |
| `<p:nvPicPr><p:cNvPr @id @name>`               | `cNvPrId`, `name`     |                                 |
| any other `<p:nvPicPr>` child                  | `nvPicPrTail[]`       |                                 |
| `<p:blipFill><a:blip @r:embed>`                | `mediaRelId`          | resolves through slide rels     |
| any other `<p:blipFill>` child                 | `blipFillTail[]`      | stretch, srcRect, …             |
| `<p:spPr><a:xfrm><a:off>` / `<a:ext>`          | `position`, `size`    |                                 |
| any other `<p:spPr>` child                     | `spPrTail[]`          | prstGeom etc.                   |
| `<p:style>`                                    | `styleRaw`            |                                 |

## `<p:grpSp>` internals

| OOXML path                                                  | Model field        | Notes                            |
| ----------------------------------------------------------- | ------------------ | -------------------------------- |
| `<p:nvGrpSpPr><p:cNvPr @id @name>`                          | `cNvPrId`, `name`  |                                  |
| `<p:grpSpPr><a:xfrm><a:off>` / `<a:ext>`                    | `position`, `size` |                                  |
| `<p:grpSpPr><a:xfrm><a:chOff>` / `<a:chExt>` (sibling pair) | `chOffExtRaw`      | preserved verbatim as one slice  |
| any other `<p:grpSpPr>` child                               | `grpSpPrTail[]`    |                                  |
| children (`<p:sp>`, `<p:pic>`, `<p:grpSp>`, …)              | `children[]`       | recursive                        |

## `<p:txBody>` internals

| OOXML element        | Model field              | Notes                                       |
| -------------------- | ------------------------ | ------------------------------------------- |
| `<a:bodyPr>`         | `bodyPrRaw`              | margins, autofit — opaque                   |
| `<a:lstStyle>`       | `lstStyleRaw`            | per-level defaults — opaque                 |
| `<a:p>`              | `paragraphs[]` entry     | typed                                       |

## `<a:p>` internals

| OOXML element                  | Model field                           | Notes                                |
| ------------------------------ | ------------------------------------- | ------------------------------------ |
| `<a:pPr @lvl @algn …>`         | `properties.{level, alignment}`       |                                      |
| any other `<a:pPr>` child      | `properties.opaqueProps[]`            |                                      |
| `<a:r>`                        | `runs[]` entry, `isLineBreak=false`   |                                      |
| `<a:br>`                       | `runs[]` entry, `isLineBreak=true`    | empty `text`, `properties` from `<a:rPr>` if present |
| `<a:fld>`                      | `runs[]` entry — opaque text run      | preserved verbatim under `opaqueProps`               |
| `<a:endParaRPr>`               | `endParaRPrRaw`                       | preserved verbatim                   |

## `<a:r>` / `<a:rPr>` internals

| OOXML path                      | Model field                              | Notes                                 |
| ------------------------------- | ---------------------------------------- | ------------------------------------- |
| `<a:r><a:t>...</a:t></a:r>`     | `text`                                   |                                       |
| `<a:rPr @b="1">`                | `properties.bold = true`                 | absence = false; `b="0"` = false      |
| `<a:rPr @i="1">`                | `properties.italic = true`               |                                       |
| `<a:rPr @u="…">`                | `properties.underline`                   | "sng"/"dbl"/… → string; "none" → false |
| `<a:rPr @strike="…">`           | `properties.strike = true`               |                                       |
| `<a:rPr @sz="3200">`            | `properties.fontSizeHundredths`          | hundredths-of-a-point                 |
| `<a:rPr><a:latin @typeface="…">`| `properties.fontFamily`                  | east-asian/cs preserved opaquely      |
| `<a:rPr><a:solidFill><a:srgbClr @val="RRGGBB">` | `properties.color`       | direct hex                            |
| anything else under `<a:rPr>`   | `properties.opaqueProps[]`               | schemeClr, lumMod, hlinkClick, …      |

## Slide rels graph (`ppt/slides/_rels/slideN.xml.rels`)

Resolved into `RelationshipGraph` keyed by the **slide part path**.

| Type URI suffix                                                     | Used for                                  |
| ------------------------------------------------------------------- | ----------------------------------------- |
| `/relationships/slideLayout`                                        | binds slide → layout (`layoutPartPath`)   |
| `/relationships/notesSlide`                                         | binds slide → notesSlide                  |
| `/relationships/image`                                              | picture media references                  |
| `/relationships/hyperlink`                                          | text-run hyperlinks (preserved opaquely)  |
| `/relationships/chart`, `/diagramData`, `/diagramLayout`, …         | preserved verbatim                        |
| `/relationships/oleObject`, `/audio`, `/video`, …                   | preserved verbatim                        |

## Presentation rels graph (`ppt/_rels/presentation.xml.rels`)

| Type URI suffix                                | Used for                                 |
| ---------------------------------------------- | ---------------------------------------- |
| `/relationships/slide`                         | binds presentation → each slide          |
| `/relationships/slideMaster`                   | preserved verbatim                       |
| `/relationships/notesMaster`                   | preserved verbatim                       |
| `/relationships/handoutMaster`                 | preserved verbatim                       |
| `/relationships/theme`                         | preserved verbatim                       |
| `/relationships/presProps`, `/viewProps`, `/tableStyles` | preserved verbatim             |

## Anything else

If we encounter an element under `<p:spTree>` that doesn't match a row above, it becomes an `OpaqueShape` whose `raw.subtree` is the parser's verbatim subtree. Same for non-introspected children of `<p:txBody>`, `<a:p>`, `<a:rPr>`, etc., which land in their respective `*Tail` / `opaqueProps[]` slots.

The serializer re-emits opaque subtrees byte-for-byte (modulo XML serializer normalization documented in `serializer.md` §Whitespace).
