# PPTX — Parser

## Algorithm

```
parsePptx(buffer):
  container = OoxmlContainer.load(buffer)
  presentationXml = container.readText("ppt/presentation.xml")
  presentationTree = parseXml(presentationXml)
  presentationRootAttrs = extractRootAttrs(presentationTree, "p:presentation")

  slideOrder, sldSz, notesSz, presOpaqueTail =
    parsePresentationHead(presentationTree, container)

  // Resolve slide rels & part paths in display order.
  presRels = container.relationships.get("ppt/presentation.xml")
  slidePartPaths = []
  for sldId in slideOrder:
    rel = presRels.find(r => r.id === sldId.rId)
    slidePartPaths.push(normalize(rel.target, base="ppt/"))

  slides = slidePartPaths.map((path, idx) =>
    parseSlide(container, path, slideOrder[idx])
  )

  masters = enumerateOpaque(container, "ppt/slideMasters/")
  layouts = enumerateOpaque(container, "ppt/slideLayouts/")
  theme   = enumerateOpaque(container, "ppt/theme/")
  notesSlides = enumerateOpaque(container, "ppt/notesSlides/")
  media   = enumerateMedia(container, "ppt/media/")

  idGen = seedIdGen(slideOrder, container)

  return PptxSnapshot {
    format: "pptx",
    revision: 0,
    root: PptxPresentation {
      id: mintNodeId(),
      slides, slideSize: sldSz, notesSize: notesSz,
      masters, layouts, theme, notesSlides, media,
      presentationRootAttrs, presentationOpaqueTail: presOpaqueTail,
      idGen,
    },
    partHashes: container.allHashes(),
    dirty: emptyDirty(),
    container,
  }
```

## parsePresentationHead

Walk `<p:presentation>`'s children in order. Recognized children populate
typed fields; everything else lands in `presentationOpaqueTail`.

| Child                | Typed                                       |
| -------------------- | ------------------------------------------- |
| `<p:sldIdLst>`       | `slideOrder: [{ slideId, rId }]`            |
| `<p:sldSz>`          | `slideSize` (cx/cy attrs, optional `type`)  |
| `<p:notesSz>`        | `notesSize`                                 |
| anything else        | `presentationOpaqueTail.push(opaqueXml(c))` |

`slideOrder` preserves source order — that IS the display order.

## parseSlide

```
parseSlide(container, partPath, sldIdEntry):
  xml = container.readText(partPath)
  tree = parseXml(xml)
  slideRootAttrs = extractRootAttrs(tree, "p:sld")

  cSld = findChild(tree, "p:sld/p:cSld")
  cSldAttrs = extractRootAttrs(cSld, "p:cSld")
  spTree = findChild(cSld, "p:spTree")

  spTreeHead, shapesXml, slideOpaqueTail =
    classifySpTreeChildren(spTree)
  // spTreeHead   = leading <p:nvGrpSpPr> + <p:grpSpPr> for the spTree itself
  // shapesXml    = remaining children that look like shapes
  // slideOpaqueTail catches non-spTree tail elements: clrMapOvr, transition,
  // timing, extLst, hf, …

  shapes = shapesXml.map(c => parseShape(c, slideRelsPath(partPath), container))

  layoutPartPath = resolveSlideLayout(container, partPath)
  notesSlidePartPath = resolveNotesSlide(container, partPath)

  return Slide {
    id: mintNodeId(),
    partPath,
    slideId: sldIdEntry.slideId,
    relId:   sldIdEntry.rId,
    layoutPartPath, notesSlidePartPath,
    shapes,
    slideOpaqueTail,
    slideRootAttrs, cSldAttrs, spTreeHead,
  }
```

## parseShape

Dispatch on the element tag. Anything not matched becomes `OpaqueShape`.

```
parseShape(c, slideRelsPath, container):
  switch c.tag:
    case "p:sp":   return parseSp(c, slideRelsPath, container)
    case "p:pic":  return parsePic(c, slideRelsPath, container)
    case "p:grpSp":return parseGrpSp(c, slideRelsPath, container)
    default:       return parseOpaqueShape(c)
```

### parseSp

```
parseSp(c, slideRelsPath, container):
  nvSpPr = findChild(c, "p:nvSpPr")
  cNvPr  = findChild(nvSpPr, "p:cNvPr")
  nvPr   = findChild(nvSpPr, "p:nvPr")
  ph     = nvPr ? findChild(nvPr, "p:ph") : null

  cNvPrId = parseInt(cNvPr.attrs["@_id"])
  name    = cNvPr.attrs["@_name"] ?? ""
  placeholder = ph
    ? { type: ph.attrs["@_type"], idx: ph.attrs["@_idx"] ? parseInt(ph.attrs["@_idx"]) : undefined }
    : undefined
  nvSpPrTail = childrenExcept(nvSpPr, ["p:cNvPr","p:nvPr"]) // typically <p:cNvSpPr>; nvPr re-emitted by serializer

  spPr = findChild(c, "p:spPr")
  xfrm = findChild(spPr, "a:xfrm")
  position = xfrm ? parseOff(findChild(xfrm, "a:off")) : undefined
  size     = xfrm ? parseExt(findChild(xfrm, "a:ext")) : undefined
  spPrTail = childrenExcept(spPr, ["a:xfrm"])

  styleRaw = findChild(c, "p:style") ? captureOpaque(findChild(c, "p:style")) : undefined

  txBodyXml = findChild(c, "p:txBody")
  if txBodyXml:
    return TextShape {
      kind: "text", id: mintNodeId(), cNvPrId, name, position, size,
      placeholder,
      nvSpPrTail, spPrTail, styleRaw,
      txBody: parseTxBody(txBodyXml),
    }
  else:
    return OpaqueShape {
      kind: "opaque", id: mintNodeId(), cNvPrId, name, position, size,
      tag: "p:sp", raw: captureOpaque(c),
    }
```

### parsePic

```
parsePic(c, slideRelsPath, container):
  nvPicPr = findChild(c, "p:nvPicPr")
  cNvPr   = findChild(nvPicPr, "p:cNvPr")
  cNvPrId = parseInt(cNvPr.attrs["@_id"])
  name    = cNvPr.attrs["@_name"] ?? ""
  nvPicPrTail = childrenExcept(nvPicPr, ["p:cNvPr"])

  blipFill = findChild(c, "p:blipFill")
  blip     = findChild(blipFill, "a:blip")
  mediaRelId = blip.attrs["@_r:embed"]
  blipFillTail = childrenExcept(blipFill, ["a:blip"]) // stretch, srcRect, …

  spPr = findChild(c, "p:spPr")
  xfrm = findChild(spPr, "a:xfrm")
  position = parseOff(findChild(xfrm, "a:off"))
  size     = parseExt(findChild(xfrm, "a:ext"))
  spPrTail = childrenExcept(spPr, ["a:xfrm"])
  styleRaw = findChild(c, "p:style") ? captureOpaque(findChild(c, "p:style")) : undefined

  // Resolve mediaPartPath through the slide's rels graph.
  rels = container.relationships.get(slideRelsPath) ?? []
  rel  = rels.find(r => r.id === mediaRelId)
  mediaPartPath = normalize(rel.target, base=dirname(slideRelsPath).replace("/_rels", ""))

  return Picture {
    kind: "pic", id: mintNodeId(), cNvPrId, name, position, size,
    mediaRelId, mediaPartPath,
    nvPicPrTail, blipFillTail, spPrTail, styleRaw,
  }
```

### parseGrpSp

```
parseGrpSp(c, slideRelsPath, container):
  nvGrpSpPr = findChild(c, "p:nvGrpSpPr")
  cNvPr     = findChild(nvGrpSpPr, "p:cNvPr")
  cNvPrId   = parseInt(cNvPr.attrs["@_id"])
  name      = cNvPr.attrs["@_name"] ?? ""

  grpSpPr   = findChild(c, "p:grpSpPr")
  xfrm      = findChild(grpSpPr, "a:xfrm")
  position  = xfrm ? parseOff(findChild(xfrm, "a:off")) : undefined
  size      = xfrm ? parseExt(findChild(xfrm, "a:ext")) : undefined
  // Capture <a:chOff>+<a:chExt> as a single opaque slice for verbatim re-emit.
  chOffExtRaw = captureOpaqueRange(xfrm, ["a:chOff","a:chExt"])
  grpSpPrTail = childrenExcept(grpSpPr, ["a:xfrm"])

  childShapeXml = childrenWhere(c, tag in ["p:sp","p:pic","p:grpSp","p:cxnSp","p:graphicFrame","mc:AlternateContent","p:contentPart"])
  children = childShapeXml.map(child => parseShape(child, slideRelsPath, container))

  return GroupShape {
    kind: "group", id: mintNodeId(), cNvPrId, name, position, size,
    chOffExtRaw, grpSpPrTail, children,
  }
```

### parseOpaqueShape

Capture the entire element subtree. Best-effort xfrm extraction for the
renderer:

```
parseOpaqueShape(c):
  spPr = findChildAnywhere(c, "p:spPr") ?? findChildAnywhere(c, "p:grpSpPr")
  xfrm = spPr ? findChild(spPr, "a:xfrm") : null
  position = xfrm ? parseOff(findChild(xfrm, "a:off")) : undefined
  size     = xfrm ? parseExt(findChild(xfrm, "a:ext")) : undefined
  cNvPr = findChildAnywhere(c, "p:cNvPr")
  cNvPrId = cNvPr ? parseInt(cNvPr.attrs["@_id"]) : 0
  name    = cNvPr ? (cNvPr.attrs["@_name"] ?? "") : ""

  return OpaqueShape {
    kind: "opaque", id: mintNodeId(), cNvPrId, name, position, size,
    tag: c.tag, raw: captureOpaque(c),
  }
```

## parseTxBody

```
parseTxBody(node):
  bodyPrRaw   = findChild(node, "a:bodyPr")    ? captureOpaque(...) : undefined
  lstStyleRaw = findChild(node, "a:lstStyle")  ? captureOpaque(...) : undefined
  paragraphs = []
  for child in node.children:
    if child.tag === "a:p":
      paragraphs.push(parseTextParagraph(child))
  return TextBody { bodyPrRaw, lstStyleRaw, paragraphs }
```

## parseTextParagraph

```
parseTextParagraph(p):
  pPr = findChild(p, "a:pPr")
  properties = parsePPr(pPr)
  runs = []
  endParaRPrRaw = undefined
  for child in p.children:
    if child === pPr: continue
    if child.tag === "a:r":           runs.push(parseTextRun(child))
    elif child.tag === "a:br":        runs.push(parseLineBreak(child))
    elif child.tag === "a:fld":       runs.push(parseField(child))   // text="" + opaque
    elif child.tag === "a:endParaRPr": endParaRPrRaw = captureOpaque(child)
    else:                             runs.push(parseOpaqueRun(child))
  return TextParagraph { id: mintNodeId(), properties, runs, endParaRPrRaw }
```

## parsePPr

| Attr / child         | Field                       |
| -------------------- | --------------------------- |
| `@_lvl`              | `level = parseInt(...)`     |
| `@_algn`             | `alignment` (l/ctr/r/just)  |
| anything else        | `opaqueProps[]`             |

Note: PPTX uses `algn="ctr"` for center, `"l"` for left, `"r"` for
right, `"just"` for justify. Anything else (`"justLow"`, `"dist"`,
`"thaiDist"`) falls back to "justify" with the original value preserved
in `opaqueProps`.

## parseTextRun

```
parseTextRun(r):
  rPr = findChild(r, "a:rPr")
  properties = parseRPr(rPr)
  t = findChild(r, "a:t")
  text = t ? collectText(t) : ""
  return TextRun { id: mintNodeId(), properties, text }
```

## parseRPr

| OOXML                                                | Field                                    |
| ---------------------------------------------------- | ---------------------------------------- |
| `@_b="1"`                                            | `bold = true`                            |
| `@_b="0"`                                            | `bold = false`                           |
| `@_i="1"` / `@_i="0"`                                | `italic`                                 |
| `@_u="sng"` / `@_u="dbl"` / etc.                     | `underline = "..." `                     |
| `@_u="none"`                                         | `underline = false`                      |
| `@_strike="..."`                                     | `strike = true` (any value other than "noStrike") |
| `@_sz="3200"`                                        | `fontSizeHundredths = 3200`              |
| `<a:latin @typeface="...">` (child)                  | `fontFamily`                             |
| `<a:solidFill><a:srgbClr @val="RRGGBB">` (descendant)| `color = "RRGGBB"`                       |
| anything else                                        | `opaqueProps[]`                          |

The `<a:rPr>` carrier itself is preserved per-run (its full attribute
set goes into `opaqueProps[]` if any non-introspected attrs survive,
e.g. `lang`, `dirty`, `kern`).

## parseLineBreak

```
parseLineBreak(node):
  rPr = findChild(node, "a:rPr")
  properties = rPr ? parseRPr(rPr) : {}
  return TextRun { id: mintNodeId(), properties, text: "", isLineBreak: true }
```

## seedIdGen

```
seedIdGen(slideOrder, container):
  maxSlideId = max(slideOrder.map(e => e.slideId), 255) // PowerPoint floor
  maxRelId   = max(parseRelIds(container.allRels()))    // global high-water
  maxSlideIdx= max(slidePartIndices(container.parts), 0)
  maxMediaIdx= max(mediaPartIndices(container.parts), 0)
  return {
    nextSlideId:        maxSlideId + 1,
    nextRelId:          maxRelId + 1,
    nextSlidePartIndex: maxSlideIdx + 1,
    nextMediaPartIndex: maxMediaIdx + 1,
  }
```

## enumerateOpaque

For each part under the given prefix that is XML, build an
`OpaquePart { partPath, raw: captureOpaqueRoot(parseXml(text)) }`. The
`raw` slice is for code paths that need to *read* the part structurally
(theme color resolution at render time). The serializer does NOT
re-stringify it — it copies bytes from the container cache.

`enumerateMedia` is similar but stores `bytes` and `sha256`.

## Errors

The parser is **strict**:

- Missing `ppt/presentation.xml` → throw `PptxParseError("missing-main-part")`.
- XML parse failure → throw `PptxParseError("invalid-xml", { part, cause })`.
- A `<p:sldId>` whose `rId` does not resolve in `presentation.xml.rels` → throw `PptxParseError("dangling-slide-rel", { sldId, rId })`.
- A `<p:pic>` whose `r:embed` does not resolve in the slide's rels → throw `PptxParseError("dangling-image-rel", { slidePart, rId })`.
- Unrecognized element under `<p:spTree>` is **not** an error — it becomes `OpaqueShape`.

All errors are structured (`name`, `code`, `cause`, `partPath?`). Never silent.

## Determinism

`mintNodeId()` is the only source of randomness. In tests we inject a deterministic counter so snapshots are stable. The parser is otherwise pure.
