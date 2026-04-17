# PPTX — Serializer

## Goal

Produce an OOXML zip such that:

1. Every part **not touched** by any command since the last load is byte-identical to the loaded bytes.
2. Every part **touched** by commands is re-serialized from the model in a shape PowerPoint / LibreOffice / Keynote / Google Slides accept without a repair dialog.
3. Opaque shapes / parts are re-emitted from their captured raw slices (modulo the documented XML serializer normalization caveat).

## Algorithm

```
serializePptx(snapshot):
  container = snapshot.container.clone()

  // Per-part decision tree.
  if snapshot.dirty.presentation:
    container.writeText("ppt/presentation.xml", serializePresentationXml(snapshot.root))

  for slide in snapshot.root.slides:
    if snapshot.dirty.slides.has(slide.partPath):
      container.writeText(slide.partPath, serializeSlideXml(slide))

  for relsPath in snapshot.dirty.relationships:
    container.writeText(relsPath, serializeRelsXml(container.relationships.get(relsPath)))

  for mediaPath in snapshot.dirty.media:
    container.writeBytes(mediaPath, snapshot.root.media.get(mediaPath).bytes)

  if snapshot.dirty.contentTypes:
    container.writeText("[Content_Types].xml", serializeContentTypesXml(container.contentTypes))

  // Drop parts that commands removed (delete-slide, delete-picture, …).
  for path in snapshot.removedParts:
    container.removePart(path)

  return container.serialize()
```

The default for a freshly-parsed snapshot: **no dirty flags set, no removed parts → output is byte-identical to input** (the load-bearing invariant, asserted in the no-edit roundtrip test).

## serializePresentationXml

```
serializePresentationXml(presentation):
  sldIdLst = presentation.slides.map(s => ({
    "p:sldId": [],
    ":@": { "@_id": String(s.slideId), "@_r:id": s.relId },
  }))

  body = []
  // Important: opaque tail items were captured in source order.
  // sldIdLst is re-emitted at the position the parser observed (typically
  // first), driven by a small `sldIdLstAnchor` pointer recorded at parse
  // time on `presentationOpaqueTail`. If absent (no slides at parse time),
  // sldIdLst is emitted before sldSz.
  body = mergeSldIdLstWithTail(sldIdLst, presentation.presentationOpaqueTail)

  body.push({ "p:sldSz": [], ":@": { "@_cx": ..., "@_cy": ..., ...(type ? { "@_type": type } : {}) } })
  if presentation.notesSize:
    body.push({ "p:notesSz": [], ":@": { ... } })

  tree = [
    { "?xml": [], ":@": { "@_version": "1.0", "@_encoding": "UTF-8", "@_standalone": "yes" } },
    { "p:presentation": body, ":@": presentation.presentationRootAttrs },
  ]
  return serializeXml(tree)
```

## serializeSlideXml

```
serializeSlideXml(slide):
  spTreeChildren = [
    ...slide.spTreeHead,                       // <p:nvGrpSpPr>, <p:grpSpPr>
    ...slide.shapes.map(serializeShape),
  ]

  cSld = {
    "p:cSld": [{ "p:spTree": spTreeChildren }],
    ":@": slide.cSldAttrs,
  }

  // slideOpaqueTail was captured in source order — it includes
  // <p:clrMapOvr>, <p:transition>, <p:timing>, <p:extLst>, etc.
  body = [cSld, ...slide.slideOpaqueTail]

  tree = [
    { "?xml": [], ":@": { "@_version": "1.0", "@_encoding": "UTF-8", "@_standalone": "yes" } },
    { "p:sld": body, ":@": slide.slideRootAttrs },
  ]
  return serializeXml(tree)
```

## serializeShape

| Node                  | Output                                            |
| --------------------- | ------------------------------------------------- |
| `TextShape`           | `serializeSp(s)` (rebuild `<p:sp>`)               |
| `Picture`             | `serializePic(s)` (rebuild `<p:pic>`)             |
| `GroupShape`          | `serializeGrpSp(s)` (rebuild `<p:grpSp>`)         |
| `OpaqueShape`         | `s.raw.subtree` (verbatim slice)                  |

### serializeSp (TextShape)

```
serializeSp(s):
  nvSpPr = {
    "p:nvSpPr": [
      { "p:cNvPr": [], ":@": { "@_id": String(s.cNvPrId), "@_name": s.name } },
      ...s.nvSpPrTail,             // typically <p:cNvSpPr>; nvPr if any
    ],
    ":@": {},
  }

  spPrChildren = []
  if s.position && s.size:
    spPrChildren.push(buildXfrm(s.position, s.size))
  spPrChildren.push(...s.spPrTail)
  spPr = { "p:spPr": spPrChildren, ":@": {} }

  body = [nvSpPr, spPr]
  if s.styleRaw: body.push(s.styleRaw.subtree[0])
  body.push(serializeTxBody(s.txBody))
  return { "p:sp": body, ":@": {} }
```

### serializePic (Picture)

```
serializePic(s):
  nvPicPr = {
    "p:nvPicPr": [
      { "p:cNvPr": [], ":@": { "@_id": String(s.cNvPrId), "@_name": s.name } },
      ...s.nvPicPrTail,
    ],
    ":@": {},
  }
  blipFill = {
    "p:blipFill": [
      { "a:blip": [], ":@": { "@_r:embed": s.mediaRelId } },
      ...s.blipFillTail,
    ],
    ":@": {},
  }
  spPr = { "p:spPr": [buildXfrm(s.position, s.size), ...s.spPrTail], ":@": {} }

  body = [nvPicPr, blipFill, spPr]
  if s.styleRaw: body.push(s.styleRaw.subtree[0])
  return { "p:pic": body, ":@": {} }
```

### serializeGrpSp (GroupShape)

```
serializeGrpSp(s):
  nvGrpSpPr = {
    "p:nvGrpSpPr": [
      { "p:cNvPr": [], ":@": { "@_id": String(s.cNvPrId), "@_name": s.name } },
      // <p:cNvGrpSpPr>, <p:nvPr> are part of the spPr tail to keep typing tight
    ],
    ":@": {},
  }
  xfrm = buildXfrmWithChild(s.position, s.size, s.chOffExtRaw)
  grpSpPr = { "p:grpSpPr": [xfrm, ...s.grpSpPrTail], ":@": {} }
  body = [nvGrpSpPr, grpSpPr, ...s.children.map(serializeShape)]
  return { "p:grpSp": body, ":@": {} }
```

### buildXfrm

```
buildXfrm(position, size):
  return {
    "a:xfrm": [
      { "a:off": [], ":@": { "@_x": String(position.xEmu), "@_y": String(position.yEmu) } },
      { "a:ext": [], ":@": { "@_cx": String(size.cxEmu), "@_cy": String(size.cyEmu) } },
    ],
    ":@": {},
  }
```

`buildXfrmWithChild` for groups inserts the captured `<a:chOff>` /
`<a:chExt>` slice between `<a:off>` and `<a:ext>` — wait, OOXML's
ordering inside `<a:xfrm>` for a group is: `off, ext, chOff, chExt`.
We follow that exact order, splicing in the captured `chOffExtRaw.subtree` after
`<a:ext>`.

## serializeTxBody

```
serializeTxBody(tb):
  body = []
  if tb.bodyPrRaw:   body.push(tb.bodyPrRaw.subtree[0])
  if tb.lstStyleRaw: body.push(tb.lstStyleRaw.subtree[0])
  for p in tb.paragraphs:
    body.push(serializeTextParagraph(p))
  return { "p:txBody": body, ":@": {} }
```

## serializeTextParagraph

```
serializeTextParagraph(p):
  body = []
  pPr = serializePPr(p.properties)
  if pPr: body.push(pPr)
  for r in p.runs:
    body.push(serializeTextRun(r))
  if p.endParaRPrRaw: body.push(p.endParaRPrRaw.subtree[0])
  return { "a:p": body, ":@": {} }
```

## serializePPr

Build `<a:pPr>` if any typed property is set or any opaque prop survived.
Order of attributes: `marL, marR, lvl, indent, algn, …` — we only emit
`@_lvl` and `@_algn` from typed fields and append opaque attrs verbatim.

If `properties` has no typed fields and no `opaqueProps`, omit `<a:pPr>`
entirely.

## serializeTextRun

```
serializeTextRun(r):
  if r.isLineBreak:
    return { "a:br": serializeRPrChildren(r.properties), ":@": serializeRPrAttrs(r.properties) }
  rPr = serializeRPr(r.properties)
  body = []
  if rPr: body.push(rPr)
  body.push({ "a:t": [{ "#text": r.text }], ":@": (r.text.startsWith(" ") || r.text.endsWith(" ")) ? { "@_xml:space": "preserve" } : {} })
  return { "a:r": body, ":@": {} }
```

## serializeRPr

Build `<a:rPr>` with attributes:

| Field                       | Output attribute              |
| --------------------------- | ----------------------------- |
| `bold === true`             | `@_b="1"`                     |
| `bold === false`            | `@_b="0"`                     |
| `italic`                    | `@_i="1"` / `@_i="0"`         |
| `underline === string`      | `@_u="..."`                   |
| `underline === false`       | `@_u="none"`                  |
| `strike === true`           | `@_strike="sngStrike"`        |
| `fontSizeHundredths`        | `@_sz="3200"`                 |

And children:

| Field                       | Child element                                   |
| --------------------------- | ----------------------------------------------- |
| `fontFamily`                | `<a:latin typeface="..."/>`                     |
| `color`                     | `<a:solidFill><a:srgbClr val="..."/></a:solidFill>` |
| `opaqueProps[]`             | each entry's subtree appended verbatim          |

If no typed field is set and `opaqueProps` is empty, omit `<a:rPr>`.

## serializeContentTypesXml

Re-emit the content-types tree with the typed `<Default>` and
`<Override>` collections from `container.contentTypes`. Untouched
defaults (e.g. for `xml`, `rels`) survive unchanged.

When `add-slide` mints a new slide part, the serializer adds a
`<Override PartName="/ppt/slides/slideN.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` entry.

When `delete-slide` removes a slide part, the corresponding `<Override>`
is dropped.

When `insert-image` adds a media file with a new extension, the
serializer registers a `<Default Extension="…" ContentType="image/…"/>`
entry if not already present.

## serializeRelsXml

For each dirty rels part, walk `container.relationships.get(path)` and
emit a `<Relationships>` document with `<Relationship Id Type Target [TargetMode]>` entries in the recorded order.

## Whitespace and the byte-stability caveat

For untouched parts the container re-emits the original bytes. Byte stability is guaranteed.

For touched parts, `fast-xml-parser`'s builder may differ from the original byte representation in inconsequential ways: attribute order inside an element, single vs double quotes around attribute values, whether a self-closing tag is `<x/>` or `<x></x>`. **PowerPoint, Keynote, LibreOffice, and Google Slides all accept either form.**

The acceptance criterion is therefore:

- Untouched parts: **byte-identical** (SHA-256 match).
- Touched parts: **structurally equivalent and PowerPoint-accepted** (verified by re-parsing the output and asserting model equality).

## Determinism

The serializer is pure — no RNG, no `Date.now()`. Output for a given
snapshot is bit-stable across runs. (RNG only enters at command time
via `mintNodeId()`, and tests inject a deterministic counter.)

## Errors

Every serializer step throws structured `PptxSerializeError` on failure (missing part, dangling rel, invalid model invariant, etc.). Never silent.
