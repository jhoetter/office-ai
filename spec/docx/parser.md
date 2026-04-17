# DOCX — Parser

## Algorithm

```
parseDocx(buffer):
  container = OoxmlContainer.load(buffer)
  documentXml = container.readText("word/document.xml")
  documentTree = parseXml(documentXml)
  rootAttrs = extractRootAttrs(documentTree, "w:document")
  bodyTree = findChild(documentTree, "w:document").children → "w:body"
  body = parseBody(bodyTree, container, ctx)
  comments = parseComments(container)
  return DocxSnapshot {
    format: "docx",
    revision: 0,
    root: DocxDocument {
      id: mintNodeId(),
      body, comments,
      documentRootAttrs: rootAttrs,
      containerSerial: 0,
    },
    partHashes: container.allHashes(),
    container,   // attached to the snapshot for the serializer
  }
```

## parseBody

Walk the children of `w:body` in order. For each child element:

- `w:p` → `parseParagraph`
- `w:tbl` → `Table { kind: "table", id, raw: capture(child) }`
- `w:sectPr` → `SectionBreak { kind: "section-break", id, raw: capture(child) }`
- anything else → `OpaqueBlock { kind: "opaque-block", id, raw: capture(child) }`

Where `capture(node)` snapshots `{ tag, attrs, subtree }` from the
fast-xml-parser preserveOrder representation.

## parseParagraph

```
parseParagraph(pNode):
  pPr = findChild(pNode, "w:pPr")     // optional
  properties = parseParagraphProperties(pPr)
  children = parseInlineSiblings(pNode, except="w:pPr")
  if children.length === 0:
    children = [emptyRun(properties.runProps)]
  return Paragraph { id: mintNodeId(), properties, children }
```

## parseParagraphProperties

Read recognized children into the typed fields. Anything not recognized
goes into `opaqueProps[]`.

The paragraph's terminal `w:rPr` (the paragraph mark's run properties) is
treated as opaque and stored under `opaqueProps[]` with tag `w:rPr`.

## parseInlineSiblings

For each remaining child of `w:p`:

- `w:r` → `parseRun`
- `w:hyperlink` → `parseHyperlink`
- `w:ins` / `w:del` → `parseRevisionWrapper`
- `w:commentRangeStart` / `w:commentRangeEnd` → range marker nodes
- `w:permStart`, `w:bookmarkStart`, `w:proofErr`, `w:fldSimple`, etc. → `OpaqueInline`

## parseRun

```
parseRun(rNode):
  rPr = findChild(rNode, "w:rPr")
  properties = parseRunProperties(rPr)
  children = []
  for child in rNode.children except "w:rPr":
    if child.tag === "w:t" or "w:delText":
      text = collectText(child)
      preserve = child.attrs["xml:space"] === "preserve"
      children.push({ kind: "text", id: mintNodeId(), text, xmlSpacePreserve: preserve })
    elif child.tag === "w:br":
      children.push({ kind: "break", id: mintNodeId(), breakType: child.attrs["w:type"] })
    elif child.tag === "w:tab":
      children.push({ kind: "tab", id: mintNodeId() })
    elif child.tag === "w:drawing":
      children.push({ kind: "drawing", id: mintNodeId(), raw: capture(child) })
    else:
      children.push({ kind: "opaque", id: mintNodeId(), raw: capture(child) })
  return Run { id: mintNodeId(), properties, children }
```

## parseHyperlink

```
parseHyperlink(hNode):
  rId = hNode.attrs["r:id"]
  anchor = hNode.attrs["w:anchor"]
  children = hNode.children where tag === "w:r" → parseRun
  // Non-run children inside a hyperlink (rare; e.g. nested commentRangeStart)
  // are wrapped as Run with a single opaque child to keep typing tight; tracked
  // in build log if encountered.
  return Hyperlink { id: mintNodeId(), relationshipId: rId, anchor, children }
```

## parseRevisionWrapper

```
parseRevisionWrapper(node):
  return RevisionWrapper {
    kind: "revision",
    id: mintNodeId(),
    revisionType: node.tag === "w:ins" ? "ins" : "del",
    author: node.attrs["w:author"] ?? "",
    date:   node.attrs["w:date"] ?? "",
    revisionId: node.attrs["w:id"] ?? "",
    children: parseInlineSiblings(node),
  }
```

## parseComments

```
parseComments(container):
  if !container.parts.has("word/comments.xml"): return []
  tree = parseXml(container.readText("word/comments.xml"))
  commentsRoot = findChild(tree, "w:comments")
  return commentsRoot.children
    .filter(tag === "w:comment")
    .map(c => DocxComment {
      id:       c.attrs["w:id"],
      author:   c.attrs["w:author"] ?? "",
      initials: c.attrs["w:initials"],
      date:     c.attrs["w:date"] ?? "",
      body:     parseBody({children: c.children}, container, ctx),
    })
```

## Errors

The parser is **strict**:

- Missing `word/document.xml` → throw `DocxParseError("missing-main-part")`.
- XML parse failure → throw `DocxParseError("invalid-xml", { part, cause })`.
- Unrecognized element under `w:body` is **not** an error — it becomes opaque.

All errors are structured (`name`, `code`, `cause`, `partPath?`). Never
silent.

## Determinism

`mintNodeId()` is the only source of randomness. In tests we inject a
deterministic counter so snapshots are stable. The parser is otherwise pure.
