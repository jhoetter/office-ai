# DOCX — Serializer

## Goal

Produce an OOXML zip such that:

1. Every part **not touched** by any command since the last load is
   byte-identical to the loaded bytes.
2. Every part **touched** by commands is re-serialized from the model in a
   shape Word/LibreOffice/Google accept without a repair dialog.
3. Opaque blocks are re-emitted verbatim (their fast-xml-parser subtree
   round-trips losslessly).

## Algorithm

```
serializeDocx(snapshot):
  container = snapshot.container.clone()

  if bodyDirty(snapshot):
    documentXml = serializeDocumentXml(snapshot.root)
    container.writeText("word/document.xml", documentXml)

  if commentsDirty(snapshot):
    if snapshot.root.comments.length > 0:
      ensureCommentsPart(container)
      commentsXml = serializeCommentsXml(snapshot.root.comments)
      container.writeText("word/comments.xml", commentsXml)
    else:
      container.removePart("word/comments.xml")
      removeRel(container, type: ".../comments")

  return container.serialize()
```

## bodyDirty / commentsDirty

The command bus marks the snapshot's `dirty` flags when it dispatches a
command that modifies the body or comments. The serializer trusts these
flags. Default for a freshly-parsed snapshot: `dirty = false` everywhere
→ output is byte-identical to input.

## serializeDocumentXml

```
serializeDocumentXml(doc):
  bodyChildren = []
  for block in doc.body:
    bodyChildren.push(serializeBlock(block))
  bodyChildren.push(serializeTrailingSectPr(doc))   // section properties at the end of body
  documentTree = [
    { "?xml": [], ":@": { "@_version": "1.0", "@_encoding": "UTF-8", "@_standalone": "yes" } },
    { "w:document": [{ "w:body": bodyChildren }], ":@": doc.documentRootAttrs },
  ]
  return serializeXml(documentTree)
```

## serializeBlock

| Node                  | Output                                |
| --------------------- | ------------------------------------- |
| `Paragraph`           | `serializeParagraph(p)`               |
| `Table { raw }`       | the captured subtree, unchanged       |
| `SectionBreak { raw }`| the captured subtree, unchanged       |
| `OpaqueBlock { raw }` | the captured subtree, unchanged       |

## serializeParagraph

```
serializeParagraph(p):
  pNode = { "w:p": [], ":@": {} }
  pPr = serializeParagraphProperties(p.properties)
  if pPr: pNode["w:p"].push(pPr)
  for child in p.children:
    pNode["w:p"].push(serializeInline(child))
  return pNode
```

## serializeParagraphProperties

Build `<w:pPr>` from typed fields **plus** opaqueProps in the order they
were originally encountered (preserveOrder discipline). Strategy: emit
typed children first in OOXML's canonical order (style → numPr → ind →
spacing → jc), then append opaqueProps verbatim. A small follow-up may
need to interleave them more carefully for Word's stricter parsers; the
current strategy is documented in `edge-cases.md`.

## serializeRun

```
serializeRun(r):
  rNode = { "w:r": [], ":@": {} }
  rPr = serializeRunProperties(r.properties)
  if rPr: rNode["w:r"].push(rPr)
  for child in r.children:
    rNode["w:r"].push(serializeRunChild(child))
  return rNode
```

`serializeRunChild`:

| Child               | Output |
| ------------------- | ------ |
| `text { text, xmlSpacePreserve }` | `<w:t xml:space="preserve">text</w:t>` if preserve else `<w:t>text</w:t>` |
| `break { breakType }`             | `<w:br w:type="…"/>` (no breakType → `<w:br/>`) |
| `tab`                             | `<w:tab/>` |
| `drawing { raw }`, `opaque { raw }` | the captured subtree, unchanged |

## serializeHyperlink

```
serializeHyperlink(h):
  hNode = { "w:hyperlink": h.children.map(serializeRun), ":@": {} }
  if h.relationshipId: hNode[":@"]["@_r:id"] = h.relationshipId
  if h.anchor: hNode[":@"]["@_w:anchor"] = h.anchor
  return hNode
```

## serializeRevisionWrapper

```
serializeRevisionWrapper(rev):
  tag = rev.revisionType === "ins" ? "w:ins" : "w:del"
  attrs = { "@_w:id": rev.revisionId, "@_w:author": rev.author, "@_w:date": rev.date }
  return { [tag]: rev.children.map(serializeInline), ":@": attrs }
```

## serializeCommentsXml

```
serializeCommentsXml(comments):
  root = { "w:comments": comments.map(serializeComment), ":@": COMMENTS_ROOT_ATTRS }
  return serializeXml([XML_DECL, root])

serializeComment(c):
  return {
    "w:comment": c.body.map(serializeBlock),
    ":@": {
      "@_w:id": c.id,
      "@_w:author": c.author,
      "@_w:date": c.date,
      ...(c.initials ? { "@_w:initials": c.initials } : {}),
    },
  }
```

## ensureCommentsPart

When adding the first comment to a doc that has no comments part:

1. Add `word/comments.xml` with `COMMENTS_ROOT_ATTRS` (the standard `xmlns:w` set).
2. Add a relationship to `word/_rels/document.xml.rels`:
   ```
   <Relationship Id="rIdN" Type=".../comments" Target="comments.xml"/>
   ```
3. Add the override to `[Content_Types].xml`:
   ```
   <Override PartName="/word/comments.xml"
             ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
   ```

These three updates are the only `[Content_Types].xml` edits performed
this session.

## Whitespace and the byte-stability caveat

For untouched parts the container re-emits the original bytes. Byte
stability is guaranteed.

For touched parts (typically just `word/document.xml` and possibly
`word/comments.xml`), `fast-xml-parser`'s builder may differ from the
original byte representation in inconsequential ways: attribute order
inside an element, single vs double quotes around attribute values,
whether a self-closing tag is `<x/>` or `<x></x>`. **Word and LibreOffice
both accept either form.**

The acceptance criterion is therefore:

- Untouched parts: **byte-identical** (SHA-256 match).
- Touched parts: **structurally equivalent and Word-accepted** (verified by
  re-parsing the output and asserting model equality).

## Errors

Every serializer step throws structured `DocxSerializeError` on failure
(missing part, invalid model invariant, etc.). Never silent.
