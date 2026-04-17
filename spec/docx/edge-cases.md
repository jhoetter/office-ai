# DOCX — Edge Cases

> Known hard cases and how we handle them. Every entry has either a code
> path or a build-log deferral.

## Whitespace in `<w:t>`

OOXML uses `xml:space="preserve"` to keep leading/trailing whitespace.
fast-xml-parser's `trimValues: false` keeps it during parse. The model
stores `xmlSpacePreserve: boolean` per text leaf. The serializer emits
`xml:space="preserve"` only when the leaf was created with that attribute
or when the runtime detects leading/trailing whitespace in the text.

## Empty paragraphs and empty runs

Word represents an empty paragraph as `<w:p/>` (no children) or `<w:p><w:pPr/></w:p>`.
Our parser normalizes these to `Paragraph { children: [emptyRun] }` for
editing convenience. The serializer collapses `Paragraph { children: [] }`
to `<w:p/>` and `Paragraph { children: [emptyRun] }` to a single empty
`<w:r/>` — both Word-acceptable.

## Comment ID collisions

Existing `commentId`s are integers as strings (`"0"`, `"1"`, ...).
`docx:add-comment` mints `String(max(existing) + 1)`. We never reuse an
id that has appeared in any historical snapshot of the file (we don't
have history beyond load, but we do scan the loaded `word/commentsExtended.xml`
to avoid collisions there).

## Hyperlink relationship rewriting

When inserting a hyperlink (deferred to P1), we mint a fresh `rId`
through `RelationshipGraph.mintId()`. Existing hyperlinks are preserved
verbatim; their `rId`s are stable across save.

## Image content-type registration

Inserting an image of a MIME type not previously present in
`[Content_Types].xml` requires adding a `<Default>` entry for that
extension. P1.

## Run-property ordering inside `<w:rPr>`

OOXML's strict schema requires children of `<w:rPr>` to follow a defined
order. Our serializer emits typed children in that canonical order
(`rStyle → rFonts → b → bCs → i → iCs → u → strike → color → sz → szCs → highlight`)
followed by `opaqueProps[]` in the order encountered. This works for all
fixtures we tested; if a stricter parser complains, we promote the
problematic property out of opaque into typed.

## `<w:p>` paragraph mark `<w:rPr>`

The "paragraph mark" (the pilcrow's run properties) lives at the end of
`<w:pPr>` as a child `<w:rPr>`. We preserve it as an opaque entry under
`paragraph.properties.opaqueProps[]` and re-emit it inside `<w:pPr>` in
its original position.

## Unicode normalization

We **never** normalize Unicode (no NFC/NFD shifts). Text in/out of the
model is the same byte sequence. Editor input from a browser may already
be NFC; that's the user's choice. We don't second-guess it.

## BOM

Parts may start with a UTF-8 BOM. We preserve the BOM on untouched
parts. On dirty parts we re-emit without a BOM (Word accepts both).

## Unknown elements with `mc:Ignorable`

`mc:Ignorable="..."` lets a producer hint that listed elements may be
ignored by older parsers. We preserve the attribute on the root and
treat any element listed in `mc:Ignorable` exactly like any other
unrecognized element: opaque, byte-preserved.

## VBA-enabled documents (.docm)

`.docm` files include a `vbaProject.bin` part. We treat it as opaque
(do not parse, do not execute). Out of scope per [`prompt.md`](../../prompt.md)
line 224.

## `w:body`'s trailing `<w:sectPr>`

Word commonly puts a single `<w:sectPr>` at the very end of `<w:body>`
(the document-level section properties). We model this as a
`SectionBreak` block at the appropriate position, preserved verbatim.

## Nested tables

Tables are opaque this session (P1), so nested tables are also opaque
and round-trip cleanly.

## Read-only protection / digital signatures

`word/settings.xml` may declare write-protection. We **do not enforce**
it in this session — we save anyway. The build log warns. A real
production deployment should respect this flag.

## File size

The parser materializes the whole document tree in memory. For our 80%
target (<= 100 page documents) this is fine. Beyond 1000 pages, perf
work is required; not in this session.
