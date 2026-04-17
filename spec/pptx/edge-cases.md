# PPTX — Edge Cases

> Known hard cases, how we handle them, and what triggers a "this is
> opaque, edit it in PowerPoint" fallback.

## Slide-id accounting

`<p:sldId @id="…">` is **stable across reorders and never reused after
delete.** PowerPoint mints starting at 256 and increments
monotonically; surviving slides keep their original ids forever.

- `add-slide` picks `max(existing slideId) + 1`, lower bound `256`.
- `delete-slide` removes the entry from `<p:sldIdLst>` but does **not**
  free the id — `idGen.nextSlideId` is monotonic.
- `move-slide` reorders `<p:sldIdLst>` children. **Slide-ids stay put.**
  The slide originally at index 3 with id 261 moved to index 0 still
  has id 261.

This matches PowerPoint's behavior. Any "renumber" attempt would
invalidate every internal reference (transitions, hyperlinks targeting
slide ids, custom-show definitions, etc.).

## Slide part-path stability

The OOXML zip path of a slide is also stable across reorders. Slide 1
in display order is **not** required to be `slides/slide1.xml`. PowerPoint
sometimes leaves gaps (`slide1.xml`, `slide3.xml`, `slide7.xml`) after
deletes/moves. We honor that:

- `delete-slide` drops the part path; it is **not** reused.
- `add-slide` mints `slide{nextSlidePartIndex}.xml` (always strictly
  greater than any existing index).

## Notes-slide cleanup on delete

A slide may have an attached `notesSlide` referenced from its rels
(`Type=".../notesSlide"`). On `delete-slide`:

1. Walk the slide's rels graph for `notesSlide` rels.
2. For each: drop the notes-slide part, drop the notes-slide's own rels
   file, drop the corresponding `<Override>` in `[Content_Types].xml`.
3. Remove the notes-slide entry from `presentation.notesSlides`.

This prevents zombie notes parts that reference deleted slides.

## Picture media dedup

Two `insert-image` calls with identical bytes must NOT produce two
copies of the bytes in `ppt/media/`.

- The dedup index is keyed by SHA-256 of the bytes.
- It is built lazily on first `insert-image` call by walking
  `presentation.media`.
- If a media part already has the digest, the handler reuses it. It
  also tries to reuse an existing `Type=".../image"` rel on the same
  slide pointing to that part; if none exists it mints a new rel.

## Theme color resolution

Shape and run colors may be:

- `<a:srgbClr val="RRGGBB"/>` — literal hex. Used as-is.
- `<a:sysClr val="windowText" lastClr="000000"/>` — system color with
  a literal fallback; we use `lastClr`.
- `<a:schemeClr val="accent1"/>` — theme reference. Resolved at
  **render time** by walking slide → layout → master → theme. The
  model never materializes the resolved value — saving the file emits
  the original `<a:schemeClr>` reference unchanged.

For tints/shades / lumMod / lumOff: P0 ignores them in the rendered
preview (uses the base color). The original XML is preserved so the
file roundtrips perfectly when opened in PowerPoint.

## Placeholder inheritance

A placeholder shape (`<p:nvSpPr><p:nvPr><p:ph type="title">`) can omit
its `<a:xfrm>` and inherit position/size from the matching placeholder
in the layout (which may itself inherit from the master). Our model
stores **only what the slide's XML declares**:

- If the slide declares `<a:xfrm>` → typed `position`/`size` on the
  shape.
- If not → `position`/`size` are `undefined` and the renderer walks
  the layout/master chain on demand for visual fidelity.

`pptx:set-position` / `pptx:set-size` always emit a fresh `<a:xfrm>`
on the slide's shape. This is PowerPoint's behavior when the user
drags a placeholder: the slide takes ownership of the geometry.

## Group shapes — moving the whole group

Moving a `GroupShape` updates only the group's `<a:xfrm>` `<a:off>`. The
children's coordinates and the group's `<a:chOff>`/`<a:chExt>` (the
internal coordinate space) stay fixed. PowerPoint does the same — the
children "ride" the group transform.

We do not yet move children of a group individually (deferred). If a
caller passes a child shape's id to `set-position`, the handler resolves
it through the group recursion and rejects `not-applicable` with a
helpful message.

## Notes & Z-order

Slide shapes are stored in document order, which IS the z-order
(later = on top). The renderer paints in this order. `add-text-box`,
`insert-image`, and `add-slide` (which can clone layout placeholders)
append to the end → top of z-stack — same as PowerPoint's "Insert".

Z-order editing (bring forward / send to back) is deferred (not in the
prompt's listed P0 commands).

## Empty paragraphs and `<a:endParaRPr>`

OOXML requires every `<a:p>` to either contain at least one run or
have an `<a:endParaRPr>`. Our model preserves `endParaRPrRaw` per
paragraph; `set-text` rebuilds paragraphs but always re-attaches the
original first paragraph's `endParaRPrRaw` so the rebuilt text body
remains schema-valid.

## Unicode, line-breaks, and `<a:br>`

We treat `\n` in `set-text` as a paragraph separator. PowerPoint also
supports soft line breaks via `<a:br>` (Shift-Enter); the model
represents these as a `TextRun { isLineBreak: true, text: "" }`. The
renderer's HTML overlay translates a `<br>` element on input into a
`<a:br>` run on save. Bidirectional roundtrip preserved.

CJK / RTL text: stored byte-for-byte under `<a:t>` with no
normalization. The browser handles shaping in the HTML overlay; the
SVG `<text>` rendering relies on the system fonts. No bidi-mark
rewriting.

## XML entity escaping

`&`, `<`, `>`, and `"` inside `<a:t>` are escaped on serialize via
`fast-xml-parser`'s default entity handling. Already-escaped entities
in the source survive the parse → serialize cycle unchanged. We do
NOT normalize between numeric and named entities.

## Untouched-but-rewritten parts

`fast-xml-parser` may emit attribute order that differs from the
original on touched parts. This is fine for any consumer (PowerPoint,
LibreOffice, Keynote, Google Slides) and is documented in
`serializer.md` §Whitespace. The byte-identical guarantee applies only
to **untouched parts** (served from `OoxmlContainer.partsCache`).

## Things that legitimately abort the import

- Missing `ppt/presentation.xml` or `[Content_Types].xml` → import
  throws `PptxParseError("missing-main-part")`. Surface as a clear
  error in the UI; never crash silently.
- A slide whose `r:id` does not resolve in
  `presentation.xml.rels` → `PptxParseError("dangling-slide-rel")`. The
  document is structurally invalid; we don't try to "guess" the slide.
- A picture with `<a:blip @r:embed>` pointing nowhere →
  `PptxParseError("dangling-image-rel")`.

We do **not** abort on:

- Unknown shape elements (become `OpaqueShape`).
- Unknown root-level `<p:presentation>` children (preserved in
  `presentationOpaqueTail`).
- Unknown attributes (preserved on the carrying element via the parser's
  attribute-pass-through).

## "Open in PowerPoint" fallback

There is no current code path that surfaces a "fallback" toast — the
editor handles every shape kind by either editing it (typed) or
showing a labeled placeholder (opaque). Per [`prompt.md`](../../prompt.md)
line 175 we may add this for hard-failing import paths in P10 polish:

- A future "graceful-degradation" surface in `apps/web` could detect
  shapes whose `kind === "opaque"` and `tag === "p:graphicFrame"` (charts /
  SmartArt / tables) and offer a "Edit in PowerPoint" link. Out of
  scope this run; tracked in `docs/build-log/pptx.md`.
