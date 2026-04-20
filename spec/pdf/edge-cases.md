# PDF — Edge Cases

> Known hard cases, how we handle them, and what triggers a fallback
> or warning.

Cross-references: engine fallback in
[`engine-strategy.md`](./engine-strategy.md);
serializer constraints in [`editing-pipeline.md`](./editing-pipeline.md);
form caveats in [`form-engine.md`](./form-engine.md);
performance under stress in [`performance.md`](./performance.md).

## Encrypted PDFs

PDF supports user-password (decrypt-to-read) and owner-password
(restrict-permissions). Both encrypt the document body; the trailer
declares `/Encrypt` referring to the encryption dictionary.

**Handling:**

- On parse, the engine throws a typed `PasswordRequired` error
  (PDF.js: `PasswordException` with `code === 1`).
- The viewer surfaces a modal: _"This document is password-protected.
  Enter the password to decrypt."_
- On submit, `PdfAgent.fromBuffer(bytes, { password })` re-attempts.
- The password is held in memory for the session and **never
  written**. Closing the document discards it.
- The CLI accepts `--password <pw>` on every command; missing
  password on encrypted input exits with code `1` and the message
  _"password required for this PDF; pass --password"_.
- The model's `PdfMetadata.encryption` flags both `hasUserPassword`
  and `hasOwnerPassword` so the UI shows the security pane.
- Permission flags (`/P` bits — print, copy, modify, …) are
  surfaced to the UI but **not enforced** by us; we are the document
  owner once decrypted. We surface the permissions banner so users
  know the original author intended restrictions.

## Corrupt files / broken xref

Common in scanner output and emails truncated mid-attachment.

**Handling:**

- The engine attempts xref recovery automatically (PDF.js does this
  by linearly scanning for `obj` markers; PDFium does the same).
- On successful recovery, the snapshot carries
  `warning: "xref-recovered"` so the UI can show a "this file was
  recovered; consider re-saving" banner.
- On full failure, the agent throws `error: "parse-failed"` with the
  underlying engine message. The CLI exits with code `2`.
- Recovered files saved through us emit a clean xref via incremental
  save. Content is preserved best-effort.

## Very wide / very long pages

Engineering drawings, cartograms, banner PDFs.

**Handling:**

- Page placeholders are sized from `PdfPage.width × PdfPage.height`,
  so the scrollbar is correct.
- Bitmap rasterization is **clamped** by the canvas size limit
  (browser-dependent, ~16k×16k). The renderer scales the bitmap
  request down and overlays the text layer at the user's requested
  zoom — text remains crisp; only the canvas image scales.
- A subtle banner appears at zoom levels that triggered clamping:
  _"Rendering at reduced fidelity due to page size."_

## Missing fonts

PDFs may reference fonts that are neither embedded nor available to
the system (deeply rare but real).

**Handling:**

- PDF.js substitutes a system font of the same general category
  (serif → Times, sans → Helvetica, mono → Courier).
- We surface this in `office-agent pdf list-fonts`:
  `{ "name": "FooBar", "embedded": false, "substituted": true,
"substitutedAs": "Helvetica" }`.
- For pages with substituted fonts, the substituted glyphs may be
  visually different from the original. The viewer shows a per-page
  badge (gear icon) that opens a popover listing the substitutions.
- The user can manually trigger the PDFium fallback (which has its
  own font substitution and may render the document differently).

## CMap edge cases

Custom or non-standard CMaps map character codes to glyphs in ways
PDF.js sometimes glyph-substitutes incorrectly (the rendered glyph
is right, but the underlying Unicode codepoint is wrong, breaking
search and copy).

**Handling:**

- `selectEngine()` detects custom CMaps via `/CIDSystemInfo` not in
  the standard registry and switches to PDFium for that document.
- For partial cases (e.g. one font has a custom CMap, others are
  standard), PDF.js handles the standard fonts and we accept the
  fidelity hit on the non-standard one. A future polish item is
  per-font engine routing.

## Signed-then-modified

A PDF may carry one or more signatures over a `/ByteRange` of the
file. Our incremental-save preserves the signed bytes; the
signature stays valid.

**Handling:**

- On parse, signatures are detected and surfaced in
  `signatureCount` + the signature panel.
- Edits via incremental save (the default) are appended; the
  signature's byte range is unchanged. The signature stays valid for
  the original revision; readers that walk `/Prev` correctly show
  the document as "signed at revision 1, modified at revision 2".
- Edits via `exportFile({ rewrite: true })` (full re-serialize) emit
  a `warning: "signature-broken-on-rewrite"` and a confirmation
  modal. The user must explicitly opt in. After the rewrite the
  signature pane shows "broken — re-sign required".
- `office-agent pdf list-signatures` reports `valid: true | false`
  per signature and indicates the byte-range coverage.

## Linearized PDFs

Optimized for HTTP range-request streaming; the first page's xref
section appears near the file start.

**Handling:**

- On parse, `metadata.linearized` is set.
- Incremental save preserves the linearization hint at the file
  start; the appended delta is at EOF as usual. The file remains
  byte-correct but no longer linearization-optimized for the new
  delta. We do not re-linearize on save (re-linearization requires
  a full rewrite, which would break signatures).
- A future polish item: `office-agent pdf optimize --linearize`
  emits a re-linearized output.

## Embedded JavaScript

PDFs may carry JS in `/Catalog/OpenAction`, `/Catalog/AA`,
form-field `/AA`, or per-page `/AA`.

**Handling:**

- We **do not execute** any embedded JS. Sandboxed / no-op.
- On parse, the snapshot carries `warning: "embedded-js"` if any JS
  is detected. The UI shows a banner: _"This document contains
  embedded scripts that have been blocked for your safety."_
- Form calc-order resolution (which would normally use JS) falls
  back to the static dependency analysis described in
  [`form-engine.md`](./form-engine.md).
- The CLI's `inspect` output flags `"hasJavaScript": true`.

## Pure-XFA forms

A PDF whose form is XFA-only (no AcroForm fallback) cannot be filled
in our viewer.

**Handling:**

- The viewer shows a banner: _"This form was authored in XFA
  dynamic-forms format and requires Adobe Acrobat. The static
  layout is shown for reference only."_
- Field-level interaction is disabled.
- `list-form-fields` flags `"backing": "xfa-only"`.
- A best-effort `office-agent pdf convert-xfa-to-acroform` is on the
  roadmap (P2); not in this session.

## Very large files (>200 MB)

- The original buffer is held in memory for incremental save —
  unavoidable given the byte-preservation invariant.
- For 200 MB files the viewer behaves correctly but memory peaks at
  ~1.5 GB (see [`performance.md`](./performance.md)).
- For files > 500 MB, the viewer surfaces a confirmation:
  _"This document is large (X MB). Open anyway?"_ — the user can
  decline and the file is not loaded.
- The CLI handles arbitrary sizes; OOM is the OS's problem.

## Thousands of annotations

Some legal-review documents carry thousands of highlights and
comments.

**Handling:**

- The annotations panel virtualizes its list (only visible rows are
  in the DOM).
- Per-page annotation overlay only mounts overlays for the visible
  pages (window ±2).
- A spatial index per page (R-tree) accelerates hit testing
  regardless of count.
- The agent's `applyCommands([…1000 highlights…])` batches the
  underlying pdf-lib writes into a single incremental save pass.

## Non-Latin scripts (Arabic / CJK / Cyrillic / Hebrew / Thai)

- **Arabic / Hebrew (RTL)**: bidi reordering applied to the plain-
  text projection (UAX #9). See
  [`text-layer.md` § RTL and bidi](./text-layer.md).
- **CJK**: handled via the engine's CMap path; vertical writing
  mode detected from the text matrix `b` component. CJK pages with
  custom CMaps trigger PDFium fallback.
- **Cyrillic**: handled identically to Latin.
- **Thai**: combining marks render via the engine; selection works
  per character cluster (not per codepoint).

## Scanned PDFs without OCR

`hasTextLayer === false` for every page.

**Handling:**

- Search returns empty.
- Selection drags don't paint highlight (no text underneath the
  pointer).
- The viewer shows a per-page banner: _"Run OCR (German + English)
  to make this page selectable and searchable."_
- One click triggers `packages/pdf-ocr.addTextLayer(buffer,
[pageIndex], "deu+eng")`. Output replaces the snapshot's bytes via
  incremental save.

## Files with attachments

`/Names/EmbeddedFiles` holds attached files (auxiliary spreadsheets,
related PDFs, etc.).

**Handling:**

- Surfaced in the Attachments sidebar tab (read-only).
- `office-agent pdf list-attachments --out ./attachments/` writes
  each attachment to disk.
- We do not currently support adding/removing attachments via the
  bus; CLI-only `office-agent pdf attach` is on the roadmap.

## PDFs with custom CMaps

See "CMap edge cases" above. `selectEngine()` triggers PDFium.

## PDF/A files

`/Metadata` declares conformance via `<pdfaid:part>` and
`<pdfaid:conformance>`.

**Handling:**

- We **open** PDF/A files normally.
- We **do not** certify outputs as PDF/A — saves remove the PDF/A
  badge. The viewer shows a _"This was a PDF/A file; saving will
  drop PDF/A conformance"_ banner before the first edit.
- A best-effort `office-agent pdf convert-to-pdfa` is roadmap (P2).

## Files with broken / overflowing /MediaBox

Some legacy PDFs declare a `/MediaBox` that doesn't match the
content stream's drawing extent.

**Handling:**

- We honor `/MediaBox` for layout (it's what every other reader does).
- Content drawn outside `/MediaBox` is clipped during render (the
  engine handles this).
- We never silently rewrite `/MediaBox` to match content.

## Mixed orientation pages

A document where pages 1-10 are portrait and pages 11-20 are
landscape.

**Handling:**

- Each page placeholder is sized from its own `width × height`.
- The continuous-scroll layout adapts.
- Two-up view aligns pages by the taller of the pair to keep them
  vertically centered.

## Files with no `/Catalog`

Malformed / truncated.

**Handling:**

- The engine throws on parse; surfaced as `error: "parse-failed"`.
- No recovery is attempted (without a catalog, there's nothing to
  recover).
