# PDF — Reference Analysis (Step A)

> Clean-room study of the open-source and proprietary PDF stack before
> writing a single line of our own code. Notes are derived from public
> documentation, public READMEs, ISO 32000-2, and the marketing
> material of the closed-source players. **No code from the references
> is copied or transliterated.** This file informs the spec; it is not
> the spec.

References surveyed (matches [`prompt-pdf.md`](../../prompt-pdf.md)
§Reference Repositories):

- `mozilla/pdf.js` (Apache 2.0) — canonical browser parser/renderer.
- `Hopding/pdf-lib` (MIT) — pure-TS writer with incremental save.
- `pdfcpu/pdfcpu` (Apache 2.0, Go) — page-level operations API surface.
- `embedpdf/embed-pdf-viewer` (MIT) — framework-agnostic viewer with a
  plugin/headless split.
- `anaralabs/lector` (MIT) — composable React headless toolkit.
- `chromium/pdfium` (BSD-3-Clause) via `@embedpdf/pdfium` WASM build.
- `ArtifexSoftware/mupdf.js` (AGPL — **architecture only, no code
  copied, never shipped**).
- ISO 32000-2:2020 / ISO 32000-1:2008 — canonical truth.
- Acrobat Web, Nutrient (PSPDFKit), Apryse (PDFTron), PDF Expert,
  ChatPDF, Shadow Reader — public docs and marketing only.

The DOCX/XLSX/PPTX clean-room work in [`spec/docx/analysis.md`](../docx/analysis.md)
and [`spec/pptx/analysis.md`](../pptx/analysis.md) is the most
important reference for _us_: PDF reuses `@officeai/core`
(`CommandBus`, `DocumentSnapshot`, `mintNodeId`, `freezeSnapshot`),
`@officeai/realtime` (codec + Y.js bridge), `@officeai/comments`
(thread model), and the same architectural invariants (headless first,
opaque-blob preservation, pending/approved/working tri-state,
agent-first CLI). PDF differs only where the format itself differs —
binary file with a cross-reference table instead of an OOXML zip,
incremental update instead of part-replacement, AP streams instead of
typed shape XML.

---

## 1. Format model

A PDF file is a sequence of indirect objects (`N gen R`) referenced by
a cross-reference table (`xref`) terminated by a trailer (`%%EOF`).
Each object is a number, string, name, dictionary, array, stream, or
indirect reference. The catalog (`/Catalog` in the trailer) anchors
the document tree:

```
Trailer → /Root → Catalog
  ├── /Pages          → Page tree (recursive /Pages and /Page nodes)
  ├── /Outlines       → Bookmark tree (linked-list dictionaries)
  ├── /Names          → Named destinations, embedded files, JS actions
  ├── /AcroForm       → Form fields + /XFA (deprecated)
  ├── /MarkInfo       → Tagged-PDF flag
  ├── /StructTreeRoot → Tagged-PDF structure tree (PDF/UA backbone)
  ├── /Metadata       → XMP RDF stream
  ├── /OpenAction     → JS action (we sandbox/no-op)
  └── /AA             → Additional actions (we sandbox/no-op)
```

Each `/Page` carries `/MediaBox`, `/CropBox`, `/Resources`
(`/Font`, `/XObject`, `/ColorSpace`, `/Pattern`, `/ExtGState`),
`/Contents` (one or more content streams), `/Annots` (annotation
references), and optional `/StructParents` linking back to the
struct tree.

### Decisions for our model

- We mirror this with a _typed projection_ (`PdfDocument` →
  `PdfPage[] + PdfOutline + PdfAnnotation[] + PdfFormField[] +
PdfAttachment[] + PdfMetadata + signatureCount`), **not** a faithful
  reconstruction of the object graph. The original byte buffer is
  always retained on the agent for incremental save. See
  [`document-model.md`](./document-model.md).
- The full object graph never enters our typed model — that road leads
  to the tar pit of re-implementing PDF semantics PDF.js / pdf-lib /
  PDFium have already nailed. We expose what the editor _operates on_
  (pages, annotations, fields, comments, metadata) and treat the rest
  as opaque bytes preserved verbatim by incremental save.

---

## 2. Rendering pipeline (PDF.js as canonical reference)

PDF.js separates the work across a worker and the main thread:

1. **Worker thread** parses the file, builds the cross-reference,
   resolves indirect references, decodes streams (FlateDecode,
   DCTDecode, LZWDecode, JBIG2Decode, JPXDecode, CCITTFaxDecode,
   ASCII85Decode, ASCIIHexDecode, RunLengthDecode), and walks each
   page's content stream to produce an **operator list** (`OPS.save`,
   `OPS.transform`, `OPS.setFont`, `OPS.showText`, `OPS.constructPath`,
   `OPS.fill`, `OPS.stroke`, `OPS.paintImageXObject`, …).
2. **Main thread** receives the operator list and replays it onto a
   `<canvas>` 2D context. Fonts arrive as raw FontFace blobs; PDF.js
   substitutes a system font when a Type3 / unmapped CMap font cannot
   be reconstructed.
3. **Text-layer overlay** is built from `getTextContent()` — invisible
   absolutely-positioned `<span>`s aligned to the canvas, enabling
   selection without re-rasterizing text.
4. **Annotation-layer overlay** renders interactive annotations
   (links, AcroForm widgets, free-text, popups) as DOM elements above
   the canvas.
5. **Struct-tree-layer overlay** renders the tagged structure tree as
   ARIA-labelled DOM, hidden visually but read by screen readers.

### What PDF.js gets approximately right (our wins)

- **Type3 fonts** — PDF.js substitutes; PDFium renders correctly. We
  trigger PDFium fallback when `/Font/Subtype /Type3` is detected.
- **DeviceN / NChannel / Separation color spaces** — PDF.js degrades
  to RGB approximation; PDFium renders with proper tint transforms.
- **Color-managed CMYK** — PDF.js ignores `/OutputIntent`; PDFium
  honors the embedded ICC profile.
- **Custom CMaps** — PDF.js sometimes glyph-substitutes; PDFium
  consistently maps via `/CIDSystemInfo`.

### Decisions for our renderer

- Default engine is **PDF.js** (smaller bundle, mature text-layer and
  AcroForm story, worker-thread isolation already done). See
  [`engine-strategy.md`](./engine-strategy.md).
- Lazy **PDFium-WASM** fallback via `@embedpdf/pdfium` for
  fidelity-critical signals. Selected by `selectEngine()` — pure
  function, no globals.
- Rendering is virtualized: only `±2` pages around the viewport are
  rasterized at any time, rest are placeholders sized from the page's
  user-units. See [`rendering-pipeline.md`](./rendering-pipeline.md).
- Raster cache uses an LRU with a memory budget (default 200 MB),
  evicting least-recently-painted page bitmaps when the budget is hit.
- `OffscreenCanvas` is used for thumbnail rasterization in a worker
  pool when available; the fallback is the main-thread `<canvas>`.

---

## 3. Edit / incremental-update model (pdf-lib as canonical reference)

PDF supports three save modes:

1. **Full re-serialize** — write a fresh xref table and trailer; all
   objects are renumbered; signatures break.
2. **Incremental update** — append new objects + a new xref section +
   a new trailer to the original bytes. Original bytes are byte-for-byte
   preserved. Signatures over the prior byte range remain valid.
3. **Linearization** — special-cased optimized layout for fast web
   view; not a save mode but a constraint.

`pdf-lib` supports both #1 and #2. pdfcpu (Go) is the canonical
reference for incremental update done correctly — its xref handling
is the most rigorous open-source implementation we surveyed.

### Decisions for our serializer

- **Default save: incremental update.** Append only the objects we
  changed (new annotations, mutated form-field values, mutated page
  rotation, new outline entries) plus a fresh `/Prev`-chained xref
  section. Original bytes — including signed regions — are
  byte-preserved.
- **Full re-serialize on opt-in** (`exportFile({ rewrite: true })`)
  for "save as cleaned copy" — strips orphaned objects and renumbers.
  Marked clearly as signature-breaking in the UI.
- **Page-level edits** (rotate, reorder, delete, insert, merge,
  split, crop, watermark, page-numbers, set-metadata) flow through
  pdf-lib in `packages/pdf-edit`. Each function takes/returns
  `Uint8Array`; pure; no DOM.
- **Annotations** are written natively (with valid AP streams) so they
  appear in Adobe Acrobat. See [`annotation-model.md`](./annotation-model.md).
- We **never** silently re-encode streams. A FlateDecode stream we
  didn't touch comes back byte-for-byte.

---

## 4. Annotation model

Native PDF annotations are dictionaries under each page's `/Annots`
array. Every annotation has a `/Subtype` (`Highlight`, `Square`,
`Ink`, `FreeText`, `Stamp`, `Link`, `Popup`, `Widget`, `Redact`, …),
a `/Rect`, an optional `/Contents`, an optional `/T` (author), and an
optional `/AP` (appearance stream — the visual representation).

PDF readers that respect the spec render the `/AP` stream when
present. Readers that don't (rare, but Apple Preview historically had
gaps) fall back to drawing from the typed fields. The robust strategy
is **always emit AP streams**, even when redundant.

FDF / XFDF are sidecar formats for shipping annotations between
applications. XFDF is XML; FDF is binary-ish. Both round-trip into
native annotations on import.

### Decisions for our annotation pipeline

- `packages/pdf-annotations` defines a **typed** `PdfAnnotation` union
  (highlight, underline, strikethrough, squiggly, sticky note,
  free-text, ink, line/arrow, rectangle, ellipse, polygon, polyline,
  stamp, link, redaction). See [`annotation-model.md`](./annotation-model.md).
- Each typed annotation has an **AP-stream emitter** producing the
  exact content-stream operators an Adobe-compatible viewer expects.
- **XFDF JSON I/O** for the agent: import XFDF → typed annotations →
  AP streams; export typed annotations → XFDF for round-trip with
  Acrobat / Foxit.
- Roundtrip test: annotate via our writer → reopen with PDF.js →
  annotation visible at the correct coordinates with the correct
  color/style.

---

## 5. Form / AcroForm model

`/AcroForm` is a flat dictionary holding the form-field tree (`/Fields`),
default appearance (`/DA`), default resources (`/DR`), and signature
flags (`/SigFlags`). Each field is one of: text (`/Tx`), button (`/Btn`
— checkbox, radio, push), choice (`/Ch` — combo, list), or signature
(`/Sig`).

XFA is a parallel form system layered on top of AcroForm. Adobe is
deprecating XFA. We render the AcroForm fallback if present and
display a clear "this form requires Adobe Acrobat" banner otherwise.

### Decisions for our form pipeline

- `packages/pdf-forms` enumerates widgets, gets/sets values, validates
  (regex + `/MaxLen`), and resolves calc-order. **No JS execution**
  (no `/AA`, no `/F`, no `/V` JS) — sandboxed/no-op.
- **Flatten on save** option bakes values into the page content stream
  - drops the field, producing a non-fillable but visually identical
    PDF. See [`form-engine.md`](./form-engine.md).
- **Signature fields** are detected, displayed read-only, and
  reported. We do not sign cryptographically (deferred). We do place
  visible signature images via the annotation pipeline.

---

## 6. Accessibility model (PDF/UA, ISO 14289-1)

Tagged PDF carries a `/StructTreeRoot` whose nodes carry `/S` (role:
`/H1`, `/P`, `/Figure`, `/Table`, …), `/Alt` (alt text), `/Lang`
(language), and `/ActualText`. The struct tree drives reading order
for screen readers. PDF.js renders this tree as ARIA-labelled hidden
DOM aligned with the text layer.

### Decisions for our a11y pipeline

- We expose the struct tree as a **typed `PdfStructTree`** that the
  viewer renders as hidden ARIA-labelled DOM aligned with the text
  layer. Screen readers (VoiceOver, NVDA) get a faithful reading
  experience.
- **Reflow mode** — single-column accessibility view that linearizes
  the struct tree and ignores layout. Font-size adjustable.
- Untagged PDFs degrade gracefully: text-layer reading order is used
  as a proxy.
- Viewer chrome targets WCAG 2.2 AA from day one.
- See [`accessibility.md`](./accessibility.md).

---

## 7. What competitors get wrong (our wins)

- **Chrome built-in (PDFium)** — no annotations, no comments, no
  outline editing, no dark mode, ugly default toolbar. We win on
  every read-side affordance.
- **Adobe Acrobat Web** — slow first paint, heavy chrome, locked
  behind sign-up gates for non-trivial features, privacy-hostile
  (everything uploads). We win on speed and privacy.
- **Most viewers' dark modes** invert images and figures along with
  the text. Shadow Reader is the only widely-known viewer that does
  smart-invert correctly. We bake smart-invert into the default. See
  [`dark-mode.md`](./dark-mode.md).
- **AI-native PDF chat tools (ChatPDF, Humata)** are upload-and-pray;
  citations are loose; the viewer is bolted-on. Our CLI is the AI
  surface — agents drive the same headless API the UI uses, with
  precise page+rect anchoring.
- **None expose a structured per-mutation diff** an agent can
  introspect/approve/reject. We inherit DOCX/XLSX/PPTX's tri-state
  approach.
- **None treat the agent as a first-class user.** Same gap as the
  sibling formats: agent API and CLI are primary; the viewer UI is a
  skin over the same headless agent.

---

## 8. Engine fallback decision

| Concern                    | PDF.js                                  | PDFium-WASM (`@embedpdf/pdfium`) | mupdf.js (study only) |
| -------------------------- | --------------------------------------- | -------------------------------- | --------------------- |
| License                    | Apache 2.0 ✓                            | Apache 2.0 (engine BSD-3) ✓      | AGPL ✗ shippable      |
| Bundle (gzipped)           | ≈350 KB                                 | ≈3.5 MB lazy                     | ≈4 MB                 |
| Worker isolation           | First-class                             | Yes (lazy boot)                  | Yes                   |
| Text-layer                 | Mature                                  | Available, lower-level           | Available             |
| AcroForm                   | Mature widget layer                     | Available, lower-level           | Available             |
| StructTree                 | Available                               | Available                        | Available             |
| Type3 fonts                | Substitutes                             | Renders                          | Renders               |
| DeviceN / Separation       | RGB approx                              | Correct                          | Correct               |
| Custom CMaps               | Sometimes substitutes                   | Correct                          | Correct               |
| ICC color management       | Limited                                 | Full                             | Full                  |
| Headless Node renderer     | `pdfjs-dist/legacy` + `@napi-rs/canvas` | Yes via WASM                     | Yes via WASM          |
| Streaming / range requests | Yes                                     | No                               | No                    |
| Time to first paint        | Best                                    | Worse (lazy)                     | Worse (lazy)          |

**Decision:** PDF.js as the default engine; lazy PDFium-WASM as a
fidelity fallback triggered by the `selectEngine()` heuristic
(uncommon color space, custom CMap, Type3 font, allowlist hit, or
explicit user opt-in). The interface in
[`packages/pdf-engine`](../../packages/pdf-engine/src/types.ts) is
engine-agnostic so a future swap to PDFium-primary is internal.
Detail: [`engine-strategy.md`](./engine-strategy.md).

---

## 9. Edge-case landscape

Documented in [`edge-cases.md`](./edge-cases.md). Highlights that
shape the spec:

- **Encrypted PDFs** — password prompt; in-memory decrypt; never
  persist the password.
- **Broken xref** — pdf-lib + PDF.js both attempt recovery. We surface
  a `pdf:repair-attempted` warning so the user knows.
- **Signed-then-modified** — incremental save preserves the signed
  byte range; full re-serialize warns and breaks the signature
  intentionally.
- **Linearized PDFs** — preserved on incremental save (we don't
  re-linearize); flagged on full re-serialize.
- **Custom CMaps** — auto-trigger PDFium fallback.
- **Very wide / very long pages** — virtualized rendering caps the
  rasterized area at the viewport `+ ±2 pages` budget.
- **PDFs with embedded JavaScript** — sandboxed; never executed.
  Surfaced as a banner.

---

## 10. What's missing from the references that we add

- **Headless agent + CLI as the AI surface.** No surveyed viewer
  exposes its full read+edit surface as a `office-agent pdf …` CLI
  with JSON-by-default output. See [`cli.md`](./cli.md).
- **Per-mutation diff + tri-state approval** for agent-source edits.
  Inherited from the sibling formats.
- **Smart-invert dark mode by default**, not an opt-in afterthought.
- **Engine-agnostic abstraction with a documented migration path**
  (PDF.js today, PDFium-primary spec'd-and-stubbed for tomorrow).

---

## Summary

The PDF stack is mature: PDF.js solves the read path; pdf-lib solves
the incremental write path; PDFium fills the fidelity gaps PDF.js
can't close. Our value is _not_ in re-implementing any of these — it
is in (a) wrapping them behind a clean engine abstraction, (b)
exposing a typed projection over the binary that flows through the
shared `CommandBus`, (c) shipping the AI integration surface as a
first-class CLI, and (d) hitting the 10/10 reading-experience bar
with smart-invert dark mode, virtualized smooth scroll, full
keyboard reachability, and tagged-PDF accessibility from day one.

Next: produce the rest of `spec/pdf/*` (feature-scope, engine-strategy,
document-model, rendering-pipeline, text-layer, annotation-model,
form-engine, editing-pipeline, accessibility, dark-mode, search,
agent-commands, cli, keyboard-shortcuts, performance, edge-cases,
acceptance-criteria).
