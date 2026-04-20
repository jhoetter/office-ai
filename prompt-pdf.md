# Build a Best-in-Class, AI-Native Online PDF Viewer & Editor

## Mission

You are a senior software architect and engineer. You will autonomously complete a full build of a browser-embeddable, AI-native PDF viewer and light editor — slotting in next to the existing DOCX, XLSX, and PPTX editors in this monorepo — in one continuous session.

The bar: the result must feel like the **best online PDF viewer on the public internet**. Faster than Chrome's built-in viewer, better looking than Adobe Acrobat Web, more capable than Preview.app, more agent-friendly than ChatPDF, and respectful of user privacy in a way none of the proprietary players are.

Work in this exact sequence, without skipping ahead:

1. Spec PDF (analysis → spec) → 2. Build PDF → 3. Validate PDF → 4. Polish to a 10/10 reading experience

Do not start Step 2 until Step 1 is complete. Do not declare Step 4 done without measurable evidence (Lighthouse, performance traces, fixture roundtrip, accessibility audit).

---

## Non-Negotiable Quality Bar

**Visual Fidelity & Roundtrip Integrity.** Every viewer/editor operation must satisfy this bar:

- Open any real-world PDF (PDF 1.4 → PDF 2.0, ISO 32000-2) produced by Adobe Acrobat, Microsoft Word "Save as PDF", LaTeX (pdfTeX/XeLaTeX/LuaTeX), Chrome "Print to PDF", macOS Quartz, LibreOffice, Google Docs export, scanners (with embedded OCR), and InDesign.
- Render pixel-faithfully against a known-good baseline (PDFium / Adobe rendering).
- For any non-destructive operation (annotate, fill form, add page, rotate, redact), save back to PDF.
- Reopen in Adobe Acrobat Reader, Preview.app, and Chrome's built-in viewer — must show no "this file was modified / repair needed" warnings.
- Every PDF object the editor did not touch must be **byte-preserved** in the output (incremental update strategy preferred for edits).
- No corruption, no silent data loss, no broken signatures (unless a signed region was deliberately edited, in which case the user is warned first).

This is the only acceptance criterion that cannot be traded away. Everything else is scope.

---

## Legal Constraint (Clean-Room Approach)

You will analyze reference repositories to extract concepts, patterns, and architectural decisions. You will then build a fresh implementation from a specification you derive — not a fork, not a verbatim port.

**Allowed:** Study public code, extract architecture concepts, describe behavior and algorithms at the conceptual level, implement independently from first principles + the ISO 32000-2 spec. You are absolutely allowed to use open source libraries for building, e.g. MIT, Apache, BSD, or LGPL with dynamic linking.

**Not allowed:** Copy code verbatim, lightly rename identifiers, ship any AGPL-licensed component as a runtime dependency to end users (study only), import reference repos that are AGPL/GPL as packages.

**Runtime dependencies permitted** (MIT / Apache 2.0 / BSD only):

- `pdfjs-dist` (Apache 2.0) — Mozilla PDF.js parser/renderer, the proven core for browser PDF rendering
- `pdf-lib` (MIT) — pure-TS PDF write/modify (page ops, form fill, basic annotations, incremental save)
- `@embedpdf/pdfium` (Apache 2.0, via PDFium) — WebAssembly build of PDFium for cases where PDF.js falls short (color spaces, exotic fonts, fidelity-critical rendering); used as an optional fallback engine
- `pdfjs-serverless` (Apache 2.0) — for headless Node.js rendering in CI / agent contexts
- `tesseract.js` (Apache 2.0) — OCR for scanned PDFs (lazy-loaded, on demand)
- `fontkit` / `@pdf-lib/fontkit` (MIT) — font subsetting/embedding when adding text
- Y.js (MIT) — CRDT collaboration primitives (already used by sibling editors)
- Zod (MIT) — runtime schema validation
- Any other MIT/Apache/BSD library if justified in the spec

**Forbidden as runtime deps (study/inspiration only):**

- MuPDF.js (AGPL) — read the architecture, do not ship
- PPTist's renderer (AGPL) — irrelevant here, listed for symmetry
- Any "free trial" SDK requiring a license key (PSPDFKit/Nutrient, Apryse/PDFTron, Foxit) — study marketing material and public docs for feature inspiration only

---

## Reference Repositories & Products to Study

Study these before speccing. Read architecture docs, main source files, public design write-ups. Do not copy. Understand.

### Open-source viewers & libraries

- https://github.com/mozilla/pdf.js (Apache 2.0) — **the canonical reference**. Two-thread architecture (worker = parser + operator-list builder; main = canvas + text-layer + annotation-layer). Study how text-layer overlays canvas for selection, how the AcroForm/XFA layers are rendered, how the structure tree drives accessibility.
- https://github.com/embedpdf/embed-pdf-viewer (MIT) — modern framework-agnostic viewer with a clean plugin/headless architecture. Excellent reference for **plugin system, framework-agnostic core, ready-made vs headless dual-mode shipping**.
- https://github.com/anaralabs/lector (MIT) — composable headless PDF toolkit for React. Reference for **headless component composition** and how to expose PDF.js cleanly to app code.
- https://pdfjskit.com/ — UI layer on top of PDF.js with themes, RTL, search improvements. Reference for **what a polished PDF.js UI looks like**.
- https://github.com/wojtekmaj/react-pdf (MIT) — most popular React wrapper around pdfjs-dist. Reference for **component API ergonomics**.
- https://github.com/Hopding/pdf-lib (MIT) — pure-TS PDF writer. Reference for **incremental-save-correct PDF writing, AcroForm fill-and-flatten, cross-reference table maintenance**.
- https://github.com/foliojs/pdfkit (MIT) — Node.js PDF generation. Reference for **font embedding and content streams**.
- https://github.com/chromium/pdfium (BSD-3-Clause) — the engine in Chrome and many SDKs. Reference for **canonical rendering correctness** (what PDF.js gets approximately right; PDFium gets exactly right).
- https://github.com/ArtifexSoftware/mupdf.js (AGPL — study only) — fastest WASM PDF stack today. Reference for **what's possible** at the performance ceiling.
- https://github.com/pdfcpu/pdfcpu (Apache 2.0) — Go PDF library. Reference for **page-level operations API surface** (split, merge, optimize, extract, watermark, stamp, encrypt).

### Proprietary viewers to benchmark against (UX/feature inspiration only)

- **Adobe Acrobat Web** (https://acrobat.adobe.com) — the floor for "professional PDF". Replicate or beat: search precision, comment threading, form filling, fill-and-sign UX.
- **Nutrient (PSPDFKit) Web SDK** (https://nutrient.io) — the modern enterprise gold standard. Note: 15+ annotation types, real-time collaboration, AI chat & redaction, JSON-first API. Benchmark our annotation set, agent API ergonomics, and collaboration latency against it.
- **Apryse (PDFTron) WebViewer** — legacy enterprise champion. Note: in-browser DOCX/XLSX editing alongside PDF, fastest large-file handling. Benchmark our 1000-page handling against it.
- **PDF Expert (Readdle)** (https://pdfexpert.com) — best-in-class macOS/iOS reading UX. Benchmark our reading mode, scroll feel, and "feels like paper" polish against it.
- **ChatPDF** / **Humata** / **Paperpal** — AI-native PDF chat. Benchmark our chat-with-PDF, citation linking back to source pages, multi-document chat, and structured extraction.
- **Browser built-ins** — Chrome (PDFium), Firefox (PDF.js), Safari (Quartz). These are the baseline. Our viewer must be obviously and immediately better than any of them within the first 30 seconds of use.
- **Sumatra PDF**, **Okular**, **Skim** — open-source readers with cult followings. Study what their fans love (speed, keyboard shortcuts, vim-style navigation, distraction-free reading).
- **Shadow Reader** (https://shadowreader.io) — specialist dark-mode PDF viewer. Reference for **how to do dark mode that doesn't destroy figures, photos, and diagrams**.
- **Readwise Reader** — best-in-class web reading UX (not PDF-only, but the bar for "I want to read this carefully"). Reference for **highlighting, note-taking, "ghost reader" feel**.

### Format Standards (canonical truth, always prefer over implementations)

- **ISO 32000-2:2020** — PDF 2.0 specification (the source of truth). https://iso.org/standard/75839.html
- **ISO 32000-1:2008** — PDF 1.7 (still the most common in the wild)
- **PDF Association** — https://pdfa.org (community guidance, errata, accessibility patterns)
- **PDF/UA (ISO 14289-1)** — accessibility profile (tagged PDF, structure tree, alt text)
- **PDF/A (ISO 19005)** — archival profile (font embedding, no encryption, deterministic rendering)
- **WCAG 2.2** — the accessibility bar our viewer's UI must meet
- **Adobe technical notes** — XFA, JavaScript for Acrobat, Forms Data Format (FDF/XFDF)

---

## Project Structure

This sits inside the existing monorepo (DOCX/XLSX/PPTX already exist). Add:

```
/
  packages/
    core/              # already exists — reuse command bus, plugin system, presence
    pdf/               # NEW: PDF viewer/editor — parser-glue, model, renderer, agent API
    pdf-engine/        # NEW: thin abstraction over PDF.js (+ optional PDFium WASM fallback)
    pdf-annotations/   # NEW: annotation model + serializer (FDF/XFDF + native PDF annots)
    pdf-forms/         # NEW: AcroForm read/fill/flatten + signature field handling
    pdf-edit/          # NEW: page-level operations (rotate, reorder, split, merge, crop, redact)
    pdf-ai/            # NEW: text extraction, chunking, embeddings, chat-with-pdf, structured extraction
    agent/             # already exists — add PDF agent commands
    realtime/          # already exists — reuse for presence + comment sync
    comments/          # already exists — reuse comment threads
    ui/                # already exists — reuse design system primitives
  apps/
    web/
      app/
        pdf-viewer/    # NEW: the production-grade reading UI (the showcase)
  spec/
    pdf/               # NEW
  fixtures/
    pdf/               # NEW: real-world PDF corpus
  tests/
    roundtrip/pdf/     # NEW
    agent/pdf/         # NEW
  docs/
    build-log/pdf.md   # NEW
```

The packages above are deliberately fine-grained so that someone using only the viewer doesn't pull in OCR, AI, or editing code.

---

## Phase Structure

### Step A: Analyze

Before writing a single line of spec or code, deeply study the reference repos and products. Specifically answer:

1. **Engine choice.** PDF.js alone, PDF.js + PDFium WASM as fidelity fallback, or pure PDFium WASM? What are the tradeoffs (bundle size, license, fidelity, font support, color management, performance)? Recommend, justify, and write down the decision.
2. **Rendering pipeline.** Worker → operator list → main thread → canvas → text-layer + annotation-layer + struct-tree-layer. Where are the slow paths? Where do we use OffscreenCanvas? Where do we cache (rasterized pages? operator lists? text layers)?
3. **Memory model.** How do we keep memory bounded for a 1000-page document? What's our LRU strategy for rasterized pages? When do we evict text layers?
4. **Edit model.** How does pdf-lib do incremental updates? When do we incremental-save vs. fully re-serialize? How do we never break digital signatures?
5. **Annotation model.** How are PDF annotations represented natively (AP streams)? How do we round-trip our own annotations back into native PDF annotations so they're visible in Adobe?
6. **Form model.** AcroForm vs XFA. Field types. Calculation order. JavaScript actions (we will sandbox or no-op these).
7. **Accessibility model.** Tagged PDF, structure tree, marked content, role mapping. How do we expose this to screen readers in our UI?
8. **What competitors get wrong.** Slow first paint. Ugly default UI. Bad dark mode (inverts photos). Tiny click targets on annotations. No keyboard shortcuts. No agent API. Privacy-hostile (everything uploaded). Document each weakness — these become our wins.

Write analysis notes in `/spec/pdf/analysis.md`. These notes inform the spec but are not the spec.

---

### Step B: Spec

Produce the specification for PDF. The spec is the contract for the build. It must be complete enough that someone could implement it independently.

**Required spec documents:**

#### `/spec/pdf/`

- `feature-scope.md` — exactly what is IN the 80% and what is explicitly OUT; no ambiguity
- `engine-strategy.md` — engine choice (PDF.js vs PDFium-WASM vs hybrid), justification, fallback rules, bundle-size budgets
- `document-model.md` — our in-memory representation: `PdfDocument`, `PdfPage`, `PdfAnnotation`, `PdfFormField`, `PdfOutline`, `PdfStructTree`, `PdfRenditionState`. Include zod schemas.
- `rendering-pipeline.md` — main vs worker thread responsibilities; canvas + text-layer + annotation-layer + struct-layer composition; OffscreenCanvas usage; raster cache LRU policy; progressive rendering; print rendering pipeline
- `text-layer.md` — text extraction, glyph→character mapping, RTL & bidi handling, ligature handling, search hit-highlighting, copy-to-clipboard fidelity
- `annotation-model.md` — every supported annotation type, its native PDF representation, AP stream generation strategy, FDF/XFDF import/export
- `form-engine.md` — AcroForm field types, calculation order, format/validate scripts (sandboxed JS subset), flatten-on-save option, signature field handling
- `editing-pipeline.md` — page-level ops (insert/delete/reorder/rotate/crop/split/merge); incremental save strategy; how we never break signed regions; redaction model (visual + content stream + metadata scrub)
- `accessibility.md` — how we use the structure tree for screen reader output; reflow mode; high-contrast mode; keyboard navigation map (every action reachable from keyboard); WCAG 2.2 AA conformance plan
- `dark-mode.md` — color-scheme model: invert background only / smart invert (preserve images, code blocks, dark text on color highlights); per-document override; thumbnail consistency
- `search.md` — text search, regex search, semantic search (AI), fuzzy search, search across multiple PDFs; result ranking; navigation between hits; saved searches
- `ai-features.md` — chat-with-pdf, summarization, structured extraction, citation linking, embeddings store, multi-document chat, OCR-on-demand for scanned PDFs
- `collaboration.md` — presence, shared cursors over the page, comment threads anchored to PDF coordinates (stable across rotation/zoom), annotation sync via Y.js
- `agent-commands.md` — the complete list of agent-callable commands; for each: name, parameters, effect on model, PDF-bytes impact, example
- `cli.md` — the `office-agent pdf` subcommand surface
- `keyboard-shortcuts.md` — every action's keyboard shortcut, with modifier conventions matching macOS / Windows / Linux norms; vim-style mode for power users
- `performance.md` — measurable targets (first-page TTI, scroll FPS at 60, memory cap for 1000-page doc, search latency, annotation create latency)
- `edge-cases.md` — encrypted PDFs, password-protected PDFs, broken xref tables, linearized vs non-linearized, PDFs with embedded JavaScript, scanned PDFs needing OCR, very-large PDFs (>500MB), PDFs with thousands of annotations, PDFs with broken fonts, PDFs with custom CMaps
- `acceptance-criteria.md` — measurable done criteria: which fixture files must render-match the baseline, which annotation roundtrips must be lossless, which agent commands must work, which user flows must work, accessibility audit pass

**Spec quality bar:** Same as in `prompt.md`. A spec doc is done when it is self-contained, precise, honest about uncertainties, and actionable.

Do NOT begin building until the spec passes this bar.

---

### Step C: Build

Implement based on the spec. Follow this sub-order:

1. **Engine glue first** — wrap PDF.js cleanly; add the optional PDFium WASM fallback path; test against fixtures for visual fidelity (pixel diff against PDFium baseline)
2. **Document model** — pure TS, no DOM, no React; load → model; serialize → bytes; round-trip integrity
3. **Renderer (headless)** — render-to-canvas (Node.js with `node-canvas` or via PDF.js's serverless build); produces PNG/JPEG per page; this is the foundation for the agent API and CI tests
4. **Command bus** — every edit (annotate, fill, page-op) goes through commands; reuse `packages/core/command-bus`
5. **Agent API** — expose every operation programmatically; test headlessly before any UI exists
6. **Renderer (browser)** — connect model to browser UI; canvas + text-layer + annotation-layer + struct-tree-layer; virtual scrolling for pages
7. **Reading UI** — the showcase. Header (toolbar), sidebar (outline / thumbnails / annotations / search), main pane, status bar. Beautiful by default. Smooth.
8. **Annotation tools** — every annotation type, with hit testing, drag/resize, color/style picker, undo/redo
9. **Form filling** — render AcroForm widgets, validate, save filled
10. **Editing** — page ops: rotate / reorder / insert / delete / split / merge / extract / crop / redact
11. **AI features** — chat sidebar, semantic search, summarize, extract, OCR on demand
12. **Collaboration** — Y.js sync of annotations + comments + presence
13. **Polish** — dark mode, keyboard shortcuts, accessibility, performance pass, "feels like paper" details

**Build discipline:**

- Write rendering-fidelity tests before implementing each rendering feature (compare to PDFium baseline)
- Write roundtrip tests before implementing each editing feature
- Every command must be testable headlessly (no browser required)
- Every commit must not reduce roundtrip pass rate or fidelity score
- Every commit must not regress the perf budget (have a perf CI job)
- Keep `/docs/build-log/pdf.md` updated with non-trivial decisions

---

### Step D: Validate

Before declaring complete:

- [ ] **Rendering fidelity:** all `/fixtures/pdf/` render with ≤ X% pixel diff vs PDFium baseline (X to be set in spec, target: < 0.5% for non-CJK text fixtures; < 2% overall)
- [ ] **Roundtrip integrity:** every fixture survives load → no-op save → load with byte-equivalent unchanged objects
- [ ] **Edit roundtrip:** every supported edit operation applied to every fixture produces a file that opens cleanly in Adobe Acrobat Reader, Preview, and Chrome
- [ ] **Form roundtrip:** filled forms display correctly when reopened in Acrobat
- [ ] **Annotation roundtrip:** annotations created in our viewer are visible in Adobe; annotations created in Adobe survive a load + no-op save in ours
- [ ] **Signature integrity:** signed PDFs that we open and save (without touching signed regions) keep their signatures valid
- [ ] **Agent API:** all commands work headlessly via `office-agent pdf …`
- [ ] **Performance budget:** met (see `performance.md`)
- [ ] **Accessibility:** axe-core / Lighthouse audit passes WCAG 2.2 AA on the viewer UI; tagged PDFs are readable by VoiceOver / NVDA via our text/struct layers
- [ ] **License audit:** no AGPL or proprietary runtime dependency
- [ ] **Visual QA:** side-by-side comparison of our viewer vs Chrome built-in vs Adobe Acrobat Web on 10 hand-picked "hero" documents — ours must look obviously, immediately better

---

## The 80% Scope

### PDF Viewer — In Scope (the table-stakes)

**Reading**
- Open any real-world PDF (PDF 1.4 → 2.0); local file, URL, drag-drop, paste, share-target
- Continuous, single-page, two-page (book/cover), and presentation view modes
- Smooth zoom: pinch, ctrl+wheel, fit-width, fit-page, fit-actual, custom %, double-tap-to-zoom
- Pan: click-drag (hand tool), arrow keys, Page Up/Down, Home/End
- Rotation per-page and per-document
- Sidebar with: **outline** (bookmarks), **thumbnails** (with reorder/select/extract), **annotations list**, **search**, **AI chat**, **attachments**
- Outline navigation (deep links into document)
- Page number jump (Cmd/Ctrl+G)
- Goto-link annotations and external URI annotations
- Named destinations
- Document properties dialog (metadata, security, fonts used)
- Print (with proper page selection, scaling, even/odd, booklet)
- Download original / download modified
- Share (link with permissions, when collab enabled)

**Reading polish (the differentiators)**
- **Dark mode** that smart-inverts: page background dark, text light, but **photos and color-rich figures preserved as-is**; toggle per-document; system-aware; remembers preference
- **Reading ruler** / line focus mode (dim everything except current line)
- **Distraction-free** mode (chrome hides; tap to reveal)
- **Read aloud** (TTS, with sentence highlight following narration)
- **Continuous scroll inertia** that feels like Apple Preview / PDF Expert
- **Gesture support**: pinch-zoom, two-finger swipe, three-finger forward/back through history
- **Vim-style keymap** (j/k scroll, gg/G top/bottom, /search, ?reverse-search, n/N next/prev hit) as opt-in
- **Mini-map** (the entire document as a slim ribbon down the side, click to jump)
- **"Where am I" indicator** that surfaces the current section heading from the outline as you scroll
- **Persistent reading position** per document (remember scroll, zoom, mode across sessions)

**Search**
- Full-text search with hit highlighting, hit count, prev/next, jump to result
- Match case, match whole word, regex, **fuzzy** (typo-tolerant)
- **Semantic search** ("show me the part where they discuss return policy") — AI-powered
- **Search across multiple open documents**
- Search panel shows snippets with surrounding context

**Selection & extraction**
- Text selection (per-glyph, per-word double-click, per-paragraph triple-click)
- Rectangular (lasso) selection for image extraction
- Copy text with formatting preserved (or as plain)
- Copy as Markdown
- Right-click → "Translate selection", "Define", "Search the web", "Ask AI about this"
- Extract image at original resolution

**Annotations** (all roundtrip into native PDF annotation objects with valid AP streams)
- Highlight, underline, strikethrough, squiggly (text markup annotations)
- Sticky note / popup comment (with thread, replies, resolve)
- Free text annotation
- Free-hand drawing (ink) with smoothing and pressure (when supported)
- Shapes: line, arrow, rectangle, ellipse, polygon, polyline
- Stamp (preset + custom image stamps; "Approved" / "Draft" / signature stamp)
- Caret / insert text mark
- File attachment annotation (attach a file at a point)
- Link annotation (URI / goto-page / goto-named-destination)
- **Redaction**: visual blackout + content stream removal + metadata scrub (cannot be reversed by selecting text under the box; not Photoshop-over-text)
- Color, opacity, line width, font, line dash for every applicable annotation
- Rich annotation list panel: filter by author, type, page, date; jump to; resolve; reply

**Forms (AcroForm)**
- Render every field type: text, textarea (multiline), checkbox, radio, combobox, listbox, button, signature, push-button, file-attachment
- Fill, validate (regex / format constraints from `/V`, `/MaxLen`, `/AA` actions where safe)
- Calculation order respected
- Save filled (preserve as fillable) **or** flatten (bake values into content stream, render-only)
- Reset form
- Import/export FDF/XFDF
- For XFA: render the AcroForm fallback if present; otherwise show a clear "this form requires Adobe Acrobat" message with a "convert to AcroForm" option (best-effort)

**Page-level editing**
- Rotate single page or selection (90° / 180° / 270°)
- Reorder pages (drag in thumbnail panel)
- Insert blank page / insert pages from another PDF
- Delete pages
- Extract pages (save as new PDF)
- Split (by range, by size, by bookmark)
- Merge multiple PDFs
- Crop pages (single, range, all; with margin presets)
- Add watermark / stamp / page numbers / header / footer
- Resize / change page size

**Signatures**
- Display existing signatures with validity badge (signature panel)
- Show signing certificate chain
- Detect changes after signing (LTV / DSS dictionary)
- Optional: place a visible signature image + sign with PKCS#12 (deferred; spec must say)

**Accessibility**
- Tagged PDF reading order respected for screen readers
- Reflow view (single-column, font-size adjustable, ignores layout)
- High-contrast mode
- Full keyboard navigation: every action reachable, every annotation focusable
- ARIA labels on all UI controls
- WCAG 2.2 AA conformance for the viewer chrome
- Skip-links to: toolbar, sidebar, page content
- Captions / TTS for read-aloud
- Respects `prefers-reduced-motion`

**Performance**
- First page rendered in < 600 ms on a mid-range laptop (cold load, 50 MB PDF over local file)
- Scroll at 60 fps with virtual page rendering (only ±2 pages around viewport are in canvas)
- Memory bounded: 1000-page document under 600 MB RSS in browser
- Search returns first hit in < 200 ms for 500-page doc
- Annotation create-and-render < 50 ms
- Worker-based parsing (never blocks the UI thread)
- Lazy load: thumbnails, text-layer, AI features all on-demand

**Privacy & security**
- 100% client-side rendering by default; nothing leaves the browser unless the user opts in (e.g. to use cloud AI)
- Local AI model option (WebLLM / wllama) for "ask the PDF" without ever uploading the file
- Sandbox embedded JavaScript in PDFs (AcroForm `/AA` actions, document-level scripts) — by default, do not execute
- Encrypted PDF support: prompt for password, decrypt locally
- Clear visual indication when a PDF contains: JS, external links, attachments, forms, encryption, signatures

**Collaboration** (reuse `packages/realtime` + `packages/comments`)
- Presence: see other readers' cursors and current page
- Annotation sync via Y.js (CRDT, no server lock)
- Comment threads anchored to PDF coordinates (page + rect), stable across rotate/zoom
- @-mentions in comments, resolve / reopen
- Activity feed

### PDF — Explicitly Out of Scope

- **Editing the existing text inside a page** (i.e. re-flowing paragraphs already rasterized into the content stream). This is a tar pit. We support text *overlay* and redaction; not text *editing*.
- **XFA dynamic forms** beyond rendering the AcroForm fallback (Adobe themselves are deprecating this)
- **3D / PRC content** (preserve, do not render)
- **Multimedia annotations** (movie / sound / 3D / RichMedia — preserve as-is, do not play)
- **PDF/A creation/conformance certification** (we may *open* PDF/A files; we do not certify outputs as PDF/A)
- **Embedded JavaScript execution** (security; we sandbox / no-op)
- **Native code signing** of digital signatures (deferred — placing visible signature images is in scope; cryptographic signing is opt-in deferred)
- **Heavy-duty OCR** beyond on-demand `tesseract.js` for selectable-text overlay (we are not a Document AI / IDP product)
- **Bates numbering, redaction logs, legal hold** workflows (defer)
- **Server-side conversion** to other formats beyond text/markdown extraction (defer)

---

## The AI-Native Design (Most Important Section)

This is the differentiator. The viewer is built so that AI agents are first-class users — not an afterthought bolted on.

### Core Principle: Everything Is a Command

Same invariant as DOCX/XLSX/PPTX. No direct model mutation. Every change — human click, agent call — flows through the **command bus** in `packages/core`.

```typescript
type PdfCommand =
  | { type: 'pdf:add-annotation'; payload: AddAnnotationPayload }
  | { type: 'pdf:update-annotation'; payload: UpdateAnnotationPayload }
  | { type: 'pdf:delete-annotation'; payload: { id: string } }
  | { type: 'pdf:fill-form-field'; payload: { fieldName: string; value: FormFieldValue } }
  | { type: 'pdf:rotate-pages'; payload: { pages: number[]; angle: 90 | 180 | 270 } }
  | { type: 'pdf:reorder-pages'; payload: { from: number[]; to: number[] } }
  | { type: 'pdf:insert-pages'; payload: { at: number; sourceBuffer: ArrayBuffer; sourcePages?: number[] } }
  | { type: 'pdf:delete-pages'; payload: { pages: number[] } }
  | { type: 'pdf:redact'; payload: { rects: PdfRect[]; replacement?: RgbColor } }
  | { type: 'pdf:add-text-overlay'; payload: { page: number; rect: PdfRect; text: string; style: TextStyle } }
  | { type: 'pdf:add-image-overlay'; payload: { page: number; rect: PdfRect; data: ArrayBuffer } }
  | { type: 'pdf:add-watermark'; payload: { text?: string; image?: ArrayBuffer; opacity: number; pages: number[] | 'all' } }
  | { type: 'pdf:add-bookmark'; payload: { title: string; destination: PdfDestination; parent?: string } }
  | { type: 'pdf:flatten-form' }
  | { type: 'pdf:flatten-annotations' }
  | { type: 'pdf:set-metadata'; payload: PdfMetadataPatch }
```

### The Agent API (Headless-First)

Same shape as the sibling format agents:

```typescript
interface PdfAgent {
  // Read
  getSnapshot(): PdfDocumentSnapshot;
  getPage(pageIndex: number): PdfPageSnapshot;
  getText(range: PdfRangeSpec, opts?: { format?: 'plain' | 'markdown' | 'html' }): string;
  getOutline(): PdfOutlineNode[];
  getAnnotations(filter?: PdfAnnotationFilter): PdfAnnotation[];
  getFormFields(): PdfFormField[];
  search(query: PdfSearchSpec): PdfSearchResult[];

  // AI
  ask(question: string, opts?: AskOpts): Promise<AskResult>; // returns answer + citations
  summarize(opts?: SummarizeOpts): Promise<string>;
  extract<T>(schema: ZodSchema<T>): Promise<T>; // structured extraction
  ocr(pages?: number[]): Promise<void>; // scanned → text-layer

  // Render (headless)
  renderPage(pageIndex: number, opts: RenderOpts): Promise<Uint8Array>; // PNG/JPEG/WebP
  renderThumbnail(pageIndex: number, opts?: ThumbOpts): Promise<Uint8Array>;

  // Write (everything goes through command bus)
  applyCommand(command: PdfCommand): Promise<Mutation>;
  applyCommands(commands: PdfCommand[]): Promise<Mutation[]>;

  // Diff & Review (mirrors sibling formats)
  getDiff(from: PdfDocumentSnapshot, to: PdfDocumentSnapshot): PdfDocumentDiff;
  getPendingMutations(): Mutation[];
  approveMutation(id: string): void;
  rejectMutation(id: string): void;
  rollback(id: string): void;

  // I/O
  importFile(buffer: ArrayBuffer): Promise<void>;
  exportFile(opts?: { incremental?: boolean }): Promise<ArrayBuffer>;
}
```

This must work **headlessly** — zero DOM, zero React, zero browser. An agent in Node.js on a server must load a PDF buffer, query it, edit it, render pages to PNG, and export bytes.

### AI Features (the showcase)

These are the things that turn a PDF viewer into a tool people *prefer*.

- **Chat with PDF** — sidebar chat. Every answer cites source pages with hover-preview and click-to-jump. Multi-turn, with conversation memory scoped to the document.
- **Summarize** — full doc, current section, current page, current selection. Choice of length and style (executive / detailed / bullet / outline).
- **Translate** — selection or full doc, with side-by-side display.
- **Structured extraction** — "extract all dates and amounts from this contract as JSON matching this schema". Powered by zod schemas via `ai`-SDK / function calling.
- **Compare two PDFs** — visual diff (per-page rendering side-by-side with highlighted differences) + text diff (with semantic similarity for re-flowed paragraphs) + structural diff (annotations, form fields, metadata).
- **Smart redaction** — "redact all PII", "redact all phone numbers", "redact every mention of $name". Preview before commit. Output a redaction log.
- **OCR on demand** — for scanned PDFs, run `tesseract.js` per-page to add an invisible text layer (selectable, searchable, AI-accessible).
- **Auto-tagging** — propose a tag tree (PDF/UA structure) for an untagged PDF; human reviews + applies.
- **Q&A pre-warming** — embed the document in the background once opened; chat is instant from then on.
- **Multi-document chat** — open several PDFs, ask "compare the cancellation clauses across these contracts".

### The Human Review Flow

Same model as DOCX/XLSX/PPTX. Agent-source mutations are staged into a pending queue (annotations show as "proposed", page operations as "preview"). Human approves / rejects / approves-all / rolls back. UI marks proposed annotations distinctly from approved ones (e.g. dashed border).

### The CLI / Programmatic Interface

Extend the existing `office-agent` CLI:

```bash
# inspect
office-agent pdf inspect --file report.pdf
# → page count, page sizes, fonts, annotations summary, forms summary, security, signatures

# read
office-agent pdf read --file report.pdf --pages 1-5 --format markdown
# → markdown of pages 1–5

office-agent pdf read --file report.pdf --bbox "page=3,x=100,y=200,w=400,h=300" --format text

office-agent pdf outline --file report.pdf
# → JSON outline tree

office-agent pdf metadata --file report.pdf
# → JSON metadata

# render
office-agent pdf render --file report.pdf --pages 1-3 --out ./pages/ --format png --dpi 150

office-agent pdf thumbnail --file report.pdf --page 1 --out cover.webp --width 400

# search
office-agent pdf search --file report.pdf --query "cancellation" --format json
# → array of {page, bbox, snippet}

office-agent pdf ask --file report.pdf --question "What is the notice period?"
# → answer + citations (page, bbox)

office-agent pdf summarize --file report.pdf --style bullet
# → summary

office-agent pdf extract --file invoice.pdf --schema ./invoice.schema.json
# → JSON matching the schema

# edit
office-agent pdf rotate --file report.pdf --pages 2,4,6 --angle 90 --out report-rotated.pdf

office-agent pdf reorder --file report.pdf --order "1,3,2,4-end" --out reordered.pdf

office-agent pdf split --file report.pdf --by bookmark --out ./chapters/

office-agent pdf merge --files a.pdf b.pdf c.pdf --out merged.pdf

office-agent pdf extract-pages --file report.pdf --pages 1-10 --out chapter1.pdf

office-agent pdf redact --file report.pdf --pattern "phone" --pattern "email" --out redacted.pdf --log redaction.json

office-agent pdf fill-form --file form.pdf --data ./values.json --flatten --out filled.pdf

office-agent pdf annotate --file report.pdf --annotations ./annotations.xfdf --out annotated.pdf

office-agent pdf watermark --file report.pdf --text "DRAFT" --opacity 0.2 --out wm.pdf

office-agent pdf ocr --file scan.pdf --lang deu+eng --out searchable.pdf

# diff
office-agent pdf diff --before v1.pdf --after v2.pdf --format html --out diff.html
```

The CLI is the primary interface for agents in server-side pipelines. Pipeable, scriptable, composable.

---

## Import / Export Requirements

### Import

- Accept any valid PDF 1.4 → 2.0
- Accept linearized and non-linearized
- Accept encrypted PDFs (prompt for password; decrypt in-memory; never write the password)
- Accept files produced by Acrobat, Word, Chrome, macOS Quartz, LaTeX, LibreOffice, Google Docs, InDesign, generic scanners, mobile apps
- Surface a clear error on broken xref / corrupted streams; never crash silently
- Recover gracefully from minor corruption when possible (rebuild xref, salvage pages)

### Export

- Default: **incremental update** (preserve original bytes; append updates) — minimizes diff and keeps signatures intact
- Optional: **full re-serialize** (compact, optimized) — for "save as cleaned copy"
- Output validated against `qpdf --check` in CI
- No false-warning dialogs in Acrobat / Preview / Chrome
- Preserve all PDF objects we did not touch byte-for-byte
- Embedded fonts: re-embed; subset; never strip
- Color spaces: preserve original (RGB / CMYK / Lab / ICC); never silently convert

### Import/Export API

```typescript
// Browser
const viewer = await PdfViewer.fromBuffer(arrayBuffer)
const outputBuffer = await viewer.export({ incremental: true })
const blob = new Blob([outputBuffer], { type: 'application/pdf' })

// Node.js (headless)
const agent = await PdfAgent.fromBuffer(fs.readFileSync('document.pdf'))
await agent.applyCommand({ type: 'pdf:add-annotation', payload: { ... } })
const output = await agent.exportFile({ incremental: true })
fs.writeFileSync('document-annotated.pdf', output)
```

---

## Fixture Corpus

Before building, collect real-world PDFs. This is not optional. Aim for breadth (lots of producers / styles) and depth (the hard cases).

### Producers (collect ≥ 2 fixtures from each)

- Adobe Acrobat (modern + legacy)
- Microsoft Word "Save as PDF"
- Apple Pages / Preview
- Chrome "Print to PDF"
- LibreOffice
- Google Docs / Sheets / Slides export
- LaTeX (pdfTeX, XeLaTeX, LuaTeX) — academic papers
- InDesign — magazine layouts, brochures
- Generic scanners (Canon, HP, Fujitsu) — scanned + OCR'd
- Mobile (CamScanner, Adobe Scan) — auto-cropped scans

### Categories

- **Business documents**: 5 invoices (Rechnungen), 5 contracts (Verträge), 3 letters, 3 reports
- **Academic**: 3 papers with figures, references, tables; 1 LaTeX-typeset thesis (200+ pages)
- **Forms**: 3 fillable AcroForm tax/admin forms (e.g. German Steuerformular), 1 XFA form (to test fallback)
- **Marketing**: 2 magazine-style PDFs from InDesign; 2 product brochures with rich graphics
- **Books & long docs**: 1 book (300+ pages); 1 user manual (100+ pages with TOC)
- **Scans & OCR**: 3 pure scans (no text layer); 2 scans with OCR text layer
- **Signed & encrypted**: 2 digitally signed PDFs; 2 password-protected PDFs; 1 with permissions restrictions
- **Annotated**: 2 with comments and highlights from Adobe Acrobat
- **Edge cases**: 1 huge file (>200 MB), 1 with 1000+ pages, 1 with thousands of annotations, 1 with non-Latin scripts (Arabic / CJK / Cyrillic), 1 with broken xref (recovery test), 1 with custom CMaps, 1 PDF/A
- **Multilingual**: at least one fixture per major script (Latin, German with umlauts, French with accents, Arabic RTL, CJK, Cyrillic, Thai, Hebrew)

If you cannot access real files during the build, generate realistic synthetic fixtures using Python (`reportlab`, `pypdf`, `pdfplumber`) — but flag these as synthetic and plan to replace with real-world files before production.

---

## Architecture Principles (Non-Negotiable)

1. **Headless-first.** Parser, model, command bus, serializer, AND a node-based renderer all run in Node.js with zero DOM. The browser is a rendering surface. This is what makes the agent API real.

2. **Commands are the only mutation path.** Direct model mutation is forbidden outside the parser. Everything else flows through the command bus. This invariant enables diffs, review, rollback, multi-agent coordination, undo.

3. **PDF bytes are the source of truth, not the model.** The in-memory model is a working surface. The PDF file is what's saved. When in doubt, keep the bytes closer to the original (incremental update by default).

4. **Opaque object preservation.** Any PDF object the editor doesn't understand is round-tripped verbatim. We never silently drop unknown content. We never reformat objects we didn't change.

5. **Signature-aware editing.** The editor knows which byte ranges are covered by signatures. It refuses (or loudly warns + invalidates) any edit that would break a signature.

6. **Engine isolation.** PDF.js is wrapped behind `packages/pdf-engine`. Switching to PDFium WASM (or adding it as a fidelity fallback) is an internal change, not a viewer rewrite.

7. **Two-thread discipline.** Parsing, content-stream interpretation, search, OCR, embeddings — all in workers. The main thread does only canvas painting, hit testing, and DOM updates. The UI thread budget is 16 ms per frame; we measure it.

8. **Progressive everything.** First page paints fast (< 600 ms). Other pages render lazily as the user scrolls. Text layer renders after canvas. Thumbnails render at low DPI on demand. AI embeddings build in the background after first paint.

9. **Privacy by default.** Nothing leaves the browser unless the user opts in. The default AI provider is local (WebLLM); cloud AI is opt-in and clearly labeled.

10. **Accessibility is a feature, not a checkbox.** The viewer's chrome is WCAG 2.2 AA from day one. Tagged PDFs are fully accessible to screen readers via our text + struct layers. Reflow mode is first-class.

11. **Fail loudly.** Import failures, render failures, signature breakage, OCR errors — all surface as structured errors with useful messages. Never silent corruption.

12. **No fake fidelity.** When a font can't be embedded, when a color profile can't be matched, when an unusual feature (e.g. transparency group with non-standard blending) renders approximately — say so. Don't lie to the user.

---

## Output at the End

When complete, produce:

1. **`/spec/pdf/`** — all spec documents, complete and up-to-date
2. **`/packages/pdf*/`** — the implementation, split by concern as listed above
3. **`/apps/web/app/pdf-viewer/`** — the showcase reading UI
4. **`/tests/roundtrip/pdf/` and `/tests/agent/pdf/`** — passing test suites
5. **`/fixtures/pdf/`** — real-world test corpus (or synthetic with a TODO to replace)
6. **`/docs/build-log/pdf.md`** — decisions, deviations from spec, known issues, performance numbers
7. **A "10/10 demo script"** in the build log: a 60-second guided tour through the showcase features that proves we beat Chrome and Acrobat Web side-by-side
8. **Lighthouse + axe + perf trace artifacts** in `/docs/build-log/pdf-audits/`

---

## The 10/10 Bar — What "Better Than Acrobat" Looks Like

These are the moments a user notices. The build is not done until each is true.

1. **First paint is faster than Chrome's built-in viewer.** Open a 50 MB PDF — first page is visible before the user can blink.
2. **Scrolling through a 1000-page PDF is butter-smooth.** Never a stutter. Never a blank page. Never a memory leak warning.
3. **Search-as-you-type returns the first hit before you finish typing.**
4. **Dark mode looks beautiful** — text inverts crisply, photos and figures are untouched, code blocks stay legible.
5. **Selecting text feels like a native app** — no jagged misalignment with the canvas, no missing characters, copy preserves layout when it should.
6. **Annotations feel native.** Tap an annotation — instantly editable. Drag — never lags. Color picker — beautiful and fast.
7. **Form filling is delightful.** Tab between fields, validation is instant and helpful, auto-fill from a saved profile.
8. **"Ask the document" actually works.** The answer cites the exact paragraph. Click the citation — page jumps and highlights it. The latency is short enough that the user keeps using it.
9. **Compare two PDFs** is something a user will pay for. It Just Works.
10. **Reading position survives** across reloads, devices, and weeks. Pick up exactly where you left off.
11. **It's accessible.** A blind reviewer using VoiceOver / NVDA can do everything a sighted reviewer can: navigate the outline, jump pages, read structured content, fill forms, hear annotations.
12. **It's private.** A user opens a confidential contract and is shown — clearly — that nothing has left their browser. The cloud-AI option is opt-in, per-document, with a visible indicator.

---

## Start Instructions

1. Read this entire prompt twice.
2. Confirm:
   - You understand the clean-room constraint and will not copy code from AGPL projects (MuPDF, etc.).
   - You understand the visual fidelity bar and will not move forward without a pixel-diff CI job against PDFium baselines.
   - You understand the headless-first / agent-first design requirement.
   - You understand that "the best PDF viewer on the public internet" is a real bar, not a slogan — every section of this spec exists to drive toward it.
3. Set up `/packages/pdf*`, `/spec/pdf/`, `/fixtures/pdf/`, `/apps/web/app/pdf-viewer/` skeletons.
4. Collect or generate the PDF fixture corpus.
5. Begin Step A: Analyze. Study the reference repos and products. Write `/spec/pdf/analysis.md`.
6. Step B: Spec. Produce every spec doc above.
7. Step C: Build. Engine glue → model → headless renderer → command bus → agent API → browser renderer → reading UI → annotations → forms → editing → AI → collaboration → polish.
8. Step D: Validate. Pass every checklist item. Capture the 10/10 demo script and the audit artifacts.

Ask no clarifying questions. Begin.
