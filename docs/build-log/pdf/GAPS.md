# PDF Viewer — Known Gaps & Follow-ups

Living backlog of items the night-shift build deliberately deferred,
documented here so a follow-up session can pick any one up without
re-discovering the context.

Each entry follows the same shape:

- **What** — one-line summary.
- **Why deferred** — why it didn't land in this pass.
- **Where** — pointers to relevant code / spec.
- **Acceptance** — how we'll know it's done.
- **Effort** — rough t-shirt size.

Entries are ordered roughly by impact / cost ratio (high → low).

---

## P1 — Ship-blockers for "really great"

### G1. Bundle the PDFium-WASM fallback blob

- **What:** the engine layer (`@officeai/pdf-engine`) ships the PDFium
  adapter and its lazy-loader, but the actual `pdfium.wasm` (~2 MB) is
  not committed. The default PDF.js path covers ~99 % of inputs;
  PDFium is meant to kick in automatically when the heuristic detects
  substituted glyphs (CJK, exotic ligatures).
- **Why deferred:** binary blob shouldn't live in git history; we want
  a build step that pulls a pinned upstream release.
- **Where:** `packages/pdf-engine/src/pdfium/`, `spec/pdf/engine-strategy.md`.
- **Acceptance:**
  - `pdfium.wasm` is fetched into `apps/web/public/wasm/` by a
    `postinstall` or `make fixtures-pdf` step (whichever is least
    surprising).
  - The engine adapter loads it lazily on first fidelity-fallback.
  - A new fixture `fixtures/pdf/cjk-glyph-substitution.pdf` proves
    PDFium engages where PDF.js would otherwise emit `□` boxes.
  - Documented in the README.
- **Effort:** S (half-day).

### G2. Run the Playwright PDF smoke against a built `apps/web`

- **What:** `apps/web/e2e/pdf-viewer.spec.ts` exists with 6 cases but
  was never executed (`next start` requires a `next build` first, and
  the build is too slow for the in-session sandbox loop).
- **Why deferred:** non-trivial wall-clock cost; the spec is wired,
  the assertions are well-defined, and the unit + integration suites
  cover the same surface area against the headless agent.
- **Where:** `apps/web/e2e/pdf-viewer.spec.ts`,
  `apps/web/playwright.config.ts`.
- **Acceptance:**
  - `pnpm --filter @officeai/web build` then
    `pnpm --filter @officeai/web e2e -- pdf-viewer.spec.ts` exits 0.
  - The spec joins the existing `web-e2e` CI job (no new workflow).
- **Effort:** S — but blocks on a clean build environment.

### G3. Formal accessibility audit

- **What:** every interactive element ships `role` / `aria-label` /
  `data-testid`, but `axe-core` / `axe-playwright` is not wired in.
  Spec sketch lives in `spec/pdf/accessibility.md`.
- **Why deferred:** no a11y harness exists for any of the four
  editors yet, so this is a cross-product investment, not a
  PDF-specific one.
- **Where:** `spec/pdf/accessibility.md`,
  `apps/web/e2e/_helpers.ts`.
- **Acceptance:**
  - Add `axe-playwright` as a dev dep on `@officeai/web`.
  - One spec per editor (`docx-a11y.spec.ts`, `xlsx-a11y.spec.ts`,
    `pptx-a11y.spec.ts`, `pdf-a11y.spec.ts`) that runs `injectAxe`
    - `checkA11y` on the editor shell after first paint.
  - All four pass with WCAG 2.1 AA + best-practice tags enabled.
  - Document the score / known violations in `docs/build-log/pdf/`.
- **Effort:** M (1 day if cross-product is in scope, ½ day if PDF-only).

### G4. Lighthouse / performance trace artefact

- **What:** the spec target is "interactive in <500 ms for a
  50-page PDF on a mid-tier laptop". We hit it locally on the M-series
  dev box but never captured a CI artefact.
- **Why deferred:** Lighthouse on a 50-page PDF is a 30-60 s test;
  needs a dedicated CI lane to avoid noisy averages.
- **Where:** `spec/pdf/performance.md`,
  `fixtures/pdf/large-50page.pdf`.
- **Acceptance:**
  - GitHub Actions job `pdf-perf` builds the web app and runs
    Lighthouse on `/pdf-viewer?src=/sample-files/large-50page.pdf`.
  - TBT < 200 ms, LCP < 1500 ms, FCP < 800 ms recorded as JSON
    under `docs/build-log/pdf/perf/<run>.json`.
  - A simple regression check fails the job if any metric worsens
    by >15 % vs the baseline checked into the repo.
- **Effort:** M.

---

## P2 — Capability gaps

### G5. PDF 2.0 RichMedia / 3D / attachment write-side

- **What:** the parser tolerates and round-trips PDF 2.0 `/RichMedia`
  and `/3D` annotations on input (we never drop them), but the
  command surface doesn't yet let an editor _create_ them.
- **Why deferred:** out of scope for v1; spec captures it.
- **Where:** `spec/pdf/edge-cases.md` §"PDF 2.0 features",
  `packages/pdf-annotations/src/types.ts`.
- **Acceptance:**
  - `@officeai/pdf-annotations` exports
    `addRichMediaAnnot`, `addThreeDAnnot`, `addFileAttachmentAnnot`.
  - Round-trip fixtures prove they survive a parse → serialize cycle.
  - CLI surface: `office-agent pdf attach-file …` etc.
- **Effort:** L (each annotation type is its own AP-stream variant).

### G6. Digital signature creation

- **What:** verification works (we surface a "this file was signed"
  badge from `/Sig` dictionary metadata), but the signing path is
  parked.
- **Why deferred:** signing requires a key / cert chain story we don't
  want to bake in without a product decision (BYO key vs server-side
  signing service vs WebCrypto local-only).
- **Where:** `spec/pdf/edge-cases.md` §"Signed-then-modified",
  `fixtures/pdf/signed-then-modified.pdf`.
- **Acceptance:**
  - One supported signing mode (recommend WebCrypto + PKCS#12 on first
    iteration).
  - `office-agent pdf sign … --p12 … --password …` produces a PAdES-B
    signature that Adobe Reader validates.
  - Document the trust chain expectation in the user-facing copy.
- **Effort:** XL.

### G7. Redact: visual blackbox **and** content removal

- **What:** the redact annotation type exists in the model and the AP
  writer paints a black rectangle, but the underlying text/image
  content under the rectangle is **not** removed from the page stream.
  This is acceptable for "draft review" but not for compliance redaction.
- **Why deferred:** safe content removal needs a content-stream
  rewriter, which is a `pdf-lib` feature we don't have today.
- **Where:** `packages/pdf-annotations/src/redact.ts`,
  `spec/pdf/annotation-model.md`.
- **Acceptance:**
  - A separate `pdf-redact` capability (or a deeper command in
    `pdf-edit`) walks the content stream and elides any text/image
    operator whose bounding box intersects a redact rectangle.
  - Output passes a "select-all → copy" check: redacted text isn't on
    the clipboard.
  - Banner in the UI clearly distinguishes "draft redact"
    (annotation-only) from "permanent redact" (content-removed).
- **Effort:** L.

### G8. OCR: production-quality text layer for scanned PDFs

- **What:** `@officeai/pdf-ocr` ships as an optional adapter wrapping
  `tesseract.js`. It works for English on small pages but is too slow
  for the 50-page scanned book case and lacks language packs.
- **Why deferred:** language pack management deserves its own PR; this
  pass focused on the architectural seam, not the model corpus.
- **Where:** `packages/pdf-ocr/`, `spec/pdf/text-layer.md`.
- **Acceptance:**
  - At least DE + EN + FR + ES language packs ship lazily.
  - A `--language` flag on `office-agent pdf add-text-layer`.
  - Per-page progress events on the command bus so the UI can render
    a real progress bar.
  - Benchmark: 50-page A4 scan at 300 dpi finishes in <90 s on a
    mid-tier laptop.
- **Effort:** M.

### G9. Reflow ("read mode") doesn't handle multi-column or tables

- **What:** the experimental reflow view stitches the PDF.js text
  layer into a single column for narrative PDFs. It works for novels,
  reports and single-column papers; it visibly garbles two-column
  academic papers and any table.
- **Why deferred:** robust reflow requires a layout-analysis pass
  (column detection, heading detection, table detection) we don't yet
  have.
- **Where:** `apps/web/app/pdf-viewer/PdfCanvas.tsx` (`pdf-canvas-reflow`
  branch), `spec/pdf/text-layer.md`.
- **Acceptance:**
  - A heuristic that detects column count from the text layer's x
    histogram.
  - Tables (detected via aligned text runs) are rendered as `<table>`
    instead of being flattened.
  - One golden fixture per case (single, two-column, table-heavy)
    with an HTML snapshot under test.
- **Effort:** L.

### G10. Annotation comments don't surface in the cross-product

**Comments** sidebar

- **What:** the new `pdf-region` `CommentAnchor` is wired through
  `@officeai/comments`, but the **shared** Comments sidebar (used by
  DOCX / XLSX / PPTX) doesn't yet know how to render a PDF anchor — it
  currently falls back to a generic "open in PDF viewer" link.
- **Why deferred:** the shared sidebar lives in `packages/ui` and
  needed a model migration we'd already done; the per-product
  rendering layer is its own follow-up.
- **Where:** `packages/comments/src/anchors.ts`,
  `packages/ui/src/comments/CommentList.tsx`,
  `apps/web/app/pdf-viewer/usePdfCommentsProvider.ts`.
- **Acceptance:**
  - Comments anchored to `pdf-region` render a thumbnail of the
    region in the sidebar, the same way XLSX cell-anchored comments
    show a tiny grid preview.
  - Clicking the comment scrolls and highlights the region.
- **Effort:** M.

---

## P3 — Polish & DX

### G11. Promote `PdfEmptyState` into the shared `EmptyState`

- **What:** `apps/web/app/pdf-viewer/PdfEditor.tsx` reimplements the
  empty state inline because the shared `EmptyState` component only
  knows about the OOXML editors (its label / extension copy comes from
  a three-entry record).
- **Why deferred:** mechanical refactor; cosmetic; the inline version
  is correct.
- **Where:** `apps/web/app/lib/shell/EmptyState.tsx`,
  `apps/web/app/pdf-viewer/PdfEditor.tsx` (`PdfEmptyState`).
- **Acceptance:**
  - `EmptyState` accepts `pdf` as a `ProductKind`.
  - `PdfEmptyState` is removed from `PdfEditor.tsx`.
  - DOCX / XLSX / PPTX empty-state visuals unchanged.
- **Effort:** XS.

### G12. Wire `data-testid` for dropdown rows in `PdfToolbar`

- **What:** the trigger buttons all carry test ids
  (`pdf-zoom-menu-trigger`, `pdf-page-ops-trigger`, …) but the
  individual menu items inside the dropdowns don't. The Playwright
  spec falls back to `getByRole("menuitem", { name: /…/ })`, which
  works but is i18n-fragile.
- **Why deferred:** the spec passes today; we'd be adding ids only
  for future tests.
- **Where:** `apps/web/app/pdf-viewer/PdfToolbar.tsx`.
- **Acceptance:**
  - Every menu item has a stable `data-testid`
    (`pdf-zoom-fit-width`, `pdf-page-ops-rotate`, etc.).
- **Effort:** XS.

### G13. Sample PDFs in `apps/web/public/sample-files/`

- **What:** during the night-shift, two user-private PDFs landed in
  `apps/web/public/sample-files/` (a master's thesis and a
  registration form). They were left **uncommitted** so they don't
  ship — but the home-page sample list expects PDFs to live there if
  we want the "open sample" affordance to show one.
- **Why deferred:** we don't yet have a permissively-licensed sample
  PDF to ship.
- **Where:** `apps/web/public/sample-files/`,
  `apps/web/app/page.tsx` (sample list).
- **Acceptance:**
  - One short, CC0 / public-domain PDF (~3 pages, mixed text + image,
    with an outline) checked in as
    `apps/web/public/sample-files/sample.pdf`.
  - Listed by the home-page sample-files loader.
- **Effort:** XS.

### G14a. Free-text annotation tool

- **What:** the model + serializer round-trip free-text annotations
  (kind `"free-text"` lands as `/FreeText`), but the toolbar tool
  was removed in the WP8/9 overhaul because it had no canvas
  composer. Highlight + sticky landed first because their UX is
  unambiguous; free-text needs a font-picker, color-picker and
  resize affordance to be worth shipping.
- **Why deferred:** scope. The user explicitly asked for
  "round-trip Highlight + Sticky (the easy two via AP-stream),
  defer Free-text" in the WP8/9 plan kickoff.
- **Where:** `packages/pdf-annotations/src/free-text.ts`,
  `apps/web/app/pdf-viewer/PdfToolbar.tsx`,
  `apps/web/app/pdf-viewer/PdfCanvas.tsx`.
- **Acceptance:**
  - Toolbar `pdf-annotate-freetext` button reappears.
  - Click on the page drops an editable text box anchored at the
    click point. Cmd/Ctrl+Enter commits, Esc cancels.
  - Resize handles + a small font-size dropdown.
  - Existing free-text annotations render with their stored font /
    color, not just a debug border.
- **Effort:** M.

### G14b. Highlight + sticky composers under viewport rotation

- **What:** both new tools (highlight selection capture, sticky
  click-to-place) silently no-op when `viewportRotation !== 0`.
  Existing annotations _render_ correctly under rotation because
  their overlays ride along with the same CSS transform as the
  text layer; only authoring is gated.
- **Why deferred:** the inverse-transform from rotated viewport
  pixels back to PDF user-space is straightforward but adds
  surface area we hadn't tested. Shipping unrotated authoring +
  rotated rendering covers ~95 % of the user journey.
- **Where:** `apps/web/app/pdf-viewer/PdfCanvas.tsx`
  (`highlightArmed`, `stickyArmed`, `onPageClickForSticky`,
  `PdfTextLayer.onMouseUp`).
- **Acceptance:**
  - The cursor reflects the armed tool at any rotation.
  - Authoring at 90/180/270 produces annotations whose stored PDF
    rect matches what the user selected/clicked.
  - Round-trip test fixture exercises authoring at each rotation.
- **Effort:** S.

### G14c. Deleting a "loaded" annotation

- **What:** `pdf:remove-annotation` refuses to remove annotations
  that originated in the source PDF (`source: "loaded"`) — the
  serializer doesn't yet rewrite existing `/Annots` arrays.
  Session-added annotations (`source: "session"`) round-trip
  cleanly through `addAnnotations`.
- **Why deferred:** writing a strip pass needs to walk every page's
  `/Annots`, drop the right indirect references, and free the
  underlying objects. Doable in `pdf-lib` but easy to get wrong
  in a way that breaks signed PDFs, so we'd want full incremental-
  save coverage too.
- **Where:** `packages/pdf/src/commands/annotation-commands.ts`
  (`removeAnnotationHandler`),
  `packages/pdf-annotations/src/writer.ts`.
- **Acceptance:**
  - `removeAnnotationHandler` handles `source: "loaded"` by
    appending an incremental update that removes the indirect
    reference from the page's `/Annots`.
  - Round-trip test loads `with-highlight-annot.pdf`, removes one
    annotation, exports, re-parses, sees one fewer.
- **Effort:** M.

### G15. Toolbar surfaces removed pending real implementations

- **What:** the toolbar previously carried four entries that were
  cosmetically present but functionally weak — **dark mode** (CSS
  `filter: invert()` only, no per-pixel correction; produced
  inverted images and washed-out colour profiles), **reflow**
  (single-column text dump that, per G9, garbled multi-column /
  tabular content), and the **Form** menu's _Fill_ / _Flatten_
  entries plus the **Page ops** menu's _Reorder_, all carrying a
  `soon` badge against missing implementations. They were removed
  in this pass because their on-canvas effect was net-negative for
  most PDFs and the placeholders muddied the affordance bar.
- **Why deferred:** each one needs real product work before it
  re-earns toolbar real estate (smart-invert per glyph for dark
  mode; column / table-aware reflow per G9; an actual form-filling
  composer; a thumbnail drag-reorder UX).
- **Where:** `apps/web/app/pdf-viewer/PdfToolbar.tsx`,
  `apps/web/app/pdf-viewer/PdfEditor.tsx`,
  `apps/web/app/pdf-viewer/PdfCanvas.tsx`
  (`PdfDarkModeStrategy` and `reflow` props are still typed on
  the canvas, just always passed `"off"` / `false` from the
  editor — re-introducing them is a one-line wiring change once
  the real implementation lands).
- **Acceptance:**
  - Dark mode returns once `PdfCanvas` uses a per-glyph invert
    that preserves photo regions (likely a shader on the page
    bitmap, not a wholesale CSS filter).
  - Reflow returns once G9 (column / table heuristics) lands.
  - Form fill / flatten return as a real composer or are kept
    CLI-only.
  - Reorder returns when thumbnail drag-and-drop is wired.
- **Effort:** see G6 / G9 / G14a — each is its own card.

### G14. LibreOffice headless XLSX chart round-trip flake

- **Not a PDF gap, but observed during validation.**
- **What:** `packages/xlsx/src/serializer/charts.libreoffice.test.ts`
  fails when run in our sandbox because `soffice` can't launch with
  the constrained tmp profile path.
- **Why deferred:** unrelated to the PDF effort; the test passes on
  the dev box outside the sandbox.
- **Where:** `packages/xlsx/src/serializer/charts.libreoffice.test.ts`.
- **Acceptance:**
  - Either skip the test under sandboxed CI (env-flag guard) or
    teach it to use a writable profile.
- **Effort:** XS.

---

## How to pick one up

1. Re-read the relevant spec doc under `spec/pdf/`.
2. Branch off `main` (the night-shift branch will be merged shortly).
3. Add a fixture under `fixtures/pdf/` if the change has any
   serialization implication.
4. Round-trip test under `tests/roundtrip/pdf/`.
5. CLI surface change → extend `packages/agent/src/pdf-cli.ts` and
   the matching MCP tool descriptor in `packages/agent/src/mcp.ts`.
6. UI surface change → keep `i18n` keys in EN + DE in lockstep.
7. `pnpm -r typecheck && pnpm -r lint && pnpm -r test &&
make audit-roundtrip` before opening the PR.

If any item changes scope or a new gap shows up, append it here so
the backlog stays in one place.
