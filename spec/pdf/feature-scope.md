# PDF — Feature Scope

> The 10/10 PDF viewer feature inventory, bucketed by priority.
> Read this first.

| P-code  | Meaning                                                                              |
| ------- | ------------------------------------------------------------------------------------ |
| **P0**  | Lands this night-shift session. Tested. Documented.                                  |
| **P1**  | Lands as a stub or opaque-preservation only this session. Editing deferred.          |
| **P2**  | Out of scope this session. Spec-only. Not parsed/serialized beyond byte-preservation.|
| **OUT** | Explicitly excluded by [`prompt-pdf.md`](../../prompt-pdf.md) §"Explicitly Out".     |

The buckets mirror the architectural waves in
[`/Users/jhoetter/.cursor/plans/pdf_viewer_night_shift_71063370.plan.md`](https://example.invalid/local-plan).
Cross-references: [`agent-commands.md`](./agent-commands.md) for the
command surface; [`acceptance-criteria.md`](./acceptance-criteria.md)
for the test plan; [`engine-strategy.md`](./engine-strategy.md) for
engine selection.

## Read

| Feature                                                              | Status | Notes                                                       |
| -------------------------------------------------------------------- | :----: | ----------------------------------------------------------- |
| Open PDF 1.4 → 2.0 from URL, file, drag-drop, paste, share-target    | **P0** | Via `pdf-engine` PDF.js backend.                            |
| Continuous, single-page, two-page (book/cover) view modes            | **P0** | `PdfCanvas` view-mode toggle.                               |
| Pinch / Ctrl+wheel / fit-width / fit-page / actual / custom % zoom   | **P0** | See [`keyboard-shortcuts.md`](./keyboard-shortcuts.md).     |
| Pan via hand tool, arrow keys, Page Up/Down, Home/End                | **P0** |                                                             |
| Per-page rotation (display + persisted via `pdf:rotate-pages`)       | **P0** |                                                             |
| Sidebar: outline, thumbnails, annotations, search, comments         | **P0** |                                                             |
| Sidebar: attachments tab                                             | **P0** | Read-only listing.                                          |
| Page-number jump (Ctrl/Cmd+G)                                        | **P0** |                                                             |
| Goto-link + URI link annotation navigation                           | **P0** |                                                             |
| Named destinations                                                   | **P1** | Resolved on read; no editing.                               |
| Document properties dialog (metadata, security, fonts used)          | **P0** |                                                             |
| Print (page selection, scaling, even/odd)                            | **P0** | Browser print bridge; no booklet imposition.                |
| Booklet imposition print mode                                        | **P2** |                                                             |
| Download original / download modified                                | **P0** |                                                             |
| Persistent reading position per document                             | **P0** | Cookie / localStorage keyed by `partHashes` digest.         |
| Reading ruler / line-focus mode                                      | **P1** |                                                             |
| Distraction-free / chrome-hide mode                                  | **P0** | `Esc` toggles.                                              |
| Read-aloud (TTS with sentence highlight)                             | **P2** |                                                             |
| Mini-map (slim ribbon)                                               | **P2** |                                                             |
| "Where am I" outline-section indicator                               | **P1** |                                                             |
| Three-finger forward/back gesture                                    | **P2** |                                                             |

## Annotate

All annotations roundtrip into native PDF annotation objects with
valid AP streams. Wire model: [`annotation-model.md`](./annotation-model.md).

| Feature                                                              | Status |
| -------------------------------------------------------------------- | :----: |
| Highlight, underline, strikethrough, squiggly                        | **P0** |
| Sticky note / popup with thread, replies, resolve                    | **P0** |
| Free-text annotation                                                 | **P0** |
| Free-hand ink (smoothing; pressure where supported)                  | **P0** |
| Shapes: line, arrow, rectangle, ellipse, polygon, polyline           | **P0** |
| Stamp (preset "Approved" / "Draft" / custom image)                   | **P0** |
| Caret / insert-text annotation                                       | **P1** |
| File-attachment annotation                                           | **P1** |
| Link annotation (URI / goto-page / goto-named-destination)           | **P0** |
| Redaction (visual blackout + content-stream removal + metadata scrub) | **P0** |
| Color, opacity, line width, font, line dash for every applicable     | **P0** |
| Annotations panel: filter by author/type/page/date; resolve; reply   | **P0** |
| Real-time annotation cursor + live ink streaming                     | **P2** |
| FDF/XFDF import/export                                               | **P0** | XFDF first; FDF best-effort.                                |

## Edit (page-level)

Backed by `packages/pdf-edit` over `pdf-lib`. Wire model:
[`editing-pipeline.md`](./editing-pipeline.md).

| Feature                                                              | Status |
| -------------------------------------------------------------------- | :----: |
| Rotate single page or selection (90 / 180 / 270)                     | **P0** |
| Reorder pages (drag in thumbnail panel; `pdf:reorder-pages`)         | **P0** |
| Insert blank page / insert pages from another PDF                    | **P0** |
| Delete pages (`pdf:delete-pages`)                                    | **P0** |
| Extract pages (save as new PDF)                                      | **P0** |
| Split (by range / by size / by bookmark)                             | **P0** |
| Merge multiple PDFs                                                  | **P0** |
| Crop pages (margin presets; per-page or per-range)                   | **P0** |
| Add watermark / stamp / page numbers / header / footer               | **P0** |
| Resize / change page size                                            | **P1** |
| In-page text editing (re-flow existing rasterized text)              | **OUT** | Tar pit — overlay + redaction only per `prompt-pdf.md`.    |

## Forms

| Feature                                                              | Status |
| -------------------------------------------------------------------- | :----: |
| Render every AcroForm field type                                     | **P0** |
| Fill, validate (`/MaxLen`, regex), respect calc-order                | **P0** |
| Save filled (preserve as fillable)                                   | **P0** |
| Flatten on save (bake into content stream)                           | **P0** |
| Reset form                                                           | **P0** |
| Import/export FDF/XFDF                                               | **P0** |
| AcroForm-fallback render for XFA-bearing PDFs                        | **P0** |
| Sandbox embedded JS (`/AA`, `/F`, document-level)                    | **P0** | No execution.                                               |
| XFA dynamic forms beyond AcroForm fallback                           | **OUT** | Adobe is deprecating.                                       |
| Cryptographic signing (PKCS#12)                                      | **P2** | Visible signature image only this session.                  |

## Search

Wire model: [`search.md`](./search.md).

| Feature                                                              | Status |
| -------------------------------------------------------------------- | :----: |
| Per-page text index built on parse                                   | **P0** |
| Hit highlighting in text-layer; hit count; prev/next                 | **P0** |
| Match case toggle                                                    | **P0** |
| Match whole word toggle                                              | **P0** |
| Regex toggle                                                         | **P0** |
| Fuzzy / typo-tolerant search                                         | **P1** |
| Semantic search                                                      | **P2** |
| Search across multiple open documents                                | **P2** |

## Accessibility

Wire model: [`accessibility.md`](./accessibility.md).

| Feature                                                              | Status |
| -------------------------------------------------------------------- | :----: |
| Tagged-PDF struct-tree → ARIA-labelled hidden DOM                    | **P0** |
| Reflow mode (single-column accessibility view)                       | **P0** |
| Full keyboard navigation; every action reachable                     | **P0** |
| Skip-links to toolbar, sidebar, page content                         | **P0** |
| ARIA labels on all UI controls                                       | **P0** |
| WCAG 2.2 AA conformance for chrome (axe-core gate)                   | **P0** |
| `prefers-reduced-motion` honored                                     | **P0** |
| High-contrast mode                                                   | **P1** |
| Auto-tag proposal (PDF/UA generation)                                | **P2** |

## AI-via-CLI

Wire model: [`cli.md`](./cli.md). **The CLI is the AI integration
surface for this product.** No LLM logic ships in-product; all AI
features are agents calling the CLI / MCP from outside.

| Feature                                                              | Status |
| -------------------------------------------------------------------- | :----: |
| `office-agent pdf inspect / metadata / outline / list-pages`         | **P0** |
| `office-agent pdf list-annotations / list-form-fields / list-fonts`  | **P0** |
| `office-agent pdf list-attachments / list-signatures`                | **P0** |
| `office-agent pdf read --pages / --bbox --format md\|json\|text`    | **P0** |
| `office-agent pdf search --query [--regex --case]`                   | **P0** |
| `office-agent pdf chunk --strategy outline\|page\|fixed-tokens`     | **P0** |
| `office-agent pdf render / thumbnail / extract-images`               | **P0** |
| Page-ops CLI: rotate / reorder / delete / insert / merge / split /    | **P0** |
| extract / crop / watermark / page-numbers / set-metadata              |        |
| `office-agent pdf import-annotations / export-annotations`           | **P0** |
| `office-agent pdf fill-form / reset-form / flatten-form`             | **P0** |
| `office-agent pdf redact --pattern\|--rects --log`                   | **P0** |
| `office-agent pdf ocr --lang`                                        | **P0** |
| `office-agent pdf apply / diff`                                      | **P0** |
| MCP tools mirroring all of the above                                 | **P0** |
| In-product chat / summarize / extract / compare / smart-redact       | **OUT** | Per user clarification: CLI is the surface.                |

## Collaboration

Reuses `packages/realtime` + `packages/comments`.

| Feature                                                              | Status |
| -------------------------------------------------------------------- | :----: |
| Presence cursors per page (`usePublishPresence`)                     | **P0** |
| Annotation sync via Y.js + command broadcast                         | **P0** |
| Comments anchored as `pdf-region` (page + normalized rect)           | **P0** |
| @-mentions in comments                                               | **P1** |
| Resolve / reopen comment threads                                     | **P0** |
| Activity feed                                                        | **P2** |
| Live remote ink-stroke streaming                                     | **P2** | Eventual via command-bus; live streaming is polish follow-up. |

## Roundtrip-integrity bar

All features above (P0, P1) must satisfy:

- **Open + no-edit save** is byte-identical to the input (incremental
  save with zero deltas).
- **Open + P0 edit + save** changes only the touched objects; every
  other object is byte-identical to the input. Verified by SHA-256
  over `partHashes`-equivalent regions of the byte buffer.
- **Open in Acrobat / Preview / Chrome** after a P0 edit shows no
  "this file was modified / repair needed" dialog.
- **Signed PDFs** edited via incremental save keep their signatures
  valid for the prior byte range.

This is the **only** non-negotiable acceptance criterion. See
[`acceptance-criteria.md`](./acceptance-criteria.md).

## Deferred items (called out up-front)

- **Real-world fixture corpus** is replaced by 20 synthetic Python
  fixtures for night shift; flag in `docs/build-log/pdf/fixtures.md`
  and prioritize replacement before production.
- **PDFium-WASM fallback** wired structurally and exercised on 1–2
  fidelity-critical fixtures, not all 20.
- **Cryptographic signing** — read-only signature verification ships;
  visible signature images ship; PKCS#12 signing deferred.
- **Cross-format embed** (PDF → DOCX text paste, etc.) — defer; the
  embed envelope ([`apps/web/app/lib/embed/envelope.ts`](../../apps/web/app/lib/embed/envelope.ts))
  needs a new `pdf-text` variant.
- **Read-aloud / TTS / mini-map** — polish, deferred.

## Tracking

`docs/build-log/pdf.md` keeps a row per feature with status, test
reference, and any deviation from this spec.
