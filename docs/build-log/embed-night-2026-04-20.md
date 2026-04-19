# Cross-format embed (Phase 5) — Night 2026-04-20

This phase ships a structured-handoff envelope that lets a copy in
one editor paste into a different editor with full semantic
fidelity. Today this means **XLSX range → DOCX table** and **XLSX
range → PPTX text box**. The whole feature is gated on the
`NEXT_PUBLIC_OAI_EMBED` flag so the extra clipboard MIME type
doesn't ship to production users until QA signs off.

## What ships

| Source                | Target | Result                              | Status |
| --------------------- | ------ | ----------------------------------- | ------ |
| XLSX range            | XLSX   | Existing `<table data-fingerprint>` | Shipped (untouched) |
| XLSX range            | DOCX   | Typed `<w:tbl>` with cell text      | **New** |
| XLSX range            | PPTX   | New `<p:sp>` text box, TSV-rendered | **New** |
| XLSX chart            | DOCX/PPTX | PNG paste                        | Backlog (see below) |
| Other → other         | …      | falls through to existing handlers  | Unchanged |

## Wire format

```
MIME : application/x-officeai-embed+json
Body : { type: "officeai/embed", version: 1, source, createdAt, payload }
```

Payload variants (`payload.kind`):

* `xlsx-range` — an `XlsxClipboardSnapshot` plus an `originLabel`.
  Pasted into XLSX, the existing fingerprint path catches it first
  (lossless). Pasted into DOCX/PPTX, the cross-format helpers
  consume the snapshot directly without a re-parse hop.
* `xlsx-chart-image` — base64-PNG + intrinsic dimensions + chart
  kind. Designed for the future chart-to-image flow; the type and
  parser are in place so the receiving DOCX/PPTX handlers can be
  added incrementally without touching the envelope.

The envelope is written **alongside** `text/html` and `text/plain`
(not instead of), so external apps (Excel desktop, Sheets, plain
editors) keep getting the formats they understand.

## Files

```
apps/web/app/lib/embed/
  envelope.ts                         # MIME + envelope + (de)serialise + flag
  envelope.test.ts                    # 5 tests
  applyXlsxRangeToDocx.ts             # docx:insert-table + setCellContent
  applyXlsxRangeToDocx.test.ts        # 3 tests, headless DocxAgent
  applyXlsxRangeToPptx.ts             # pptx:add-text-box, TSV-rendered
  applyXlsxRangeToPptx.test.ts        # 2 tests, headless PptxAgent
```

Wired into editors:

* `apps/web/app/xlsx-editor/clipboard.ts` — `marshalClipboard` adds
  the embed string when the flag is on; `writeToSystemClipboard`
  carries it through the `ClipboardItem` path.
* `apps/web/app/xlsx-editor/XlsxEditor.tsx` — synchronous
  `onSurfaceCopy` / `onSurfaceCut` paint the embed MIME via
  `setData` so non-async paths still get it.
* `apps/web/app/editor/DocxEditor.tsx` — paste handler checks the
  embed MIME first, falls back to the existing image-paste path.
* `apps/web/app/pptx-editor/PptxEditor.tsx` — window-level paste
  listener (the slide canvas isn't tab-focusable; window scope keeps
  Cmd-V working from anywhere except form fields).

## How to try it locally

```bash
NEXT_PUBLIC_OAI_EMBED=1 pnpm --filter @officeai/web dev
```

1. Open `localhost:3000`, pick "Spreadsheet", load a sample.
2. Select a 3×3 range, Cmd-C.
3. Open `localhost:3000` in another tab, pick "Word document".
4. Cmd-V — a real `<w:tbl>` appears at the caret, undo-able as a
   single user action and visible to other realtime peers (Phase 1).
5. Repeat with "Presentation" — pastes as a TSV-rendered text box
   on the active slide.

With the flag off (the production default) the existing PM HTML
table paste / clipboard image paste paths run unchanged.

## Deliberate cross-format downgrades

The first ship picks correctness-by-construction over fidelity:

* **DOCX**: Cell formulas render as `=…` text. Style ids and merge
  regions are dropped — the typed `<w:tbl>` doesn't yet carry style
  fields, so re-emitting them through `set-cell-content` would be a
  lie. Column widths are a uniform split clamped to fit a Letter
  body width.
* **PPTX**: Renders as a single TSV text box because there is no
  `pptx:insert-table` command yet (only edits to existing
  `<a:tbl>`). The pasted box is sized 80% × 60% of the slide and
  named `XLSX paste Sheet1!A1:C5` so the user can find it in the
  shape list.

## Backlog (deferred to follow-ups, in priority order)

1. **`pptx:insert-table` command** so the PPTX paste produces a
   real `<a:tbl>`. Mostly mirroring the DOCX `insert-table` shape
   plus DrawingML's `<a:tbl>` skeleton; the round-trip path is the
   harder part because PowerPoint's table styles tree is large.
2. **DOCX cell styles**: pipe `EffectiveStyle` from
   `XlsxClipboardCell.styleId` into `RunProperties` so bold /
   italic / fill survive the cross-format paste.
3. **Chart → PNG**: render the `SheetChart` via the existing
   on-screen SVG path, rasterise via `<canvas>`, base64 it onto
   `xlsx-chart-image` payloads, and let the DOCX/PPTX handlers call
   the existing image-insert paths. The envelope is already in
   place (`payload.kind === "xlsx-chart-image"`), only the producer
   and consumer wiring is missing.
4. **Reverse direction**: DOCX table copy → XLSX range. Falls out
   of the envelope shape; needs a `docx-table` payload variant.

## Test coverage

```
apps/web vitest:                 10 / 10  passing (envelope + apply helpers)
pnpm typecheck (root):           17 / 17  green
pnpm lint:root:                  0 errors (2 pre-existing warnings)
```

Smoke-test recipe (manual): copy a range from the XLSX sample,
paste into the DOCX sample, save, re-open in Word — table renders
with the same row × column shape and cell text content.
