# PDF — Performance

> Measurable budgets for cold load, page render, memory, and search.
> Enforced by perf CI gates.

Cross-references: rendering pipeline in
[`rendering-pipeline.md`](./rendering-pipeline.md);
search budgets in [`search.md`](./search.md);
edge-case stress in [`edge-cases.md`](./edge-cases.md);
acceptance criteria in [`acceptance-criteria.md`](./acceptance-criteria.md).

## Targets

| Metric                                                     | Target                                  |
| ---------------------------------------------------------- | --------------------------------------- |
| Cold load (first paint) for 50-page text PDF (~2 MB)       | **< 600 ms p95** on a Mac M1 over local file |
| Cold load for 50-page mixed-content PDF (~10 MB)           | **< 1.5 s p95**                         |
| Page render (visible page → bitmap)                        | **< 150 ms p95**                        |
| Scroll FPS during continuous scroll on 1000-page PDF       | **60 FPS** sustained (no drops > 20 ms in any 5 s window) |
| Memory cap (RSS) on 1000-page PDF                          | **< 600 MB**                            |
| Memory cap (RSS) on 200 MB PDF (stress)                    | **< 1.5 GB**                            |
| Search: first hit on 500-page text PDF                     | **< 200 ms p95**                        |
| Search: all hits on 500-page text PDF                      | **< 800 ms p95**                        |
| Annotation create + render                                 | **< 50 ms p95**                         |
| Bus dispatch + handler + diff                              | **< 5 ms p95** for typical commands     |
| Outline parse for 1000-entry outline                       | **< 100 ms**                            |
| Thumbnail generation (lazy, per page)                      | **< 80 ms**                             |
| Bundle (eager) — viewer chrome + PDF.js                    | **< 800 KB gzipped**                    |
| Bundle (lazy) — PDFium WASM, tesseract, fonts              | budgets in [`engine-strategy.md`](./engine-strategy.md) |

These targets are the floor for "feels faster than Chrome's built-in
viewer". Anything slower is a regression.

## Virtualization strategy

The renderer uses an aggressive virtualization budget to keep memory
and frame time bounded — see
[`rendering-pipeline.md`](./rendering-pipeline.md):

- **Render window**: viewport pages **± 2**. Outside the window,
  pages are placeholder DOM elements sized from
  `PdfPage.width × PdfPage.height`. The full document scroll height
  is correct from parse time, so the scrollbar never jumps.
- **Text-layer window**: viewport pages **± 1**. The DOM cost per
  page is dozens to thousands of `<span>`s; trimming aggressively
  keeps DOM under 30 K nodes even for a 1000-page doc.
- **Raster cache**: 200 MB LRU, page-granular eviction.
- **Operator-list cache**: 20 pages LRU. Re-renders on
  rotate/zoom/dark-mode-toggle reuse the operator list, skipping
  the worker round-trip.

## Memory model

- **Original buffer** is retained for the life of the agent (needed
  for incremental save). For a 200 MB PDF that's 200 MB locked.
- **Engine state** holds the parsed cross-reference + a few
  per-page caches; ~50 MB for a typical 1000-page doc.
- **Raster cache** caps at 200 MB.
- **Text indexes** are ~3-5 KB per page; ~5 MB for a 1000-page doc.
- **Annotations + form fields + comments** are sub-megabyte for any
  realistic document.

Total RSS for a 1000-page text PDF: ~280 MB original buffer + ~50
MB engine + ~200 MB raster + ~5 MB indexes + UI + browser overhead =
~600 MB. We hit the budget.

For the 200 MB PDF stress fixture: ~200 MB original + ~80 MB engine
+ ~200 MB raster + ~5 MB indexes + UI ≈ 1.5 GB. We hit the stress
budget.

## Threading discipline

The 16 ms-per-frame budget on the main thread is **non-negotiable**.
The architecture check
([`scripts/check-architecture.mjs`](../../scripts/check-architecture.mjs))
enforces no direct `pdfjs-dist` import outside
`packages/pdf-engine/src/backends/`. This means parsing,
operator-list interpretation, font decoding, image decoding, text
extraction, search-index construction, and OCR all run in workers.

Main thread work, per frame:

1. Hit testing on pointermove (≤ 1 ms; spatial index per page).
2. DOM commit for newly visible pages (≤ 5 ms).
3. Smart-invert WebGL pass for 1-2 newly-painted pages (≤ 4 ms total).
4. Composite the visible canvases (browser-internal, ≤ 4 ms).

Total: ≤ 14 ms with 2 ms slack.

## OffscreenCanvas usage

Where supported (~95% of modern browsers):

- **Thumbnail generation** runs on `OffscreenCanvas` in a dedicated
  worker pool of 2-4 workers. Thumbnails for off-viewport pages are
  generated eagerly during idle time (`requestIdleCallback`).
- **Smart-invert pre-pass** for the next ±2 pages can run on
  `OffscreenCanvas` so the result is ready when the page enters the
  viewport.

Fallback: same code paths run on the main thread `<canvas>` with a
yield to the event loop between pages.

## Streaming and range requests

PDF.js supports HTTP range requests. When loading from a URL:

- Initial load fetches the linearization hint (first 64 KB) +
  trailer + xref table.
- Page rendering triggers per-page byte fetches as needed.
- For non-linearized PDFs, the full file is fetched up-front.

The agent's CLI runs against local files, so range requests don't
apply there — but the in-browser viewer benefits significantly for
large remote files.

## CI gates

Three perf jobs:

### `perf:cold-load`

- Boots a headless Chromium via Playwright.
- Loads each fixture from `fixtures/pdf/synthetic/`.
- Measures first-page TTI via `performance.measure`.
- Fails if any fixture exceeds the cold-load budget by > 10%.

### `perf:scroll-fps`

- Same harness; opens the 1000-page stress fixture.
- Auto-scrolls top → bottom over 30 s.
- Captures a Chrome DevTools performance trace.
- Asserts no frame > 20 ms in any 5 s window.

### `perf:memory`

- Same harness; opens the 200 MB stress fixture.
- Scrolls to mid-document, lets the LRU stabilize.
- Reads `chrome://process-internals` (or `Performance.memory`) for
  RSS.
- Fails if RSS exceeds 1.5 GB.

### `perf:bundle`

- `bundlesize` config in `apps/web/package.json`.
- Fails on any bundle exceeding its budget in
  [`engine-strategy.md`](./engine-strategy.md).

## Headless / agent perf

The CLI's perf characteristics:

| CLI command                              | Budget for 50-page PDF        |
| ---------------------------------------- | ----------------------------- |
| `office-agent pdf inspect`               | < 800 ms                      |
| `office-agent pdf metadata`              | < 400 ms                      |
| `office-agent pdf read --pages all`      | < 2 s                         |
| `office-agent pdf render --dpi 150 all`  | < 15 s (~300 ms / page)       |
| `office-agent pdf rotate --pages 1`      | < 600 ms (load + edit + save) |
| `office-agent pdf merge --files 5`       | < 3 s                         |

Measured by `tests/agent/pdf/perf.test.ts` on CI.

## Known performance traps

- **Naive text-layer**: building `<span>`s for all pages eagerly
  blows the DOM budget. We strictly window to ±1.
- **Full-document re-render on dark-mode toggle**: avoided by
  preserving the operator-list cache and only re-running the
  smart-invert pass.
- **Unbounded LRU**: solved by 200 MB cap with page-granular
  eviction.
- **Search regex catastrophic backtracking**: timeout the search
  worker at 5 s; user message rather than UI hang.
- **Loading 200 MB into memory twice** (once as the original buffer,
  once as the engine's parsed state): the engine's parsed state
  doesn't duplicate the original buffer — it indexes it.

## Cold-start UX

Within the first 600 ms after `agent.fromBuffer()`:

1. **0-50 ms**: Skeleton page placeholders sized from a bytes-only
   estimate (mean page size from the linearization hint).
2. **50-200 ms**: Real page sizes from the engine's
   `getPage(1).info`; skeleton resizes are smooth.
3. **200-400 ms**: First page rasterized; outline parsed in the
   background.
4. **400-600 ms**: Thumbnails for the first 4 pages eager-generated;
   text-layer for page 1 mounted; user can begin scrolling /
   selecting / searching.

The user-perceived first paint is the moment step 3 commits. For
text-heavy 50-page PDFs we beat the < 600 ms budget consistently on
mid-range hardware.
