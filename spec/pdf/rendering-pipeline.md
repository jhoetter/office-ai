# PDF — Rendering Pipeline

> Page render lifecycle, layer composition, devicePixelRatio handling,
> and virtualization budget for the browser viewer. The headless Node
> renderer follows the same lifecycle without the DOM layers.

Cross-references: engine choice in [`engine-strategy.md`](./engine-strategy.md);
text-layer detail in [`text-layer.md`](./text-layer.md);
annotation overlay in [`annotation-model.md`](./annotation-model.md);
performance budgets in [`performance.md`](./performance.md).

## Threads

1. **Worker** — owns parsing, content-stream interpretation,
   operator-list construction, font program decoding, image stream
   decoding (DCTDecode, JBIG2, JPX, CCITT, …), text-content extraction,
   and search-index pre-building. Owned by `pdfjs-dist`'s built-in
   worker for the PDF.js backend; owned by a custom WASM worker for
   the PDFium backend.
2. **Main** — owns canvas painting, hit testing, DOM layer composition,
   user input, and the LRU raster cache.

The main thread budget is **16 ms per frame**. Parsing or decoding on
the main thread is forbidden. The architecture check enforces no
direct `pdfjs-dist` import outside `packages/pdf-engine/src/backends/`.

## Page render lifecycle

For a page entering the viewport's render window:

```
intent → check-cache → request-from-engine → rasterize → paint →
  → text-layer → annotation-layer → struct-tree-layer → presence-cursors
```

### 1. Intent

`PdfCanvas` watches the IntersectionObserver for each placeholder
page element. When a page enters the render window (viewport ± 2
pages), it issues a `RenderIntent` to the orchestrator with `{
pageNumber, scale, rotation, viewportRect }`.

### 2. Check cache

The LRU raster cache is keyed by `(pageNumber, devicePixelScale,
rotation, engineKind)`. Hit → bitmap is reused; the page paints
synchronously. Miss → enqueue a render task.

### 3. Request from engine

The orchestrator calls `enginePage.render({ scale, canvas })`. The
engine's worker decodes the page operator-list and replays it onto
the supplied canvas. For PDF.js this happens off the main thread by
posting the rasterization task to the worker; for headless Node the
rasterization happens on `@napi-rs/canvas`.

### 4. Rasterize

The bitmap is committed to the LRU cache. If the cache exceeds its
memory budget (default **200 MB**), least-recently-painted entries
are evicted. Eviction is page-granular; we never partially evict a
page.

### 5. Paint

The bitmap is drawn into the on-screen `<canvas>` for the page.
First paint is the moment the user sees the page. Performance budget:
**< 150 ms p95** (see [`performance.md`](./performance.md)).

### 6. Text-layer

`enginePage.getTextContent()` returns `{ items, plain }`. Each item
becomes an absolutely-positioned, font-sized `<span>` overlaid on the
canvas. The text layer is selectable, copyable, and search-highlight
target. Detail: [`text-layer.md`](./text-layer.md).

### 7. Annotation-layer

`enginePage.getAnnotations()` returns lightweight projections. The
viewer composes:

- **Native annotations** rendered from their `/AP` streams (PDF.js
  does this for free; PDFium output is composited into the same
  layer).
- **Office-AI overlays** for in-progress annotations (highlight in
  flight, ink stroke being drawn, comment region indicator).

Detail: [`annotation-model.md`](./annotation-model.md).

### 8. Struct-tree-layer

If the page is tagged, the engine returns a struct-tree fragment.
The viewer renders it as hidden ARIA-labelled DOM aligned with the
text layer for screen readers. Detail: [`accessibility.md`](./accessibility.md).

### 9. Presence cursors

Other readers' cursors arrive over `useCommandBroadcast` /
`usePublishPresence`. Each is rendered as a labelled cursor sprite
in the same coordinate space as the canvas.

## DevicePixelRatio handling

The on-screen canvas uses `width = cssWidth * devicePixelRatio` and
the 2D context is scaled by `dpr`. The bitmap from the engine is
rendered at `scale = baseScale * devicePixelRatio` where `baseScale`
is the user's zoom (1.0 == 100%). This produces crisp text at retina
densities without blurring.

**Caveat:** at high zoom + high DPR (e.g. 400% on a 3x display) the
bitmap area can exceed `Canvas.maxAllowedSize` (browser-dependent,
~16k square). The orchestrator clamps `effectiveScale` to a maximum
that keeps the bitmap under the limit and overlays the text layer at
the requested zoom. Visual fidelity degrades smoothly rather than
hard-failing.

## Virtualization budget

| Window           | Budget                                   |
| ---------------- | ---------------------------------------- |
| Render window    | viewport pages **± 2**                   |
| Text-layer window| viewport pages **± 1**                   |
| Annotation window| viewport pages **± 2**                   |
| Struct-tree win  | viewport pages **± 1**                   |
| Raster cache     | **200 MB** RSS, page-granular LRU        |
| Operator-list cache | **20 pages** LRU (faster re-render on rotate / zoom changes) |

Pages outside the render window are placeholders sized from
`PdfPage.width × PdfPage.height` so the scroll height is correct from
parse time. This means the scrollbar never jumps and `Ctrl+End`
lands on the last page reliably.

## Print rendering

Print uses a separate render pass that disables the LRU cache and
streams pages at the user's chosen DPI (default 150) into the
browser print bridge. The print pass holds at most one page bitmap
in memory at a time so a 1000-page print job doesn't blow RSS.

## Headless rendering (Node)

`packages/pdf-engine` exposes the same `enginePage.render(opts)`
contract in Node. The headless backend uses
`pdfjs-dist/legacy/build/pdf.mjs` + `@napi-rs/canvas`, returning a
`Uint8Array` (PNG / JPEG / WebP). This is what
`office-agent pdf render` and `office-agent pdf thumbnail` call —
no DOM, no React, no browser.

## Cancellation

Render tasks are cancellable. When a page leaves the render window
before its bitmap is committed, the orchestrator cancels the engine
render task and discards any partial output. The engine returns a
fresh task on the next intent.

## Failure modes

| Failure                                | Handling                                                           |
| -------------------------------------- | ------------------------------------------------------------------ |
| Engine throws on `getPage(N)`          | Render placeholder shows "Page N could not be rendered" + retry button.|
| Engine returns blank text content      | Page rendered without text-layer; OCR opt-in banner shown if scan-detected.|
| Bitmap exceeds canvas size limit       | `effectiveScale` clamped; visible warning at zoom level.           |
| Cache eviction during in-flight render | The render completes, then is immediately evicted if budget exceeded. Caller is unaware.|
| Worker crashes (PDF.js)                | Engine destroyed and re-created lazily; in-flight intents replayed. |
