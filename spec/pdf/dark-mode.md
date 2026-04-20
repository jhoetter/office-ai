# PDF — Dark Mode

> Smart-invert algorithm: luminance-only invert for backgrounds and
> text, image preservation for photos and figures, opt-out per
> document.

Cross-references: rendering pipeline integration in
[`rendering-pipeline.md`](./rendering-pipeline.md);
keyboard binding in [`keyboard-shortcuts.md`](./keyboard-shortcuts.md);
accessibility interplay in [`accessibility.md`](./accessibility.md).

## Why this matters

Most PDF viewers' dark modes are naive: they invert the entire
rendered bitmap (`filter: invert(1) hue-rotate(180deg)`). The result:
photos turn into negatives, figures get unreadable, code blocks
flicker, and color-coded charts become meaningless. Adobe Acrobat
Web's dark mode does this. Chrome's PDFium viewer doesn't have a
dark mode at all. Shadow Reader is the only widely-known PDF reader
that does smart-invert correctly, and it's a Mac-only paid app.

We bake smart-invert in by default. It's the visible "10/10" win
the moment a user opens a PDF in low light.

## Algorithm

The smart-invert pass operates **per pixel** on the rasterized
canvas, after the engine has painted the page. It is implemented as
a WebGL fragment shader (with a CPU fallback) executed in a single
post-processing pass.

### Per-pixel decision

Let the source pixel be `(r, g, b)` in 0..1.

1. Compute **luminance** using BT.709 weights:

   ```
   Y = 0.2126·r + 0.7152·g + 0.0722·b
   ```

2. Compute **chroma magnitude** (max channel difference):

   ```
   chroma = max(r, g, b) − min(r, g, b)
   ```

3. **Classify** the pixel:

   - `chroma < CHROMA_BG` (default 0.04) → **near-grayscale**.
     - If `Y > Y_BG` (default 0.85) → **background** (paper white).
       Output: dark surface color (`#1a1a1a` token `--pdf-paper-dark`).
     - If `Y < Y_TEXT` (default 0.30) → **text** (near-black).
       Output: light text color (`#e8e8e8` token `--pdf-ink-dark`).
     - Else → **gray UI element** (rules, table borders, light
       shading). Output: invert luminance only — `(1−Y, 1−Y, 1−Y)`.
   - `chroma ≥ CHROMA_BG` → **chromatic** content (photo, figure,
     icon, color-coded chart). Output: **pixel preserved**.

4. **Smooth** the boundary between classes via a 1-pixel-wide
   blend zone to avoid hard edges around antialiased glyphs.

The thresholds (`CHROMA_BG`, `Y_BG`, `Y_TEXT`) are tunable via
CSS custom properties so a power user can adjust per-document if
needed.

### Image preservation

Even within "chromatic" pixels, we want photos to look untouched.
The shader treats any pixel with `chroma ≥ CHROMA_BG` as preserved —
no inversion, no luminance shift. This means:

- A black-and-white photo (low chroma but image-like) **does** get
  inverted (it lives in the near-grayscale path). This is sometimes
  imperfect for monochrome photos. Mitigation below.
- A color photo, a syntax-highlighted code listing, a chart, a
  figure with colored shapes, an icon — all preserved.

For **monochrome photos** specifically, we use the struct tree when
present: any pixel inside a `/Figure` rect bypasses the inversion
entirely. The shader receives a "preserve mask" sampled from a
per-page R8 texture marking `/Figure` regions. For untagged PDFs the
heuristic above (chroma threshold) is the only signal.

## Where the shader runs

- **WebGL2** path (preferred): the rasterized canvas is uploaded as
  a texture, the shader runs the smart-invert pass, the result is
  drawn to the visible canvas. Cost: one GPU pass per page render,
  ~1-2 ms on integrated graphics.
- **CPU fallback** (`OffscreenCanvas` 2D + `ImageData` walk): used
  when WebGL2 is unavailable or when GPU memory is tight (1000-page
  document under stress). Cost: ~30-60 ms per 1000×1300 page on a
  Mac M1 — acceptable for the rare fallback case.
- **Headless Node** path: the smart-invert is **skipped** — the
  CLI's `office-agent pdf render` always emits the producer's
  original colors regardless of viewer theme. Dark mode is a
  rendering preference, not a document property.

## Toggle and persistence

| Trigger                     | Effect                                              |
| --------------------------- | --------------------------------------------------- |
| `Ctrl+Alt+D`                | Toggle dark mode globally (system + viewer chrome). |
| Page menu → "Dark mode"     | Same as `Ctrl+Alt+D`.                              |
| `prefers-color-scheme: dark`| Default to dark mode on first open.                 |
| Page menu → "Always light"  | Per-document opt-out, persisted in localStorage.    |
| Page menu → "Always dark"   | Per-document opt-in, persisted in localStorage.    |
| Page menu → "Follow system" | Default; respects `prefers-color-scheme` changes.   |

The per-document override is keyed by `partHashes`-equivalent digest
so the same document always loads with the same preference.

## Thumbnail consistency

Thumbnails in the sidebar **always** render in light mode. Dark
thumbnails make it hard to recognize a page at a glance, especially
for figure-heavy pages. The thumbnail strip uses a `--pdf-thumb-bg`
token that is white in both themes, with a 1px ring in the active-
theme accent color around the current page.

## Accessibility interplay

- `prefers-contrast: more` boosts the `Y_BG` and `Y_TEXT` thresholds
  so the result hits 7:1 contrast (WCAG AAA) rather than 4.5:1
  (AA). See [`accessibility.md`](./accessibility.md).
- The text-layer search-hit highlight uses a different yellow in
  dark mode (`bg-yellow-700/40` instead of `bg-yellow-200/60`) to
  remain visible against the dark surface.
- Annotation colors are **never** smart-inverted — the user picked
  yellow for a highlight; it stays yellow. The annotation layer is
  composited above the smart-invert pass.
- Comment indicators (the `pdf-region` glyph) use `--accent` which
  is theme-aware; visible in both modes.

## Failure modes

| Case                                            | Handling                                              |
| ----------------------------------------------- | ----------------------------------------------------- |
| WebGL2 unavailable                              | CPU fallback. Banner: "Slower dark-mode pass active." |
| GPU memory pressure during 1000-page scroll     | Drop the smart-invert pass for off-viewport pages.    |
| Pure-grayscale scan (low chroma everywhere)     | Inversion applied everywhere → photo looks wrong; user can opt-out per document. |
| Color profile mismatch (CMYK / Lab)             | Convert to RGB at engine boundary; smart-invert sees RGB. |

## Why not "filter: invert(1) hue-rotate(180deg)"

That CSS filter is the naive approach. It:

- Inverts **every** pixel, destroying photos.
- Hue-rotates the result, making colors wrong.
- Doesn't know about `/Figure` regions.
- Doesn't smooth boundaries.
- Looks bad on antialiased text edges.

We explicitly reject it. The shader path costs us a few KB of GLSL
and a few milliseconds per render — well worth it for the visible
quality difference.

## Testing

- **Visual snapshot diff** on 20 fixtures, light vs dark, across
  3 representative page types (text-only, figure-heavy, photo).
  Asserts the smart-invert preserves chromatic content and inverts
  near-grayscale.
- **Contrast measurement** on rendered pages: text vs background
  contrast ≥ 4.5:1 in dark mode (≥ 7:1 with `prefers-contrast: more`).
- **Performance budget**: WebGL pass < 5 ms per page on integrated
  graphics; CPU fallback < 80 ms per page on a Mac M1. Enforced by
  a perf trace in CI.
