# PPTX — Renderer

## Goal

Render PPTX snapshots in the browser with editable text and direct-manipulation gestures (drag to reposition, handles to resize). Render thumbnails of the same snapshots at 1/8 scale, non-interactive. **Every gesture goes through the command bus** — the renderer never mutates the model directly.

## Technology — hybrid SVG + HTML overlay

Per the user-confirmed decision in the plan:

- **SVG for shape geometry and pictures.**
  - One `<svg>` per slide, `viewBox="0 0 ${slideSize.cxEmu} ${slideSize.cyEmu}"`, `preserveAspectRatio="xMidYMid meet"`.
  - CSS `width: 100%; height: auto;` (with a fixed aspect-ratio container) handles pixel-space scaling. Coordinates inside the SVG are in EMU — drag deltas, resize handles, hit-testing all work in EMU without per-shape DPI math.
  - `<g>` per shape, `<rect>` / `<image>` / `<path>` / `<g>` (group) per shape kind.
  - Pictures use `<image href="${objectURL}"/>` with object-URLs minted on parse.
- **HTML overlay for editable text.**
  - For each `TextShape`, a `<div contenteditable>` is positioned **on top of** the SVG using absolute positioning and CSS `transform: translate(${x}px, ${y}px)` driven by `emuToPx(emu, dpi)` (DPI = 96 by convention).
  - Font size scales with the same DPI factor.
  - On `blur` / `Ctrl+Enter`, the overlay's text is diff'd against the model and emitted as one or more `pptx:set-text` (or `pptx:format-text` for partial selections) commands. While editing, the SVG layer's `<text>` element for that shape is hidden (`visibility: hidden`).
  - This avoids reimplementing text shaping (browser does it) while keeping the SVG canvas as the source of geometric truth.

## Module layout

```
packages/pptx/src/renderer/
  layout/                 # pure TS, headless-testable, no DOM imports
    units.ts              # emuToPx / pxToEmu / EMU_PER_INCH / DPI
    slide.ts              # slideViewBox(slide), slideAspectRatio(slide)
    shape.ts              # shapeBoundingBox(shape, masterChain?, layoutChain?)
    color.ts              # resolveColor(colorRef, theme): "RRGGBB"
    placeholder.ts        # resolvePlaceholderDefaults(shape, layout, master)
  svg/
    slide.ts              # slideToSvgString(slide, ctx) → string (SSR / thumbnail)
    shapes.ts             # text-shape-svg, picture-svg, group-svg (factories)
  react/                  # Web-only; React + DOM allowed here
    SlideCanvas.tsx       # interactive: SVG + HTML overlay, gestures → commands
    SlideThumbnail.tsx    # non-interactive miniature; calls slideToSvgString
    SlidesSidebar.tsx     # vertical list of thumbnails with selection state
    selection-handles.tsx # drag/resize affordances
    text-overlay.tsx      # contenteditable HTML layer over a TextShape
  index.ts                # only re-exports `layout/` and `svg/` (headless-safe)
  react/index.ts          # separate entry — re-exports React components
```

The package's `package.json` `exports` map exposes:

- `@officeai/pptx` — agent + model + layout/svg (headless-safe).
- `@officeai/pptx/agent` — same as DOCX, agent only.
- `@officeai/pptx/renderer` — `layout/` + `svg/` only (headless).
- `@officeai/pptx/renderer/react` — React components (DOM-allowed).

`scripts/check-architecture.mjs` is extended to assert the headless package boundary: `packages/pptx/src/renderer/layout/**`, `svg/**`, `agent/**`, `commands/**`, `model/**`, `parser/**`, `serializer/**` may not import from `react`, `react-dom`, `next`, `prosemirror-view`, or DOM globals.

## Layout module — pure functions

```typescript
export const EMU_PER_INCH = 914400;
export const DEFAULT_DPI = 96;
export const EMU_PER_PX_AT_96DPI = EMU_PER_INCH / DEFAULT_DPI; // 9525

export function emuToPx(emu: number, dpi = DEFAULT_DPI): number {
  return (emu * dpi) / EMU_PER_INCH;
}

export function pxToEmu(px: number, dpi = DEFAULT_DPI): number {
  return Math.round((px * EMU_PER_INCH) / dpi);
}

export function slideViewBox(slide: Slide, slideSize: SlideSize): string {
  return `0 0 ${slideSize.cxEmu} ${slideSize.cyEmu}`;
}

export function slideAspectRatio(slideSize: SlideSize): number {
  return slideSize.cxEmu / slideSize.cyEmu;
}

export function shapeBoundingBox(shape: Shape): { x: number; y: number; cx: number; cy: number } | null {
  if (shape.position && shape.size) {
    return { x: shape.position.xEmu, y: shape.position.yEmu, cx: shape.size.cxEmu, cy: shape.size.cyEmu };
  }
  return null;
}
```

Headless tests live in `packages/pptx/src/renderer/layout/*.test.ts` and assert the round-trip identity `pxToEmu(emuToPx(e)) === e` for typical EMU values.

## Color resolution

```typescript
type ColorRef =
  | { kind: "srgb"; hex: string }                    // <a:srgbClr val="…"/>
  | { kind: "sysClr"; hex: string }                  // <a:sysClr lastClr="…"/>
  | { kind: "scheme"; name: "accent1" | … | "tx1" }  // <a:schemeClr val="…"/>
  | { kind: "unsupported"; raw: OpaqueXml };

export function resolveColor(ref: ColorRef, theme: ThemeColorScheme): string {
  switch (ref.kind) {
    case "srgb":   return ref.hex;
    case "sysClr": return ref.hex;
    case "scheme": return theme[ref.name] ?? "000000";
    case "unsupported": return "000000"; // fallback
  }
}
```

`theme` is read from the slide's master's referenced theme part on demand. Shape-level color overrides (lumMod/tint/shade) are P1 — for P0 we render the base color and re-emit the original XML on save (no information loss).

## SVG factory

`slideToSvgString(slide, ctx)` returns a serialized SVG string, used for thumbnails AND as the SSR-safe rendering for the main canvas (the React component hydrates the same SVG and adds the overlay layer + gesture handlers).

```typescript
export interface SvgRenderCtx {
  readonly slideSize: SlideSize;
  readonly theme: ThemeColorScheme;
  readonly mediaUrls: ReadonlyMap<string, string>; // mediaPartPath → object URL or data URL
}

export function slideToSvgString(slide: Slide, ctx: SvgRenderCtx): string {
  const parts = [
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${slideViewBox(slide, ctx.slideSize)}" preserveAspectRatio="xMidYMid meet">`,
    `<rect width="100%" height="100%" fill="white"/>`,
    ...slide.shapes.map((s) => shapeToSvg(s, ctx)),
    `</svg>`,
  ];
  return parts.join("");
}
```

Per shape:

| Shape kind    | SVG output                                                                                                                                                                                                 |
| ------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `TextShape`   | `<g class="shape text" data-shape-id="…"><rect/><text>…</text></g>`                                                                                                                                        |
| `Picture`     | `<g class="shape pic" data-shape-id="…"><image .../></g>`                                                                                                                                                  |
| `GroupShape`  | `<g class="shape group" data-shape-id="…">…children…</g>`                                                                                                                                                  |
| `OpaqueShape` | `<g class="shape opaque" data-shape-id="…"><rect class="placeholder"/></g>` (the renderer cannot draw arbitrary SmartArt/charts; we draw a labeled placeholder so users know where the opaque shape lives) |

Text shapes get both:

- An SVG `<text>` element (rendered in screenshots / thumbnails / SSR).
- An HTML overlay `<div contenteditable>` (only in the React canvas; the SVG `<text>` is hidden while the overlay is mounted).

## React canvas — gesture → command translation

```typescript
type Gesture =
  | { kind: "drag";   shapeId: NodeId; deltaXemu: number; deltaYemu: number }
  | { kind: "resize"; shapeId: NodeId; deltaCxEmu: number; deltaCyEmu: number; anchor: ResizeAnchor }
  | { kind: "text-edit"; shapeId: NodeId; newText: string }
  | { kind: "text-format"; shapeId: NodeId; range: TextRange; format: TextFormatPayload };

function gestureToCommands(g: Gesture, snapshot: PptxSnapshot, slideIndex: number): Command[] {
  switch (g.kind) {
    case "drag":
      return [{ type: "pptx:set-position", payload: { slideIndex, shapeId: g.shapeId, x: ..., y: ... }, source: "human", ... }];
    case "resize":
      return [{ type: "pptx:set-size",     payload: { slideIndex, shapeId: g.shapeId, width: ..., height: ... }, source: "human", ... }];
    case "text-edit":
      return [{ type: "pptx:set-text",     payload: { slideIndex, shapeId: g.shapeId, text: g.newText }, source: "human", ... }];
    case "text-format":
      return [{ type: "pptx:format-text",  payload: { slideIndex, shapeId: g.shapeId, range: g.range, format: g.format }, source: "human", ... }];
  }
}
```

`SlideCanvas.tsx` subscribes to `agent.subscribe()`, holds the current selection in local React state, listens for `pointerdown`/`pointermove`/`pointerup` on the SVG, and runs gesture-to-command on `pointerup` (drag/resize) or `blur`/`Enter` (text edit).

## Thumbnails

`SlideThumbnail.tsx` calls `slideToSvgString(slide, ctx)` and `dangerouslySetInnerHTML`s the result inside a fixed-aspect-ratio container at 200px wide. Non-interactive — clicks bubble up to the sidebar to change the active slide index. Re-renders only when the slide's `partHash` changes.

## Pending-mutation visualization

When `agent.getPendingMutations()` includes a mutation whose diff touches `shape:${shapeId}`, the canvas adds a `pending` CSS class to that shape's `<g>` (a violet outline, using the `--ai-violet` design token from `@officeai/ui`). Same convention DOCX uses for tracked-change-style staging.

## Out of scope for the renderer

- Rendering of charts (`graphicFrame`'s embedded chart parts).
- Rendering of SmartArt (`graphicFrame`'s `dgm:relIds`).
- Rendering of in-slide tables (`graphicFrame`'s `a:tbl`).
- Rendering of animations / transitions.
- Rendering of master/layout backgrounds beyond solid white. (P1 follow-up: consult the master's `<p:bg>` and render its solid/gradient/picture fill.)
- Connectors with non-trivial geometry (we draw a placeholder rect at the bounding box).
- Rich text shaping beyond what the browser does for HTML — no kerning, no advanced ligatures, no vertical text.

These are all **rendered as labeled placeholders** so users see where the unsupported content lives without breaking the canvas, and the underlying XML round-trips byte-clean on save.
