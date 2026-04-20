# PDF — Annotation Model

> Typed annotation enum, AP-stream emit rules, XFDF JSON I/O. Every
> annotation we write is a native PDF annotation visible in Adobe
> Acrobat / Preview / Chrome.

Cross-references: typed annotation projection in
[`document-model.md`](./document-model.md);
form-field annotations in [`form-engine.md`](./form-engine.md);
agent commands in [`agent-commands.md`](./agent-commands.md).

## Typed `PdfAnnotation` enum

Defined in [`packages/pdf/src/model/types.ts`](../../packages/pdf/src/model/types.ts):

```typescript
export type PdfAnnotationKind =
  | "highlight" | "underline" | "strikethrough" | "squiggly"   // text markup
  | "note" | "free-text"                                        // text
  | "ink"                                                       // free-hand
  | "line" | "rectangle" | "ellipse" | "polygon" | "polyline"   // shapes
  | "stamp" | "link" | "redaction"                              // misc
  | "unknown";                                                  // round-trip preserved

export interface PdfAnnotation {
  readonly id: NodeId;
  readonly kind: PdfAnnotationKind;
  readonly subtype: string;             // native PDF /Subtype string
  readonly pageNumber: number;
  readonly rect: PdfRect;
  readonly contents?: string;
  readonly author?: string;
  readonly color?: { r: number; g: number; b: number; a?: number };
  readonly url?: string;                // for link
  readonly destPage?: number;           // for goto-link
  readonly createdAt?: string;          // ISO-8601
  readonly nativeObjectNumber?: number; // for incremental save
}
```

| `kind`         | `/Subtype`     | Visual                                   |
| -------------- | -------------- | ---------------------------------------- |
| `highlight`    | `Highlight`    | Translucent fill over `QuadPoints`       |
| `underline`    | `Underline`    | Solid line at the descender baseline     |
| `strikethrough`| `StrikeOut`    | Solid line at mid-height                 |
| `squiggly`     | `Squiggly`     | Wavy line at the descender baseline      |
| `note`         | `Text`         | Sticky-note icon with popup on click     |
| `free-text`    | `FreeText`     | Floating text box                        |
| `ink`          | `Ink`          | Vector strokes from `InkList`            |
| `line`         | `Line`         | Line segment with optional arrowheads    |
| `rectangle`    | `Square`       | Stroked / filled rectangle               |
| `ellipse`      | `Circle`       | Stroked / filled ellipse                 |
| `polygon`      | `Polygon`      | Closed polygon                           |
| `polyline`     | `PolyLine`     | Open polyline                            |
| `stamp`        | `Stamp`        | Image stamp (preset or custom)           |
| `link`         | `Link`         | Invisible click region with action       |
| `redaction`    | `Redact`       | Black box + content-stream removal       |
| `unknown`      | (preserved)    | Round-trip the original /AP stream verbatim |

`unknown` is the catch-all for native annotation subtypes we don't
expose in the typed model (e.g. `RichMedia`, `Movie`, `Sound`, `3D`).
They are preserved through incremental save so they survive a
round-trip, but the editor cannot manipulate them.

## AP-stream emit rules

Every annotation we **write** carries an explicit `/AP` (appearance)
stream. Per the PDF spec, viewers render `/AP` when present and fall
back to drawing from the typed fields otherwise. Always emitting
`/AP` ensures consistent rendering across Acrobat, Preview, Chrome,
and any non-spec-strict viewer.

### Per-kind emit recipes

**Text markup (`Highlight` / `Underline` / `StrikeOut` / `Squiggly`)**

- `/QuadPoints` derived from the user's text selection; eight numbers
  per quad in PDF user-space.
- `/AP` content stream:
  - `Highlight`: `/Multiply` blend mode, fill rect for each quad with
    `{r,g,b}` color (default `1.0 1.0 0.0` yellow).
  - `Underline`: stroke a line at the descender for each quad.
  - `StrikeOut`: stroke a line at mid-quad-height.
  - `Squiggly`: stroke a wave (zig-zag with 2pt amplitude) at the
    descender.

**Sticky note (`Text`)**

- `/Icon` `Comment` (24×24 visual icon).
- `/AP` is the standard sticky-note glyph drawn at `/Rect`.
- `/Contents` carries the comment text.
- `/Open` is `false` by default.
- `/Popup` references the popup annotation for the threaded reply
  display.

**Free text (`FreeText`)**

- `/DA` (default appearance) sets font + size + color.
- `/Q` quadding (0=left, 1=center, 2=right).
- `/AP` content stream renders the text via `BT … ET` with the chosen
  font (one of the 14 standard fonts unless a custom font is
  embedded via fontkit).

**Ink (`Ink`)**

- `/InkList` is an array of stroke arrays, each a list of
  `[x1, y1, x2, y2, …]` user-space points captured from
  pointermove during free-hand drawing.
- `/BS` (border style) carries the line width and dash pattern.
- `/AP` content stream traces the strokes using `m`/`l` operators
  with optional Catmull-Rom smoothing for natural curves.

**Shapes (`Line` / `Square` / `Circle` / `Polygon` / `PolyLine`)**

- `/Rect` is the bounding box.
- For `Line`, `/L` is `[x1, y1, x2, y2]`; `/LE` is the line-ending
  style (e.g. `[/None /OpenArrow]` for an arrow).
- For `Polygon` / `PolyLine`, `/Vertices` is the point list.
- `/IC` is the interior color (fill); `/C` is the border color.
- `/AP` content stream draws the shape with `re`/`m`/`l`/`c`
  operators + `S`/`f`/`B` paint operators.

**Stamp (`Stamp`)**

- `/Name` is the stamp name (`/Approved`, `/Draft`, …) for presets.
- For custom image stamps, the AP stream paints an `XObject`
  reference to the embedded image.
- The image is embedded as a Form XObject so transparency works.

**Link (`Link`)**

- `/Rect` is the click region.
- `/A` (action) carries either `/URI` (external) or `/D`
  (destination, for goto-page / goto-named-destination).
- `/Border` is `[0 0 0]` (invisible) by default.

**Redaction (`Redact`)**

Two-phase:

1. **Mark** — `/Subtype /Redact`, `/QuadPoints` covering the redacted
   text or `/Rect` for an arbitrary box, `/IC` (fill color, default
   black). The mark is visually a translucent box during review.
2. **Apply** — on `pdf:apply-redactions`, the marks are converted to:
   - opaque black fill rectangles in the page content stream;
   - removal of the underlying text from the content stream (so
     selecting under the box returns nothing);
   - scrubbing of any matching annotations and form-field values;
   - a redaction log emitted to a JSON sidecar.

Until applied, redaction marks are reversible. After applied, the
text is gone.

## XFDF JSON I/O for the agent

The agent's import/export format is **XFDF** (XML, Adobe-defined),
optionally wrapped in a JSON envelope for ease of construction by
LLMs.

### XFDF wire format

```xml
<?xml version="1.0" encoding="UTF-8"?>
<xfdf xmlns="http://ns.adobe.com/xfdf/" xml:space="preserve">
  <annots>
    <highlight page="0" rect="100,200,400,220"
               color="#ffff00" opacity="0.5"
               coords="100,220,400,220,100,200,400,200">
      <contents>Important clause.</contents>
    </highlight>
    <freetext page="2" rect="50,100,250,150" color="#000000"
              fontfamily="Helvetica" fontsize="12">
      <contents>Reviewer note.</contents>
    </freetext>
  </annots>
</xfdf>
```

### JSON envelope (agent-friendly)

```json
{
  "format": "xfdf-json/v1",
  "annotations": [
    {
      "kind": "highlight",
      "pageNumber": 1,
      "rect": [100, 200, 400, 220],
      "quadPoints": [100, 220, 400, 220, 100, 200, 400, 200],
      "color": { "r": 1.0, "g": 1.0, "b": 0.0, "a": 0.5 },
      "contents": "Important clause."
    },
    {
      "kind": "free-text",
      "pageNumber": 3,
      "rect": [50, 100, 250, 150],
      "color": { "r": 0, "g": 0, "b": 0 },
      "fontFamily": "Helvetica",
      "fontSize": 12,
      "contents": "Reviewer note."
    }
  ]
}
```

The agent CLI's `office-agent pdf import-annotations --format xfdf|json`
accepts both. `office-agent pdf export-annotations --format xfdf|json`
emits the chosen serialization.

## Coordinate convention

All annotation rects, quad-points, and ink lists are in **PDF
user-space**: origin at lower-left, units 1/72 inch by default. The
viewer denormalizes against the current page rotation when drawing.
For storage, normalized rects (0..1 in unrotated user-space) are
**only** used by `PdfComment.normalizedRect` — native annotations
keep absolute coordinates because that's what the spec requires.

## Author / timestamp / color metadata

- `/T` carries the author display name. The viewer fills this from
  the realtime presence identity. Server-side CLI fills it from
  `--author "Name"` or the OS username.
- `/M` carries the modification timestamp (PDF date format, e.g.
  `D:20260420143000+02'00'`). The viewer sets this on every write.
- `/C` is the stroke color (RGB array `[r g b]` in 0..1).
- `/IC` is the interior fill color. Optional opacity via `/CA`
  (stroke) and `/ca` (fill) in 0..1.

## Round-trip guarantees

- An annotation created in our viewer (typed + AP stream) opens in
  Adobe Acrobat, Preview, and Chrome with the correct visual.
- An annotation created in Adobe Acrobat (native AP stream) survives
  a load + no-op save in our viewer (incremental save preserves it
  byte-for-byte).
- An XFDF file exported from our viewer imports cleanly into Adobe
  Acrobat ("Comments → Import → Acrobat XFDF") and vice versa.
- An `unknown` annotation (e.g. `RichMedia`) round-trips unchanged.

## Real-time sync

Annotation creation / mutation / deletion flow through the
`CommandBus` and broadcast over `useCommandBroadcast`. Other connected
clients receive the command and apply it via the same handler the
local UI used. Live ink-stroke streaming (point-by-point during draw,
not just on stroke-end) is **deferred** to a polish follow-up; final
strokes do sync.
