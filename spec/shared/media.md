# Shared — Media (video / audio) + image polish

> Status: F1. Promotes embedded video and audio from preserve-only
> to first-class with insert + Present-mode playback. Closes the
> deferred image-UX backlog (crop mode, resize-handle
> disambiguation, alt text). Affects DOCX and PPTX; XLSX still
> only embeds images.

## Why

`prompt.md` says "Embedded videos / audio: open + roundtrip, not
edit". The user's TODO list overrides that for the SME-suite
positioning: business decks rely on video and audio constantly.
Today we round-trip `.mp4` / `.mp3` cleanly as opaque media but
nobody can insert one or play it in Present mode.

Image UX deferred items from the prior night also land here:
real crop mode, resize-handle disambiguation across all three
editors, alt text persisted.

## Video / audio model

### PPTX

```ts
// packages/pptx/src/model/types.ts (additions)

export interface MediaShape {
  readonly kind: "media";
  readonly id: string;
  readonly name: string;
  readonly position: { x: number; y: number; cx: number; cy: number };
  readonly mediaType: "video" | "audio";
  /** Path inside the OOXML container, e.g. "ppt/media/video1.mp4". */
  readonly mediaPath: string;
  /** Optional poster image (also inside `ppt/media/`). */
  readonly posterPath?: string;
  /** When playback starts: "click" | "automatic" | "inClickSequence". */
  readonly trigger: "click" | "automatic" | "inClickSequence";
  /** Loop until end of slide. */
  readonly loop: boolean;
  /** Mute audio (audio shapes too — to mute their playback). */
  readonly muted: boolean;
  /** Show controls (PowerPoint default true for video). */
  readonly showControls: boolean;
  /** Verbatim attributes we don't need to change. */
  readonly raw?: Readonly<Record<string, unknown>>;
}
```

OOXML mapping: a `<p:pic>` with a `<p:nvPicPr>` containing a
`<p:nvPr>` that has an `<a:videoFile r:link="..."/>` or
`<a:audioFile>`. The `r:link` resolves through the slide's rels
to a `ppt/media/...` part. We add `MediaShape` to the slide
shape sum type.

### DOCX

```ts
// packages/docx/src/model/types.ts (additions)

export interface MediaInline {
  readonly kind: "media";
  readonly id: NodeId;
  readonly mediaType: "video" | "audio";
  readonly mediaPath: string; // word/media/...
  readonly posterPath?: string;
  readonly cx: number; // EMU
  readonly cy: number;
  readonly muted: boolean;
}
```

OOXML: a `<w:drawing>` with a DrawingML `wp:inline` containing
`a:graphic > a:graphicData` with `a:videoFile` (Word's idiom,
shared with PowerPoint). We treat it as an inline media for now;
floating layout is P1.

### XLSX

Out of scope this session. Audio/video embedded in spreadsheets
stay opaque (current behaviour).

## Insert commands

| Command                  | Payload                                                                                   | Effect                                                                                                    |
| ------------------------ | ----------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------- |
| `pptx:insert-media`      | `{ slideId; bytes; mimeType; trigger?: "click"\|"automatic"; position?; loop?: boolean }` | Adds part to `ppt/media/`, dedup by SHA-256, mints `MediaShape`, optional poster derived from first frame |
| `docx:insert-media`      | `{ paragraphId; offset; bytes; mimeType; cx?; cy? }`                                      | Adds part to `word/media/`, mints `MediaInline` at caret                                                  |
| `pptx:set-media-trigger` | `{ slideId; shapeId; trigger }`                                                           | Mutates `MediaShape.trigger`                                                                              |
| `pptx:set-media-loop`    | `{ slideId; shapeId; loop }`                                                              | Mutates `MediaShape.loop`                                                                                 |
| `pptx:set-media-poster`  | `{ slideId; shapeId; bytes; mimeType }`                                                   | Inserts/replaces poster                                                                                   |

Allowed MIME types:

```ts
const VIDEO_MIME = new Set(["video/mp4", "video/webm", "video/quicktime"]);
const AUDIO_MIME = new Set(["audio/mpeg", "audio/mp4", "audio/wav", "audio/ogg"]);
```

Anything else: reject with `unsupported-media-type`.

## Renderer

### Edit mode

`MediaShape` renders as a foreignObject in the SVG slide canvas
hosting an HTML5 `<video controls=false poster=...>` (paused) or
an `<audio>` (with a small icon and waveform glyph). The shape
is selectable, draggable, resizable like any other shape — the
existing shape interaction layer handles that.

### Present mode

`apps/web/app/pptx-editor/MediaPlayer.tsx` is mounted per
`MediaShape` on the active slide:

- `trigger === "automatic"` → play on slide enter.
- `trigger === "click"` → play on next "advance" click in the
  click sequence.
- `loop` honoured.
- `showControls` honoured.

Audio shapes render as a small floating icon during edit; in
Present mode they're invisible by default (just the audio).

### DOCX

Inline media in DOCX renders as an HTML5 `<video>` / `<audio>`
inside a sized container in the body PM, marked `editable=false`
so PM treats it as an atom. Click to select; selection handles
allow resize.

## Image polish

### Crop mode

Toggled from the Picture Format ribbon tab (or via context menu
Crop). Implementation:

- `apps/web/app/lib/image-crop.ts` exposes a `useImageCropMode()`
  hook returning a state machine: `inactive` |
  `cropping { shapeId, srcRect: { l, t, r, b /* fractional 0..1 */ } }` |
  `committing`.
- The shape's image overlay shows the **uncropped** image at low
  opacity outside the crop rect, full opacity inside, with eight
  drag handles around the crop rect.
- Commit dispatches `pptx:set-image-srcRect` /
  `docx:set-image-srcRect` writing `<a:srcRect l="..." t="..." r="..." b="..."/>`
  into the image's `a:blipFill`. Values are fixed-point per OOXML
  (multiplied by 1000 and floored).

```ts
export interface SetImageSrcRectPayload {
  readonly l: number; // 0..1
  readonly t: number;
  readonly r: number;
  readonly b: number;
}
```

### Resize handles

Shared module `apps/web/app/lib/resize-handles.ts`:

```ts
export interface ResizeHandleSpec {
  readonly id: "nw" | "n" | "ne" | "e" | "se" | "s" | "sw" | "w" | "rotate";
  readonly cursor: string;
}

export function computeResize(
  start: { x: number; y: number; w: number; h: number },
  pointer: { dx: number; dy: number },
  handle: ResizeHandleSpec["id"],
  modifiers: { shift: boolean; alt: boolean }
): { x: number; y: number; w: number; h: number };
```

Rules:

- Corner handles scale both axes; `shift` constrains aspect
  ratio.
- Edge handles scale only one axis; `shift` is ignored.
- `alt` scales from the centre instead of from the opposite
  corner/edge.
- A separate **rotate** handle floats above the top-centre at a
  fixed offset; click-drag rotates around the shape centre.

Adopted by all three editors' shape selection layers, replacing
the three independent code paths.

### Alt text

Property panel field for any image in PPTX (`<p:nvPicPr><p:cNvPr
descr="..."/>`) and DOCX (`<wp:docPr descr="..."/>`). Commands:

| Command                   | Payload                               | Effect                     |
| ------------------------- | ------------------------------------- | -------------------------- |
| `pptx:set-shape-alt-text` | `{ slideId; shapeId; descr; title? }` | Sets alt text on any shape |
| `docx:set-image-alt-text` | `{ imageId; descr; title? }`          | Sets alt text on image     |

## Round-trip invariants

1. **No-edit MP4.** A deck with an embedded MP4 round-trips
   byte-clean.
2. **Single-image crop.** Cropping one image only re-emits that
   slide's `slideN.xml`; media bytes unchanged.
3. **Alt text.** Adding alt text writes only the
   `cNvPr@descr`/`docPr@descr` attribute; rest of the OOXML is
   byte-equal.
4. **Insert MP4.** Inserting a 5MB MP4 grows the container by
   exactly the file size + the typical XML overhead (one
   `<p:pic>` + one rels entry + one ContentType override).

## Acceptance criteria

A1. **PPTX video insert + play.** Insert a 5-second MP4, switch
to Present, the video plays.

A2. **DOCX audio insert.** Insert an MP3, see an audio player in
the body, save and reopen — audio is preserved.

A3. **Crop in PPTX.** Crop an image via the new crop mode, save
and reopen in PowerPoint — crop matches.

A4. **Shift-resize.** Drag a corner handle with Shift held —
aspect ratio preserved.

A5. **Alt text.** Set alt text on a picture, save and reopen —
the value survives.

A6. **No regressions.** Existing image / shape resize Playwright
tests stay green.

## Out of scope (F1)

- Video trimming / cue points.
- Linked (not embedded) media.
- Closed captions.
- Floating images in DOCX with text wrap (use inline anchors).
- 3D models (`a:graphicFrame` with `m3d`) — preserve only.
- Animated GIFs as media (already supported as images).
