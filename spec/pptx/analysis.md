# PPTX — Reference Analysis (Step A)

> Notes derived from public READMEs, architecture docs, the OOXML
> specification (ECMA-376 5th-edition Transitional, parts dealing with
> PresentationML and DrawingML), and the `officeopenxml.com` reference.
> **No code from the references is copied.** This file informs our spec;
> it is not the spec.

References surveyed (per [`prompt.md`](../../prompt.md) §Reference Repositories):

- `gitbrent/PptxGenJS` (MIT) — generation/serialization helper. Builder API; no parser.
- `pipipi-pikachu/PPTist` (AGPL) — **architecture concepts only**. Read public READMEs and the high-level data-model write-ups; **no source code consulted**, no patterns inherited that aren't independently derivable from the spec. Treated as a "this kind of editor exists and ships" existence proof, not a reference implementation.
- ECMA-376 / `officeopenxml.com` — canonical truth, prefer over any implementation.

The DOCX clean-room work in [`spec/docx/analysis.md`](../docx/analysis.md) is the most important reference for *us*: PPTX reuses `@officeai/core` (`OoxmlContainer`, `RelationshipGraph`, `ContentTypes`, `CommandBus`, `parseXml`/`serializeXml`, `sha256Hex`) and the same architectural invariants (headless first, byte-preservation by part-hash, pending/approved/working tri-state).

---

## 1. In-memory document model

PresentationML structures a deck as:

```
Presentation (ppt/presentation.xml)
 ├─ sldIdLst → ordered list of slide rIds → ppt/slides/slide{N}.xml
 ├─ sldMasterIdLst → masters
 ├─ notesMasterIdLst → notes masters
 ├─ sldSz (width × height in EMU)
 └─ defaultTextStyle, custShowLst, etc. (preserved verbatim)

Slide (ppt/slides/slide{N}.xml)
 ├─ cSld
 │   └─ spTree (shape tree — the canvas content)
 │       ├─ sp (shape: rect, ellipse, …, with optional txBody)
 │       ├─ pic (picture — bound by blipFill to a media rel)
 │       ├─ grpSp (group of shapes — recursive)
 │       ├─ cxnSp (connector)
 │       ├─ graphicFrame (table / chart / SmartArt — opaque to us)
 │       └─ contentPart, AlternateContent, … (opaque)
 ├─ clrMapOvr (color map override; usually empty)
 ├─ transition (preserved verbatim)
 └─ timing (animations; preserved verbatim)

SlideLayout (ppt/slideLayouts/slideLayout{N}.xml)
 └─ same shape-tree structure, but with placeholders (ph) defining defaults.

SlideMaster (ppt/slideMasters/slideMaster{N}.xml)
 └─ shape tree + clrMap + theme reference.

Theme (ppt/theme/theme{N}.xml)
 └─ color scheme, font scheme, format scheme.

NotesSlide (ppt/notesSlides/notesSlide{N}.xml)  — optional, per-slide
NotesMaster (ppt/notesMasters/notesMaster1.xml) — optional
TableStyles (ppt/tableStyles.xml)               — preserved verbatim
ViewProps  (ppt/viewProps.xml)                  — preserved verbatim
PresProps  (ppt/presProps.xml)                  — preserved verbatim
```

PptxGenJS, being a builder, has no input-side model: it constructs slides/shapes from a high-level JS API and emits OOXML. Useful for:

- The exact shape of `<p:sp>`, `<p:pic>`, `<p:txBody>`, `<a:p>`, `<a:r>`, `<a:rPr>`, `<a:solidFill>`, `<a:blipFill>`, `<a:xfrm>` it emits — confirms what Word/PowerPoint accepts in practice.
- Default attribute values (e.g. `<p:nvSpPr>` boilerplate; the `<a:ext>` / `<a:off>` ordering inside `<a:xfrm>`).

PPTist (architecture-level only) confirms: editors keep an in-memory model at slide-tree granularity and re-render on mutation. Our independent design lands at the same structural shape because it falls out of OOXML directly.

### Decisions for our model

- **Typed slice for what we edit, opaque preservation for everything else.**
  - **Typed:** `Slide`, `Shape` (the union below), `TextBody { paragraphs[] → runs[] }` (paragraphs and runs only — no level, lang, smartTag promotion this round), `Position { xEmu, yEmu }`, `Size { cxEmu, cyEmu }`, `Picture` with media-rel reference.
  - **Opaque:** masters, layouts, theme, notes slides, transitions, timing, table cells, chart graphicFrames, SmartArt graphicFrames, connectors with non-trivial geometry, group shapes whose children we don't need to address (we still parse the group's bounding box so the renderer can show them; children stay opaque).
- **Shape union:**
  ```
  Shape =
    | TextShape   { kind: "text",   id, name, position, size, txBody, opaqueXfrmExtras?, raw? }
    | Picture     { kind: "pic",    id, name, position, size, mediaRelId, raw? }
    | GroupShape  { kind: "group",  id, name, position, size, children: Shape[], raw? }
    | OpaqueShape { kind: "opaque", id, name, position?, size?, rawXml }
  ```
  - `id` is the OOXML `<p:nvSpPr><p:cNvPr id="…">` int (PPTX uses ints, not GUIDs). The parser also mints a stable `NodeId` and stores it separately so commands address shapes regardless of `cNvPr` collisions across slides.
- The model is **immutable** at the surface (mutations produce new snapshots). Internally a handler may do a shallow copy + targeted splice, then `Object.freeze` in dev/test via `freezeSnapshot`.
- The model is **format-aware but renderer-agnostic.** SVG/HTML projection happens at render time. The model is the truth.

---

## 2. Parsing strategy

OOXML files are zip archives. Inside, `ppt/presentation.xml` is the manifest; the per-slide XML lives in `ppt/slides/slideN.xml`; each slide's relationships graph lives in `ppt/slides/_rels/slideN.xml.rels` and binds picture references to the `ppt/media/imageN.ext` parts.

### Decisions for our parser

- Reuse `OoxmlContainer` and `parseXml(..., { preserveOrder: true })` from `@officeai/core` exactly as DOCX does. This gives us:
  - Namespace-faithful round-trip (`p:`, `a:`, `r:`, `mc:`, `wp:`, …).
  - Byte-cache for every part loaded from the zip (the foundation of the byte-preservation invariant).
- **Two-pass parse per snapshot:**
  1. **Pass 1 — manifest:** load `[Content_Types].xml`, `_rels/.rels`, `ppt/presentation.xml`, `ppt/_rels/presentation.xml.rels`. Build `slideOrder: string[]` (ordered list of slide part paths).
  2. **Pass 2 — slides:** for each slide part, load XML, walk `cSld/spTree`, recognize `sp`/`pic`/`grpSp`/`cxnSp`/`graphicFrame`. Anything else → `OpaqueShape`. Within a recognized `sp`, parse `nvSpPr` (id+name), `spPr` (xfrm → position/size + a stash for non-xfrm spPr children we don't introspect), and `txBody` (paragraphs → runs → text + per-run rPr).
- **Layouts/masters/theme/notes** → loaded as opaque `OpaquePart` entries in the snapshot. Parser only enumerates them and computes part-hashes (so the serializer knows they're untouched). It does **not** recurse into them.
- **Picture media** → for each `<p:pic>`, resolve `<a:blipFill>/<a:blip r:embed="rIdN">` against the slide's rels graph to obtain the media part path. Compute SHA-256 of the media bytes; record the digest on the `Picture` so `insert-image` can dedup against existing media without re-reading the bytes.
- **Default unknown handling:** any element inside `spTree` that doesn't match a recognized opener becomes an `OpaqueShape` carrying `rawXml` (a *serialized* slice of the original parsed-order array, NOT a re-stringified version). On serialize the slice is re-emitted in place.

---

## 3. Serialization & untouched-parts preservation

Same invariant as DOCX (the only one that cannot be traded away): a part the editor did not touch must come back **byte-identical** on export.

PptxGenJS regenerates everything from scratch — useful as a fixture-generation helper, useless as a roundtrip reference. PPTist (per public docs) does not preserve untouched parts as a hard invariant.

### Decisions for our serializer

- Driven by `dirty: { presentation, slides: Set<partPath>, masters, layouts, theme, notesSlides: Set<partPath>, media: Set<partPath>, relationships: Set<partPath>, contentTypes }`.
- Per-part decision tree:
  - If the part is **not in `dirty`** → emit from `OoxmlContainer.partsCache` byte-identical.
  - If a slide is dirty → re-serialize from the typed `Slide` model. Untouched shapes inside a dirty slide use their `raw` slice (when present) so attribute order is preserved. Touched shapes are emitted from typed fields.
  - If `relationships` is dirty for any rels part → re-emit just that rels file.
  - If `contentTypes` is dirty → re-emit `[Content_Types].xml`.
  - If `media` is dirty (image insert/dedup) → write new media part bytes; old media bytes are kept verbatim from the cache.
- The serializer **never re-orders** sibling elements unless a typed mutation explicitly does so (e.g. `move-slide` reorders `<p:sldIdLst>` children). All other reorderings are bugs.

### Slide ID accounting (the load-bearing invariant)

`<p:sldIdLst><p:sldId id="N" r:id="rIdM"/></p:sldIdLst>` — `id` is the **stable PowerPoint slide id** (not reused across deletes; PowerPoint mints new ones starting at 256, increments). `rIdM` resolves through `ppt/_rels/presentation.xml.rels` to a `slideN.xml` part path. Our model:

- Stores slides keyed by part path (`ppt/slides/slide12.xml`), but exposes **slide index** to commands (the user-facing ordinal).
- On `add-slide`: pick the next free slide-id (`max(existing) + 1`, lower bound 256), pick the next free relationship id (`rId{max+1}`), pick the next free part path (`slide{N}.xml` where `N` is `max(existing) + 1` across the whole part directory).
- On `delete-slide`: drop the part, its rels file, its `<p:sldId>` entry, its content-type override, and any attached notes-slide part (resolved through the slide's own rels graph, type `…/notesSlide`). **Do not** renumber surviving slide-ids or part paths — that would invalidate every reference to them in the deck.
- On `move-slide`: reorder the children of `<p:sldIdLst>` only. Part paths and `rId`s never change.
- On `duplicate-slide`: deep-clone the typed slide tree, mint fresh `cNvPr` ids inside it (`max(existing on the new slide) + 1`, scoped per slide so cross-slide collisions are fine), copy the source slide's `_rels` file (so picture references to the same media parts are preserved without copying media bytes), and append a fresh `<p:sldId>` + `<p:Override>`.

---

## 4. Mutation / command pattern

PptxGenJS has no mutation pattern (builder only). PPTist's UI emits internal events; our independent design lands at the standard command-bus pattern that DOCX already uses.

### Decisions for our command bus

- The pattern from [`prompt.md`](../../prompt.md) lines 312–332 is non-negotiable. Every edit — pointer drag in the SVG canvas, contenteditable keystroke in the HTML text overlay, agent call, CLI invocation — produces a `Command<T,P>`, dispatches through `CommandBus<PptxSnapshot>`, yields a `Mutation` with a structured `DocumentDiff`.
- Re-use `CommandBus` from `@officeai/core` exactly as DOCX does. No format-specific bus.
- The hybrid renderer's React components must intercept all gestures and translate them into commands at the **gesture boundary** (`pointerup` for drag, `blur` or `Enter` for text edit, button clicks for toolbar). They must never mutate model state directly.
- Pending agent mutations live in the same approved/pending/working tri-state slice the bus already implements. Visual decoration of pending shapes happens in the renderer (a violet outline using the same `aiViolet` design token DOCX uses).

---

## 5. The hard parts

### EMU ↔ px conversion

OOXML uses **English Metric Units** for spatial coordinates: `914400 EMU = 1 inch`, `9525 EMU = 1 px @ 96 DPI`. Slide size is typically `9144000 × 6858000 EMU` (10" × 7.5" classic 4:3) or `12192000 × 6858000` (16:9). We:

- Store everything in EMU in the model.
- `emuToPx(emu, dpi=96)` and `pxToEmu(px, dpi=96)` live in `packages/pptx/src/renderer/layout/units.ts` as pure functions; both headless-tested.
- The SVG `viewBox` is set to the slide's EMU dimensions; CSS scales it to the available pixel box. This means selection coordinates, drag deltas, and resize handles all work in EMU space without per-shape math.

### Theme color resolution

Shapes reference colors as one of: `<a:srgbClr val="FF0000"/>` (literal), `<a:schemeClr val="accent1"/>` (theme reference), `<a:sysClr val="windowText" lastClr="000000"/>` (system color with literal fallback), or via percent-tints/shades (`<a:lumMod val="75000"/>`, `<a:tint val="40000"/>`).

- For P0 we resolve `srgbClr` and `sysClr` (using `lastClr`) directly. `schemeClr` is resolved by reading the slide's master's theme color scheme (`<a:clrScheme>`) — but only at render time; the model stores the unresolved reference so serialization is exact.
- Tints/shades are passed through to the SVG renderer as `filter`/derived values; for P0 we approximate (display only) and round-trip the original XML on save.

### Placeholder inheritance

A placeholder `<p:sp>` with `<p:nvSpPr><p:nvPr><p:ph type="title" idx="0"/></p:nvPr></p:nvSpPr>` inherits position/size/text-style from the matching placeholder in the slide's layout, which in turn inherits from the master. Resolving the full effective shape requires walking the chain (slide → layout → master).

- For P0, the **renderer** does the chain walk on demand for visual fidelity. The **model** stores only what the slide's own XML declares (which is what gets serialized back). This avoids ever silently materializing layout/master values into the slide's bytes.

### Notes slides, transitions, animations, timing

All preserved as opaque XML/parts. We never touch them. `delete-slide` correctly drops the attached notes slide via the slide's rels graph (type `…/notesSlide`).

### Picture media dedup

Same problem DOCX solves for `insert-image`: when the same image bytes are inserted twice, we re-use the existing media part instead of writing a duplicate.

- SHA-256 of the bytes is the dedup key (computed by `sha256Hex` from `@officeai/core`).
- The dedup index is built lazily on first `insert-image` call by walking `presentation.media` (populated by the parser).

### Group shapes (`p:grpSp`)

Group shapes have their own `xfrm` describing the group's position/size and `chOff`/`chExt` describing the child coordinate system offset/extent. For P0:

- Parse the group's `xfrm` so the renderer can place it.
- Treat the group's children as opaque (rendered via raw-XML projection or a thumbnail-style flattening). Editing children of a group is deferred to a follow-up.
- Moving a group shape (`set-position`) updates only the group's `xfrm`, not the children's coordinates (PowerPoint's behavior).

---

## 6. What references get wrong / sacrifice that we improve

- **PptxGenJS cannot read existing files.** It is a one-way builder. We build the parser it lacks.
- **PPTist (architecture inference only)** does not, by public documentation, treat byte-preservation as a hard invariant; its primary mode is to *render* PPTX into its own model and re-serialize on save. We make byte-preservation the central invariant, identical to DOCX.
- **None of the references treat agents as first-class.** Same gap as DOCX: the agent API and CLI are primary; the UI is a skin over the same headless agent.
- **None expose a structured, per-mutation diff** an agent can introspect / approve / reject. We inherit DOCX's approach.

---

## 7. What's missing from the 80% scope we need

Beyond what the references cover, our 80% requires:

- **Headless-first I/O** — load + edit + export with zero DOM (Node-only path, exercised by the agent test that mounts no React).
- **CLI** — `office-agent pptx inspect | read | set-text | add-slide | …`.
- **Pending-mutation staging** — the approved/pending/working tri-state from [`prompt.md`](../../prompt.md) lines 437–451. Already in `CommandBus`.
- **Opaque-blob preservation as a hard invariant** — verified by SHA-256 over every untouched part on every roundtrip test.
- **Hybrid SVG/HTML renderer with EMU-correct coordinates** — neither reference does this. Falls out of our headless layout module + a thin React adapter.
- **Per-feature confidence ratings** in our spec (P0/P1/P2/OUT) — keeps deferrals explicit, identical to DOCX.

---

## Summary

Our PPTX implementation borrows the **slide-tree-of-shapes / typed-properties** structural intuition that falls directly out of OOXML and is independently confirmed by the public-API surface of PptxGenJS and the architecture-level READMEs of PPTist. Everything else — the command bus, the headless agent, the staging tri-state, the strict byte-preservation, the CLI, the agent-first ergonomics, the hybrid renderer with EMU-native coordinates — is original to this project and reuses the format-agnostic pieces of `@officeai/core` exactly as DOCX does.

Next: produce `spec/pptx/*` (feature-scope, document-model, ooxml-mapping, parser, serializer, renderer, agent-commands, edge-cases, acceptance-criteria).
