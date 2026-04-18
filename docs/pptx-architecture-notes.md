# PPTX Architecture Notes — Where Slides Diverge from DOCX

> Companion to [`docs/session-summary-pptx.md`](./session-summary-pptx.md) and [`docs/build-log/pptx.md`](./build-log/pptx.md).
> Captures the patterns the slides phase uses **differently from `@officeai/docx`**, and the reasons behind each one. None of these are accidents — they all exist because a slide is a 2-D, absolutely-positioned, multi-part shape canvas, while a Word document is a single text-flow part.
>
> Future maintainers: if a divergence here ever bites you, this is the doc to start with. Inverse-direction question — "should DOCX adopt this pattern?" — is occasionally yes (see callouts).

## TL;DR — slide-specific patterns at a glance

| # | Divergence                                                      | Where it lives                                                                                                                  | Why slides need it                                                                                                                              |
| - | --------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------- |
| 1 | Snapshot owns typed OPC scaffolding (`relationships`, `contentTypes`, `removedParts`) | `packages/pptx/src/model/types.ts` — `PptxSnapshot`                                                                             | Every command can add or delete whole OOXML parts (slide N + its `_rels` + its notes). DOCX never deletes its sole `word/document.xml`.         |
| 2 | `PptxDirty` is mostly `ReadonlySet<string>`, not booleans       | `packages/pptx/src/model/types.ts` — `PptxDirty`                                                                                | Every category is 1:N (N slides × M layouts × K themes × L charts). Set granularity lets us rebuild slide 7 while leaving slides 1–25 byte-identical. |
| 3 | Three-tier opaque taxonomy: `OpaqueXml` + `OpaqueShape` + `OpaquePart` | `packages/pptx/src/model/types.ts`                                                                                              | DOCX has no whole-part opaque bucket; PPTX leaves whole `master`, `layout`, `theme`, `notesSlide` parts untouched.                              |
| 4 | Slide-level opaque buckets (`cSldHead`, `spTreeHead`, `slideOpaqueTail`, `timingTailRaw`) | `packages/pptx/src/model/types.ts` — `Slide`                                                                                    | `<p:sld>` is order-sensitive (cSld → clrMapOvr → transition → timing); per-bucket capture preserves byte order without modelling every child.    |
| 5 | Persistent monotonic counter (`PptxIdGen`) on the snapshot root | `packages/pptx/src/model/types.ts` — `PptxIdGen`                                                                                | PowerPoint requires `slideId ≥ 256` and increasing across the deck; part paths must not reuse freed integers within a session.                  |
| 6 | `cNvPrId` minted per-slide (not per-document)                   | `packages/pptx/src/commands/helpers.ts` — `maxCNvPrId`                                                                          | Animations target shapes via `<p:spTgt @spid>`; that reference resolves only inside the same slide part.                                        |
| 7 | Renderer is a direct command-dispatch hybrid (no editor library) | `packages/pptx/src/renderer/react/SlideCanvas.tsx`                                                                              | A 2-D, mostly-non-text canvas would fight ProseMirror's flat-text-with-marks model. We dispatch `pptx:set-position/-size/-text` from native React events. |
| 8 | `headless-invariant.test.ts` enforces the renderer split with both an import-allowlist *and* a DOM-globals regex | `packages/pptx/src/headless-invariant.test.ts`                                                                                  | Renderer is split into three sibling dirs (`layout/`, `svg/`, `react/`); only `react/` may touch React or `document`/`window`. DOCX has no equivalent guard. |
| 9 | `SvgRenderCtx` plumbed through every shape function (`theme`, `mediaUrls`, `charts`) | `packages/pptx/src/renderer/svg/shapes.ts` — `SvgRenderCtx`                                                                     | Theme colours are resolved *only at paint time* so the OOXML `<a:solidFill><a:schemeClr>` capture stays opaque; chart parts live in separate XML files and need a lookup table at render. |
| 10 | Dedicated `dirty-roundtrip.test.ts` invariant suite              | `packages/pptx/src/serializer/dirty-roundtrip.test.ts`                                                                          | Asserts the negative — "marking one slide dirty changes nothing else". DOCX gets this for free with one main part.                              |
| 11 | Real-world roundtrip threshold is **≥ 95 %**, not 100 %          | `tests/roundtrip/pptx/real-world-roundtrip.test.ts`                                                                             | `ppt/presentation.xml` is an *index* of slides; any insert/delete/reorder forces it to be re-emitted, which can drift on attribute order.       |
| 12 | Cross-CLI-invocation tests rely on `OFFICEAI_DETERMINISTIC_IDS=1` | `packages/agent/src/pptx-cli.test.ts`                                                                                           | Shapes have **no OOXML-stable identifier** (`cNvPrId` is non-unique across slides and unstable across edits). The deterministic minter is the test-time workaround. |

## 1. Model shape

### 1.1 Typed OPC scaffolding lives on the snapshot

```5:32:packages/pptx/src/model/types.ts
import type { ReadonlyOoxmlContainer } from "@officeai/core/ooxml/container";
import type { OpaqueXml } from "@officeai/core/ooxml/opaque";

export type NodeId = string;

export interface PptxSnapshot {
  readonly version: number;
  readonly root: PptxPresentation;
  readonly container: ReadonlyOoxmlContainer;
  readonly relationships: ReadonlyMap<string, RelationshipsSnap>;
  readonly contentTypes: ContentTypesSnap;
  readonly removedParts: ReadonlySet<string>;
  readonly dirty: PptxDirty;
}
```

Compare to DOCX, which keeps only `dirty` + `container` and lets the cloned container hold rels + content-types. The reason is `removedParts`: `delete-slide` and `insert-image-replacement` both need the serializer to drop parts from the cloned container. There is no analogue in `word/`.

**Inverse direction**: DOCX does not currently need this; if the team ever ships `delete-image` or `delete-comment` (the "comment-extended part disappears when the last comment is resolved" case), the same pattern would help.

### 1.2 `PptxDirty` is a Set-of-parts

```34:46:packages/pptx/src/model/types.ts
export interface PptxDirty {
  presentation: boolean;
  slides: ReadonlySet<string>;        // partPaths
  notesSlides: ReadonlySet<string>;
  masters: ReadonlySet<string>;
  layouts: ReadonlySet<string>;
  theme: ReadonlySet<string>;
  media: ReadonlySet<string>;
  charts: ReadonlySet<string>;
  relationships: ReadonlySet<string>; // owning partPath
  contentTypes: boolean;
}
```

`DocxDirtyFlags` is mostly singleton booleans because Word has one body, at most one comments part, etc. The Set-shaped model is what makes the byte-identity invariant tractable for a 25-slide deck — and what makes the dedicated `dirty-roundtrip.test.ts` suite (#10 below) possible.

### 1.3 Three-tier opaque taxonomy

|              | `OpaqueXml` carrier                                | `OpaqueShape` typed kind             | `OpaquePart` whole-part wrapper        |
| ------------ | -------------------------------------------------- | ------------------------------------ | -------------------------------------- |
| **Lives on** | inline (e.g. `Slide.slideOpaqueTail`, `TextRun.opaqueChildren`, `ChartPart.chartSpaceRaw`) | `Shape` discriminated union (SmartArt, unmodelled `graphicFrame` payloads) | `PptxPresentation.{masters, layouts, theme, notesSlides}` |
| **Survives**  | If the **enclosing slide is not dirty** → re-emitted byte-for-byte from container cache. Otherwise re-serialized via `serializeXmlTree`. | A slide-rebuild (it owns its own `raw: OpaqueXml`).                              | Always — these parts are never re-emitted by the slides phase. |
| **DOCX equivalent** | Yes (`OpaqueBlock`/`OpaqueInline`/`OpaqueRunChild`/`OpaqueDrawing`).                       | Partial (`OpaqueBlock` covers SmartArt-equivalent fallbacks).                    | **None** — DOCX has no whole sibling parts we deliberately ignore. |

`Slide.timingTailRaw` is a fourth hybrid: re-emitted verbatim if `Slide.animations` is untouched; *dropped and rebuilt* via `timingFromAnimations` if any animation command runs. There is no DOCX analogue — `Table.raw` and `InlineImageDrawing.raw` are simpler "bytes if present, else regenerate" markers without the model-vs-raw flip.

See `docs/build-log/pptx.md` known-issues for the loss-of-unmodelled-children caveat.

### 1.4 Slide buckets

```182:218:packages/pptx/src/model/types.ts
export interface Slide {
  readonly id: NodeId;
  readonly slideId: number;
  readonly partPath: string;
  readonly layoutRel?: string;
  readonly cSldHead: ReadonlyArray<OpaqueXml>;     // children of <p:cSld> before <p:spTree>
  readonly spTreeHead: ReadonlyArray<OpaqueXml>;   // <p:nvGrpSpPr> + <p:grpSpPr>
  readonly shapes: ReadonlyArray<Shape>;
  readonly slideOpaqueTail: ReadonlyArray<OpaqueXml>; // post-<p:cSld> opaque children
  readonly transition?: SlideTransition;
  readonly animations: ReadonlyArray<EntranceAnimation>;
  readonly timingTailRaw?: OpaqueXml;
  readonly notes?: NotesSlide;
}
```

DOCX has just `body: ReadonlyArray<BlockNode>` — flat. `<p:sld>` is order-sensitive in ways `<w:body>` isn't (`cSld → clrMapOvr → transition → timing`), and the per-bucket capture preserves byte order without forcing us to model every child kind.

### 1.5 Persistent ID minter

`PptxIdGen` (`nextSlideId`, `nextSlidePartIndex`, `nextMediaPartIndex`) is bumped by `add-slide` and `insert-image`. PowerPoint requires `slideId ≥ 256` and increasing across the deck even if slides are deleted; part paths likewise mustn't reuse a freed integer within a session. DOCX walks the body each time (`mintDocPrId`) because Word has no need to persist this state.

## 2. Identity and addressing

### 2.1 `cNvPrId` is slide-scoped

PPTX commands compute `maxCNvPrId(slide.shapes) + 1` per slide (`commands/helpers.ts`). DOCX `mintDocPrId` walks the entire body. The slide-scoping matters because `<p:spTgt @spid>` references in `<p:timing>` resolve **inside** the slide part — a document-wide minter would over-allocate and break compatibility with slides authored elsewhere.

### 2.2 `NodeId` UUIDs vs DOCX selector paths

PPTX commands take `(slideIndex, NodeId)` — a runtime-minted UUID that **isn't persisted to OOXML**. DOCX commands take selector strings like `paragraph:1/run:0/text:0..5`, parsed by `selector.ts` at every invocation, with comment IDs sourced from `w:comment[@w:id]`.

The reason slides need a different scheme: shapes have **no OOXML-stable public identifier**. `cNvPrId` is non-unique across slides and unstable across edits. The runtime `NodeId` is the only persistent option.

### 2.3 `OFFICEAI_DETERMINISTIC_IDS=1` for tests

Because `NodeId`s are UUIDs by default, chained CLI invocations (`a.pptx → b.pptx → c.pptx`) would each see a different `NodeId` for the same shape. The env var swaps the UUID minter for a deterministic counter so the *same shape parsed from disk twice produces the same `NodeId`*. Used heavily in `packages/agent/src/pptx-cli.test.ts` and the animation suite there. DOCX needs nothing equivalent because its selectors re-resolve from positional indexes.

## 3. Rendering

### 3.1 Hybrid renderer with **no editor library**

DOCX wraps ProseMirror: `schema.ts`, `doc-to-pm.ts`, `transaction-to-commands.ts`, `mount.ts` — a complete bidirectional bridge that translates PM transactions into commands.

PPTX has none of that. `SlideCanvas.tsx` dispatches commands directly:

- `pptx:set-shape-position` / `pptx:set-shape-size` from `onPointerUp` after a drag/resize.
- `pptx:set-text` from a `contenteditable` div's `onBlur`.

The shape currently being edited is **spliced out** of `svgInner` and replaced by an absolutely-positioned HTML overlay sized in slide-size percentages. This swap-out-and-overlay technique is unique to slides; PM would not have helped because slides are 2-D, absolutely positioned, and largely non-text.

### 3.2 `SvgRenderCtx` plumbed through every shape function

```19:26:packages/pptx/src/renderer/svg/shapes.ts
export interface SvgRenderCtx {
  readonly slideSize: SlideSize;
  readonly mediaUrls: ReadonlyMap<string, string>;
  readonly theme?: ThemeColorScheme;
  readonly charts: ReadonlyMap<string, ChartPart>;
}
```

The context is built once via `useMemo` keyed on `slideSize`, `mediaUrls`, `themeDefault`, `charts`, and threaded through every `*ToSvg` function. Both `SlideCanvas` (interactive) and `SlideThumbnail` (sidebar) reuse it. DOCX renders into PM nodes that carry their own attrs; there is no per-render context to thread because there is no theme resolution and no separate-part chart lookup at paint time.

### 3.3 `headless-invariant.test.ts` enforces the three-way split

`renderer/layout/` (pure layout fns) and `renderer/svg/` (pure SVG factory) must run in Node — in the agent, in tests, in the CLI. `renderer/react/` is the only directory allowed to import React or touch `document`/`window`. The invariant test walks `model/`, `parser/`, `serializer/`, `commands/`, `agent/`, `renderer/layout/`, `renderer/svg/` and rejects:

- imports of `react`, `react-dom`, `react/jsx-runtime`, `next`, `next/router`, `next/navigation`, `tailwindcss`,
- regex hits on `document.X`, `window.X`, `navigator.X`, `location.X` (after a string/comment-stripping pass to avoid matching the literal `"document.xml"`).

DOCX has no equivalent. It collapses everything into one `renderer/` because `prosemirror-view` only mounts in a browser anyway — the boundary is enforced by the dependency, not by a test. The PPTX guard is necessary precisely because `renderer/svg/` *could* trivially `import React` and nothing would catch it without this test.

### 3.4 EMU-as-SVG-coord-system

The SVG `viewBox` is `0 0 cxEmu cyEmu` (slide size in English Metric Units). All overlays — selection rectangles (`stroke-width="20000"`), animation badges (`r="90000"`), SE resize handles — are drawn directly in EMU, so they scale with the canvas without any per-zoom recalculation. Implicit in the F1.5 zoom story; called out here because it's a deliberate pattern future contributors should preserve.

### 3.5 Animation badges are SVG, not HTML

Numbered yellow circles drawn at the top-left of animated shapes use `pointer-events="none"` so they never steal clicks, and live inside the same `dangerouslySetInnerHTML` as shapes so they share the EMU coordinate system. An HTML overlay would have needed per-zoom positioning maths.

## 4. Roundtrip + testing

### 4.1 Dedicated dirty-roundtrip suite

`packages/pptx/src/serializer/dirty-roundtrip.test.ts` exercises the strict invariant: mark *one* slide dirty → reload → assert that **every other part** in `snapshot.container.parts.keys()` hashes identically (`sha256Hex`). The same file also asserts model-equivalence after a typed→raw→typed cycle on `TableShape`, `ChartShape`, and `Slide.transition + animations`.

DOCX has analogous individual cases scattered through `serialize.test.ts` and `commands/tables.test.ts` but no dedicated suite. With one main part, DOCX gets the "nothing else changed" assertion for free — anything that changes shows up in `word/document.xml`'s hash. The Set-of-parts dirty model in PPTX (#1.2 above) only makes sense if you have a test that proves the negative.

### 4.2 ≥ 95 % real-world roundtrip threshold

`tests/roundtrip/pptx/real-world-roundtrip.test.ts` accepts ≥ 95 % byte-identity on the pure-roundtrip case; `tests/roundtrip/docx/real-world-roundtrip.test.ts` asserts strict 100 %. The reason is specific to slides: `ppt/presentation.xml` is an *index* (it orders slides by `rId`), so any insert/delete/reorder forces it to be re-emitted, and the rebuild path can drift on attribute order even when no slide was deleted/added if a future code path flips `dirty.presentation`.

DOCX has no analogous "index" part — body order lives inside `word/document.xml` itself.

### 4.3 Multi-invocation CLI test pattern

`packages/agent/src/pptx-cli.test.ts` chains `runCli([...])` → `writeFileSync` → `loadDeterministic(out)` → re-resolve shape ID, repeating across `a.pptx → b.pptx → c.pptx → d.pptx`. The animation suite there exercises `add-shape-animation → reorder-shape-animations → remove-shape-animation` across four CLI processes and asserts the final on-disk model. This pattern requires the deterministic-IDs env var (#2.3) and is the single best safeguard against the "`NodeId`s shifted because we changed an unrelated part of the parser" class of bug — see the F4 `parseSlide` reorder fix in the build log for the example.

DOCX has multi-step tests but they run inside a single process and use selector-based addressing, so the failure mode this pattern guards against doesn't exist there.

## When to reach for this doc

- **Adding a new typed shape kind** (e.g. SmartArt) → §1.3 (which opaque tier?), §1.4 (where in the slide bucket order?), §2.1 (does it need a slide-scoped `cNvPrId`?), §3.2 (does the SVG factory need a new context field?).
- **Adding a new OOXML part type** (e.g. typed slide layouts) → §1.1 (snapshot scaffolding), §1.2 (a new `dirty` Set), §1.5 (does it need a counter?), §4.1 (extend the dirty-roundtrip test).
- **Lifting opaque content into a typed model** (e.g. exit animations) → §1.3 (model-vs-raw flip pattern), §4.1 (typed→raw→typed cycle).
- **Adding a new command** → §2.1 + §2.2 (addressing scheme), §4.3 (chain it through the multi-invocation CLI test).
- **Adding a renderer feature** → §3.2 (`SvgRenderCtx`), §3.3 (keep it out of `renderer/react/` if it doesn't need DOM), §3.4 (draw in EMU, not pixels).

## What we did **not** invent

For symmetry — the following are reused unchanged from `@officeai/core` / `@officeai/docx` and **don't** belong in a slides-divergences doc:

- `OoxmlContainer` (JSZip + byte-cache) — same in both formats.
- `RelationshipGraph`, `ContentTypes` helpers — typed-snapshot wrappers in PPTX (§1.1) wrap the same underlying utilities.
- `CommandBus<TSnapshot>` from `@officeai/core/commands` — generic over snapshot type, no PPTX fork.
- `HandlerContext` (`mintNodeId` + `now`) — identical for DOCX and PPTX. Slide-scoping is done by the helpers (`findSlide`, `withSlide`, `maxCNvPrId`) inside `commands/helpers.ts`, not by extending the framework.
- `parseXml` / `serializeXml` (`fast-xml-parser` with `preserveOrder`) — same configuration as DOCX.
- `sha256Hex` (via `js-sha56`) — isomorphic SHA-256 originally chosen for DOCX; reused here for media dedup and the dirty-roundtrip test.
