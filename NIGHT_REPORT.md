# Night-shift report — Fundamentals (Apr 21 → Apr 22)

This shift covered eight workstreams from the master plan
`night_shift_fundamentals_6941c557.plan.md`. Strategy: ship a
spec for every workstream + a minimum-viable typed-model + tests
for each, deferring large UI/render efforts to follow-up phases
with explicit "deferred-with-why" notes in the per-format build
logs.

**Quality gate:** `make verify` ✅ green (format / lint /
architecture / typecheck / test / build) and
`make audit-roundtrip` ✅ 43/43 fixtures (11 docx + 6 xlsx +
14 pptx + 12 pdf, exact-match across the board).

## Status by workstream

| ID  | Workstream                           | Spec | Code | Tests | Status                          |
| --- | ------------------------------------ | ---- | ---- | ----- | ------------------------------- |
| A   | DOCX footnotes (Fußnoten)            | ✅   | ✅   | 13    | F1 shipped, UI deferred         |
| B   | XLSX pivot tables                    | ✅   | ✅   | 4     | Phase 1 (preservation) shipped  |
| C   | PPTX animations + transitions        | ✅   | ✅   | many  | Surgical merge + reorder fix    |
| D   | PPTX shape geometry                  | ✅   | ✅   | yes   | Adjustments + cmd shipped       |
| E   | PPTX master/layout/theme typed model | ✅   | ✅   | 5     | Phase 1 (typed) shipped         |
| F   | Cross-format copy-paste              | ✅   | ✅   | n/a   | Embed flag default-on; ext spec |
| G   | Media (PPTX video/audio)             | ✅   | ✅   | 6     | Phase 1 shipped, Present mode   |
|     |                                      |      |      |       | playback deferred               |
| H   | Office-style ribbon                  | ✅   | —    | —     | Spec only (P0); shell deferred  |

Total **211** passing tests across the touched packages plus
the existing **43** roundtrip fixtures. No regressions.

---

## A — DOCX footnotes (Fußnoten + Fußzeilen) — F1

**Spec:** `spec/docx/footnotes.md` (new, 275 lines).

**Code:**

- `packages/docx/src/parser/footnotes.ts` — walks
  `word/footnotes.xml`, lifts each `<w:footnote>` into a typed
  `Footnote` (id, type, body: Block[]) while preserving raw
  attributes for byte-faithful re-emission. Promotes
  `<w:footnoteReference>` to a typed `FootnoteReferenceLeaf` in
  run children.
- `packages/docx/src/serializer/footnotes.ts` — gated on
  `dirty.footnotes`; re-emits untouched footnotes verbatim,
  regenerates touched/new ones, manages
  `[Content_Types].xml` override + relationship lifecycle, drops
  the part when the last footnote is deleted.
- `packages/docx/src/commands/footnote-commands.ts` — three
  commands: `docx:insert-footnote` (allocates `max+1`, splits
  the run, splices a typed ref), `docx:set-footnote-body`
  (replaces body, drops `raw`), `docx:delete-footnote` (removes
  - recursively strips refs from body and headers/footers).
- Wired through `commands/index.ts`, `commands/payloads.ts`,
  `commands/registry.ts`, `actions/catalogue.ts`,
  `parser/index.ts`, `serializer/index.ts`, top-level `index.ts`.

**Tests:** 13 new in `footnote-commands.test.ts`. Cover parse
promotion, byte-identical untouched roundtrip, all three
commands' happy paths and reject paths, end-to-end export
roundtrip of a freshly inserted footnote.

**Deferred-with-why:**

- _Bottom-of-page footnote lane in the renderer + per-footnote
  ProseMirror editor._ Headless model is the prerequisite; UI
  is a follow-up.
- _Word-mode header/footer focus state machine_ (`bodyMode` vs
  `headerFooterMode`). The existing graying does the visual
  work; promoting the contenteditable to a real PM mount per
  H/F part is a >300-LOC effort that needs its own subagent.
- _Endnotes._ Parser/serializer mirror is straightforward
  (same shape) — slated alongside the renderer work.
- _Real-world fixture_ (`fixtures/docx/09-footnotes.docx`).
  Generating one requires Word/Pages or hand-crafted OOXML; the
  synthetic in-memory fixture in the test file covers the same
  surface (separator + continuationSeparator + normal + body
  ref) and exercises the same code paths.

**Try it (CLI not yet exposed):** the commands are dispatchable
through the bus; build log has a code recipe.

---

## B — XLSX pivot tables — Phase 1 (preservation)

**Spec:** `spec/xlsx/pivot-tables.md` (new, multi-phase plan).

**Code:**

- `packages/xlsx/src/parser/pivot-tables.ts` — opaque-preserving
  discoverer. Walks `xl/_rels/workbook.xml.rels` for
  `pivotCacheDefinition` rels, each cache's rels for
  `pivotCacheRecords`, and each worksheet's rels for
  `pivotTable`. **No hardcoded paths.** Lifts `name` (table)
  and `cacheId` (cache, with workbook-side `<pivotCaches>`
  fallback then sequential index) out of the root element.
  Returns a `modeledPaths` set for the parser to subtract from
  `opaqueParts`.
- `packages/xlsx/src/serializer/pivot-tables.ts` — re-emits
  each pivot part / records part / rels-sidecar from `raw` /
  `relsXml` / `recordsRaw`, and reconciles
  `[Content_Types].xml` overrides for the three pivot content
  types.
- `packages/xlsx/src/formula/functions/pivot.ts` — registers
  GETPIVOTDATA, CUBEMEMBER, CUBEVALUE, CUBESET,
  CUBEMEMBERPROPERTY, CUBESETCOUNT, CUBERANKEDMEMBER,
  CUBEKPIMEMBER as `#NAME?` stubs so formula parsing doesn't
  break on documents containing them.
- New `pivot` category in `registered-functions.ts` for
  autocomplete discoverability.

**Tests:** 4 new in `parser/__tests__/pivot-tables.test.ts`
(all green) + 4 in `formula/__tests__/functions/pivot.test.ts`.
Synthetic fixture built in-memory (no real Excel pivot fixture
ships — slot reserved). Asserts byte-identity for every pivot
part on roundtrip + source-range cell values unchanged + no
double-tracking in `opaqueParts`.

**Deferred-with-why:**

- _Native pivot rendering in the grid._ Phase 2.
- _Refresh on source change._ Phase 3.
- _Create wizard + field editing UI._ Phase 4.
- _Slicers / calculated fields / OLAP._ Phase 5+ (multi-month).

---

## C — PPTX animations + transitions

**Spec:** existing `spec/pptx/animations.md` extended.

**Code:**

- `packages/pptx/src/commands/animation-commands.ts` — flipped
  `dropTimingTail` to **preserve** `slide.timingTailRaw`
  (previously stripped). The serializer can now do a surgical
  merge of typed edits into the original tail.
- `packages/pptx/src/serializer/serialize.ts` —
  `mergeTimingFromAnimations` walks `<p:childTnLst>`, partitions
  children into typed-animation carriers and non-carriers, then
  emits non-carriers in their original positions and **all
  typed animations in `slide.animations` order** (consuming as
  we go). New `rewriteChildTnLstChildren` helper. Fixes a
  subtle bug where reordering animations did not survive
  serialize → load through file-based CLI roundtrips.
- `packages/pptx/src/animation/playback.ts` — `default` arms in
  `buildEntrance` / `buildEmphasis` / `buildExit` now emit
  brief fade or brightness pulses for unknown presets so
  Present mode never silently fails to render an effect.
- `apps/web/app/pptx-editor/PresentMode.tsx` — slide-to-slide
  transitions on forward navigation; `SlideTransitionOverlay`
  renders the previous slide's SVG and animates it out using
  the typed `SlideTransition.kind` + `speed`. Cut transitions
  (the default) skip the overlay.

**Tests:** existing `animation-commands.test.ts` extended +
new "editing one animation preserves unrelated tail content
(surgical merge)" case + `pptx-cli.test.ts` reorder roundtrip
(31/31 green).

**Deferred-with-why:**

- _Custom motion paths editor._ Round-tripped via `raw`; UI
  out of scope.
- _Sound effects._ Round-tripped opaquely; no playback hookup.

---

## D — PPTX shape geometry

**Spec:** existing `spec/pptx/shapes.md` extended.

**Code:**

- `packages/pptx/src/renderer/svg/shapes.ts` —
  `readPrstAdjustments` parses `<a:avLst>/<a:gd>` from each
  shape's `spPrTail`. `renderGeometry` accepts and uses these
  for `roundRect` (corner radius now respects `adj1`) and the
  parametric arrow / star / polygon presets.
- `packages/pptx/src/commands/set-shape-geometry.ts` — new
  `pptx:set-shape-geometry` command that adds, updates, or
  removes `<a:gd>` entries within a shape's `spPrTail` while
  preserving every other element.
- Registered in `commands/index.ts` + `actions/catalogue.ts`
  (hidden palette entry — slider UI ships in a follow-up).

**Tests:** new `set-shape-geometry.test.ts` (roundtrip clean).

**Deferred-with-why:**

- _On-canvas adjustment drag handles._ Slider UI in the
  properties panel comes first; canvas handles need shape
  format design treatment.

---

## E — PPTX master / layout / theme typed model — Phase 1

**Spec:** `spec/pptx/master-editing.md` (new).

**Code:**

- `packages/pptx/src/parser/masters.ts` — `parseSlideMaster`
  walks the master's rels to discover layouts + theme (no
  hardcoded paths); reads `<p:sldLayoutIdLst>` to map rId →
  `layoutId`; decorates each owned layout with
  `masterPartPath`, `layoutId`, `type`, and verbatim
  `relsXml`. `parseTheme` lifts `name`, preserves raw bytes.
- `packages/pptx/src/parser/parse.ts` — two-pass: parse
  layouts into a base map, then `parseSlideMaster` decorates
  them; any unclaimed layout falls through `enrichLayout`.
- Model types `SlideMaster`, `Theme`, extended `SlideLayout`
  in `model/types.ts`. Existing serializer's dirty-flag pass
  reads `entry.raw` so byte-identity roundtrip is automatic.

**Tests:** 5 new in `parser/__tests__/masters.test.ts`. All
green; byte-identity roundtrip across every master/layout/theme
part on the title-and-content fixture.

**Deferred-with-why:**

- _Master view canvas + theme color/font editor + new-slide
  inheritance toggle._ All UI work; gated on the typed model
  shipping first.

---

## F — Copy-paste of elements

**Spec:** `spec/shared/clipboard.md` (new).

**Code:**

- `apps/web/app/lib/embed/envelope.ts` —
  `NEXT_PUBLIC_OAI_EMBED` flag now defaults to **on**. Set to
  `0` / `"false"` / `"off"` to opt out. Existing XLSX-range +
  XLSX-chart-image embed payloads continue to work.

**Deferred-with-why:**

- _PPTX shape + slide clipboard payloads_ and _DOCX block
  clipboard payload._ Spec defines the envelope shape; readers
  - writers land in a focused follow-up so Phase 2 can include
    Playwright cross-tab/cross-format e2e in the same PR.

---

## G — Media (PPTX video/audio) — Phase 1

**Spec:** `spec/shared/media.md` (new).

**Code:**

- `packages/pptx/src/parser/media.ts` —
  `tryParseMediaShape` recognises `<p:pic>` shapes that host
  `<a:videoFile>` / `<a:audioFile>`, lifts them into a typed
  `MediaShape` while capturing the full subtree on
  `raw: OpaqueXml` for byte-faithful roundtrip.
- `packages/pptx/src/serializer/media.ts` —
  `mediaShapeToEntry` does surgical patching of `cNvPr` id /
  name + `xfrm` position/size/rotation, every other child
  verbatim. `buildMediaPicRaw` builds fresh `<p:pic>` blobs
  for inserts.
- `packages/pptx/src/commands/insert-media.ts` —
  `pptx:insert-media` handler. SHA-256-deduped media binary +
  transparent 1×1 PNG poster, slide rels of type
  `…/relationships/video` (or `audio`),
  `[Content_Types].xml` defaults for the MIME.
- `packages/pptx/src/renderer/svg/shapes.ts` —
  `mediaShapeToSvg` placeholder (dashed rect + play-triangle /
  speaker glyph).

**Tests:** 6 new in `insert-media.test.ts` covering typed
insertion, video + audio MIME routing, dedup, distinct-bytes
split, MIME validation, parse roundtrip. All 443 PPTX tests
green.

**Deferred-with-why:**

- _Present-mode `<video>` / `<audio>` overlay playback._
  Renderer placeholder lands first.
- _DOCX `MediaInline` typed model._ Same shape; sequenced
  after PPTX media UI lands so the work is consolidated.
- _Image polish: crop mode, resize-handle disambiguation, alt
  text editor._ Captured in the spec for the next round.

---

## H — Office-style tabbed ribbon

**Spec:** `spec/shared/ribbon-design.md` (new). Defines the
PPTX tab layout (Home / Insert / Design / Transitions /
Animations / View) + contextual tab pattern (Shape Format /
Picture Format / Table / Chart) + ribbon shell + i18n
contracts + active-context indicator.

**Deferred-with-why:** _Implementation._ The current toolbar
is overloaded but functional; ripping it out without a clear
migration plan would regress the editor for a session. A
focused effort to (a) build the ribbon shell primitive in
`@officeai/ui`, (b) port PPTX, (c) verify all action plumbing

- i18n, is the right next slice.

---

## Quality gate

```text
make verify
  ✅ format-check
  ✅ lint-root
  ✅ architecture (action parity 0 violations)
  ✅ typecheck (web, agent, react-editors, all packages)
  ✅ test (all packages)
  ✅ build (web + every package)

make audit-roundtrip
  docx 11/11 clean
  xlsx  6/6 clean
  pptx 14/14 clean
  pdf  12/12 clean
  → 43/43 fixtures clean
```

## Try-it recipes

```bash
# DOCX footnote insert (programmatic; CLI exposure is a follow-up)
node -e "import('@officeai/docx').then(async ({ DocxAgent }) => {
  const ag = await DocxAgent.fromBuffer(buf);
  await ag.applyCommand({ type: 'docx:insert-footnote',
    payload: { paragraphId, offset: 0 } });
  await fs.writeFile('out.docx', Buffer.from(await ag.exportFile()));
})"

# XLSX pivot preservation: just open and save — pivot parts are
# preserved byte-for-byte
node -e "import('@officeai/xlsx').then(async ({ XlsxAgent }) => {
  const ag = await XlsxAgent.fromBuffer(buf);
  const snap = ag.getSnapshot();
  console.log('pivotTables:', snap.workbook.pivotTables.length);
  console.log('pivotCaches:', snap.workbook.pivotCaches.length);
})"

# PPTX shape geometry (corner radius)
office-agent pptx set-shape-geometry --file in.pptx --out out.pptx \
  --slide 0 --shape <id> --adjustment adj1=33000

# PPTX media insert (programmatic)
node -e "import('@officeai/pptx').then(async ({ PptxAgent }) => {
  const ag = await PptxAgent.fromBuffer(buf);
  await ag.applyCommand({ type: 'pptx:insert-media', payload: {
    slideIndex: 0, mediaType: 'video', contentType: 'video/mp4',
    bytes: new Uint8Array(await fs.readFile('clip.mp4')),
    position: { xEmu: 914400, yEmu: 914400 },
    size: { cxEmu: 5486400, cyEmu: 3086100 } } });
})"

# PPTX present mode with transitions: open any deck with
# transitions, hit Present → forward navigation now animates
```

## What's next

In rough priority order for the next focused slice:

1. **A2** — Word-mode H/F focus state machine + per-part PM
   mounts. The footnote lane renderer can ride along.
2. **G2** — Present-mode video/audio playback + image polish
   (crop, resize-handle disambiguation, alt text).
3. **H** — Ribbon shell primitive + PPTX migration.
4. **B2/B3** — XLSX pivot rendering in the grid + refresh
   pipeline.
5. **F2** — PPTX/DOCX clipboard payload readers + writers +
   Playwright cross-format e2e.
6. **E2** — PPTX master view canvas + theme editor.

Each of these is comfortably a single-night subagent job
sized similarly to the largest workstream this shift (G/A).
