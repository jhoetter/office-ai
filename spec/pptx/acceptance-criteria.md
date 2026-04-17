# PPTX — Acceptance Criteria

> Measurable bar for declaring the PPTX phase done. Every box must be
> green before P10 closes.

## Roundtrip integrity (the only non-negotiable bar)

For every fixture in `fixtures/pptx/synthetic/`:

- [ ] **No-edit roundtrip is byte-identical.**
  `parsePptx(buf) → serializePptx(snap) → bytesEqual(buf, output)`.
  Verified by the SHA-256 of every part being unchanged in
  `partHashes`.
- [ ] **Single-edit roundtrip is shape-clean and untouched-parts byte-identical.**
  Apply `pptx:set-text` (or a P0 command of the test's choice) →
  serialize → re-parse → assert (a) the dirty part round-trips to the
  same typed model after re-parse, and (b) every other part's
  SHA-256 matches the input.

## Parser

- [ ] All synthetic fixtures parse without error.
- [ ] A fixture with a SmartArt `graphicFrame` parses; the SmartArt
      is captured as an `OpaqueShape` with `tag === "p:graphicFrame"`.
- [ ] A fixture with an embedded chart parses; the chart is captured
      as an `OpaqueShape` with `tag === "p:graphicFrame"`. The
      `ppt/charts/*` parts are present in `container.parts` and
      preserved verbatim on serialize.
- [ ] A fixture with a `notesSlide` parses; the slide carries
      `notesSlidePartPath`; the notes-slide part is captured as
      `OpaquePart`.
- [ ] A fixture with a group shape parses; the group is typed; its
      children are typed (or opaque if non-recognized).
- [ ] A fixture with a placeholder text shape that omits `<a:xfrm>`
      parses; the shape's `position`/`size` are `undefined` and the
      renderer falls back to layout/master.

## Commands

For each command (10 total), at least one handler unit test passes:

- [ ] `pptx:add-slide` — append + insert at index, optional layout
- [ ] `pptx:delete-slide` — drops part + rels + override; cleans up
      attached notes slide
- [ ] `pptx:duplicate-slide` — clones with fresh ids, no media duplication
- [ ] `pptx:move-slide` — reorders sldIdLst only; no part renaming
- [ ] `pptx:set-text` — single-paragraph + multi-paragraph (`\n`-split)
- [ ] `pptx:set-position` — TextShape, Picture, GroupShape; rejects opaque
- [ ] `pptx:set-size` — TextShape, Picture, GroupShape; rejects opaque
- [ ] `pptx:format-text` — subset-of-runs, full-paragraph
- [ ] `pptx:insert-image` — fresh insert + dedup-on-second-insert
- [ ] `pptx:add-text-box` — appended at end of z-stack

For each command, at least one **integration test** passes:

- [ ] Parse → command → serialize → re-parse → assert the post-serialize
      typed model matches expectations.
- [ ] Untouched parts of the same fixture remain byte-identical.

## Agent

- [ ] `PptxAgent.fromBuffer(buf).then(a => a.exportFile())` returns a
      buffer byte-equal to `buf` (no-edit roundtrip via the agent surface).
- [ ] `agent.applyCommand(cmd)` returns a `Mutation` with a non-empty
      `diff.changes` for every P0 command.
- [ ] `agent.applyCommand({ source: "agent", … })` puts the mutation in
      the pending queue; `agent.getPendingMutations()` returns it; calling
      `approveMutation(id)` moves it to approved.
- [ ] `agent.toMarkdown()` returns a slide-by-slide outline (one heading
      per slide, then bulleted text).
- [ ] `agent.search({ query: "…" })` returns matches with `slideIndex`,
      `shapeId`, and a preview.
- [ ] `agent.getRange({ kind: "pptx-shape", slideIndex, shapeId })`
      returns the shape projection (text + position/size).
- [ ] **Headless invariant.** `@officeai/pptx/agent` does not transitively
      import `react`, `react-dom`, `next`, `prosemirror-view`, or any
      DOM global. Enforced by `scripts/check-architecture.mjs`.

## Renderer

- [ ] `slideToSvgString(slide, ctx)` returns a non-empty string for every
      shape kind (TextShape, Picture, GroupShape, OpaqueShape).
- [ ] `emuToPx(pxToEmu(p)) === p` (round-trip identity for typical px
      values).
- [ ] `slideAspectRatio` returns the correct ratio for both 4:3 and 16:9
      fixtures.
- [ ] React canvas: dragging a shape produces exactly one
      `pptx:set-position` command at `pointerup`.
- [ ] React canvas: resizing a shape produces exactly one `pptx:set-size`
      command at `pointerup`.
- [ ] React canvas: editing a text shape produces exactly one
      `pptx:set-text` command at `blur`.
- [ ] Thumbnails reuse `slideToSvgString` and re-render only when the
      slide's `partHash` changes.

## Web app

- [ ] Homepage has a "Open the PPTX editor" CTA next to the DOCX one.
- [ ] `/pptx-editor` route loads; uploading any synthetic fixture
      renders the deck.
- [ ] Toolbar exposes: Open, Save, Add slide, Delete slide, Duplicate
      slide, Insert image, Add text box, Bold/Italic/Underline,
      Font size, Color.
- [ ] Slides sidebar shows thumbnails; clicking a thumbnail switches
      the active slide.
- [ ] Agent prompt panel produces a hard-coded "AI" recipe whose
      mutations land in the pending queue with a violet-outlined
      visual marker on affected shapes.
- [ ] Approve / reject buttons in the pending queue work end-to-end.
- [ ] Saving downloads a valid `.pptx` whose unedited parts are
      byte-identical to the upload.

## CLI / MCP

- [ ] `office-agent pptx --help` lists every PPTX command.
- [ ] `office-agent pptx inspect -i deck.pptx` prints slide count,
      shapes per slide, text content summary.
- [ ] `office-agent pptx read -i deck.pptx --slide N --format json`
      prints the typed slide projection.
- [ ] `office-agent pptx set-text -i in.pptx -o out.pptx --slide N
      --shape-id ID --text "X"` writes the file with the edit applied.
- [ ] All other PPTX subcommands (`add-slide`, `delete-slide`,
      `duplicate-slide`, `move-slide`, `set-position`, `set-size`,
      `format-text`, `insert-image`, `add-text-box`, `apply`) work and
      have a vitest CLI test.
- [ ] MCP server exposes the PPTX tools to a connected MCP client.

## Quality gates

- [ ] `make verify` (format-check, lint, architecture, typecheck, test,
      build) is green.
- [ ] `pnpm --filter @officeai/pptx test` is green.
- [ ] `pnpm --filter @officeai/integration-tests test` is green
      (PPTX integration suite added under `tests/` mirroring DOCX).
- [ ] No new AGPL / GPL-only / SSPL / BUSL runtime dependency added
      (`make licenses` clean).

## Manual smoke

- [ ] Browser smoke via `cursor-ide-browser` MCP: open
      `http://localhost:3000`, click "Open the PPTX editor", upload
      `fixtures/pptx/synthetic/05-with-image.pptx`, exercise:
      - text edit (set-text)
      - drag a shape (set-position)
      - resize a shape (set-size)
      - add a slide (add-slide)
      - export → re-upload → renders identically
- [ ] LibreOffice headless roundtrip on the same fixture, when `soffice`
      is on PATH (skipped gracefully otherwise).

## Documentation

- [ ] `docs/build-log/pptx.md` lists every spec deviation, deferred
      feature, and known issue.
- [ ] `docs/session-summary.md` has a PPTX section describing what was
      built, what passes, what's deferred, what was harder than expected.
- [ ] `packages/pptx/README.md` no longer says "deferred" — it points
      to the spec and lists the ten commands.
- [ ] `spec/pptx/README.md` likewise.
