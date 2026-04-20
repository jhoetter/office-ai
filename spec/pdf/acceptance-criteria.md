# PDF — Acceptance Criteria

> Per-feature acceptance matrix usable as the test plan. Every box
> must be green before the PDF night-shift PR ships.

Cross-references: feature buckets in
[`feature-scope.md`](./feature-scope.md);
performance budgets in [`performance.md`](./performance.md);
edge cases in [`edge-cases.md`](./edge-cases.md);
agent surface in [`agent-commands.md`](./agent-commands.md) and
[`cli.md`](./cli.md).

## Roundtrip integrity (the only non-negotiable bar)

For every fixture in `fixtures/pdf/synthetic/` (20 fixtures, see
`docs/build-log/pdf/fixtures.md`):

- [ ] **No-edit roundtrip is byte-identical.**
      `parsePdf(buf) → serializePdf(snap, buf) → bytesEqual(buf, output)`.
      Implementation note: the serializer detects "zero pending writes"
      and returns the original buffer unchanged.
- [ ] **Single-edit roundtrip is shape-clean and untouched-bytes
      preserved.** Apply a P0 command (rotate / reorder / delete /
      annotate / fill / set-metadata) → serialize → re-parse →
      assert (a) the typed model reflects the edit, (b) every byte
      position 0..originalLength-1 of the output matches the input
      byte-for-byte (incremental update appends only).
- [ ] **Output opens cleanly.** Run `qpdf --check` against every
      edited output; zero errors and zero warnings.
- [ ] **No "this file was modified / repair needed" warnings** in
      Adobe Acrobat Reader, Preview.app, or Chrome's built-in viewer
      across the 20 fixtures (manual QA checklist documented in
      `docs/build-log/pdf/visual-qa.md`).

## Engine

- [ ] `selectEngine()` returns the expected backend for every hint
      combination ([`engine-strategy.md`](./engine-strategy.md));
      tested in
      [`packages/pdf-engine/src/select-engine.test.ts`](../../packages/pdf-engine/src/select-engine.test.ts).
- [ ] PDF.js backend opens all 20 fixtures, renders page 1, returns
      non-empty text content for non-scan fixtures.
- [ ] PDFium backend opens 3 fidelity-critical fixtures (Type3 font,
      DeviceN color, custom CMap) and produces visually correct
      rasterizations within < 5% pixel diff vs PDFium baseline.
- [ ] Lazy-load discipline: `@embedpdf/pdfium` is **not** present in
      the eager bundle (verified by
      `bundlesize`).
- [ ] Headless Node renderer produces deterministic PNGs across
      runs.

## Parser

- [ ] All synthetic fixtures parse without error.
- [ ] Encrypted fixture parses with `--password test` and rejects
      without it.
- [ ] Broken-xref fixture parses with `warning: "xref-recovered"`.
- [ ] Tagged fixture parses with non-empty `PdfStructTree`.
- [ ] PDF/A fixture parses; metadata flags PDF/A conformance.
- [ ] Multilingual fixture (CJK + Arabic + German umlauts) parses;
      text projection is bidi-correct.
- [ ] Scan fixture (no text layer) parses; every page reports
      `hasTextLayer: false`.
- [ ] AcroForm fixture parses; all field types are surfaced
      correctly.
- [ ] XFA-only fixture parses; `list-form-fields` reports
      `"backing": "xfa-only"`.
- [ ] 1000-page stress fixture parses in under 5 s (cold).

## Commands

For each typed command, at least one **handler unit test** passes:

- [ ] `pdf:rotate-pages` — single page, multiple pages, all pages,
      various deltas.
- [ ] `pdf:set-page-rotation` — absolute set, identity (no-op), each
      of 0/90/180/270.
- [ ] `pdf:reorder-pages` — identity permutation (no-op), reverse,
      single-swap, full shuffle.
- [ ] `pdf:delete-pages` — single page, multiple, contiguous,
      non-contiguous; rejects "delete all pages".
- [ ] `pdf:set-metadata` — single field, multiple fields, empty
      patch (no-op).
- [ ] `pdf:add-bookmark` — root-level, nested under a parent,
      rejects unknown parent.
- [ ] `pdf:add-comment` — first comment, multiple comments per page.
- [ ] `pdf:reply-comment` — reply to existing, rejects unknown
      parent.
- [ ] `pdf:edit-comment` — text mutation only.
- [ ] `pdf:resolve-comment` — toggle on, toggle off.
- [ ] `pdf:delete-comment` — leaf comment, comment with replies
      (cascade).

For each command, at least one **integration test** passes:

- [ ] Parse → command → serialize → re-parse → typed model matches
      expectations.
- [ ] Untouched bytes of the same fixture are byte-identical.
- [ ] `qpdf --check` reports no errors.

## Page operations (CLI)

For each operation in [`editing-pipeline.md`](./editing-pipeline.md):

- [ ] `office-agent pdf rotate` — round-trip a fixture; `--pages`
      and `--angle` validated.
- [ ] `office-agent pdf reorder` — round-trip; permutation validated.
- [ ] `office-agent pdf delete-pages` — round-trip; non-empty result
      enforced.
- [ ] `office-agent pdf insert-pages` — insert from another fixture;
      page count grows by the inserted count.
- [ ] `office-agent pdf extract-pages` — output is a fresh PDF with
      only the requested pages.
- [ ] `office-agent pdf split --by range|bookmark|size` — output
      directory contains the expected files.
- [ ] `office-agent pdf merge` — concatenated page count matches
      the sum of inputs.
- [ ] `office-agent pdf crop` — `/CropBox` reflects the margin;
      `/MediaBox` unchanged.
- [ ] `office-agent pdf watermark` — output has the watermark XObject
      composited; original content untouched.
- [ ] `office-agent pdf add-page-numbers` — text-show operators
      appended.
- [ ] `office-agent pdf set-metadata` — `/Info` and XMP both updated.

## Annotations

For each annotation kind in [`annotation-model.md`](./annotation-model.md):

- [ ] Created via the typed model + AP-stream emitter; output opens
      in Adobe Acrobat / Preview / Chrome at the correct coordinates
      with the correct color.
- [ ] FDF/XFDF round-trip: write → export → re-import → identical
      typed annotations.
- [ ] Native annotation created in Adobe Acrobat survives a load +
      no-op save in our viewer (incremental save preserves it).
- [ ] Redaction mark → apply → underlying text is gone (selecting
      the redacted region returns empty).

## Forms

- [ ] AcroForm fixture: every field type enumerated correctly.
- [ ] `office-agent pdf fill-form` populates values; output reopened
      in Acrobat shows the values rendered.
- [ ] `office-agent pdf flatten-form` produces a non-fillable PDF
      that still displays the values.
- [ ] `office-agent pdf reset-form` restores defaults.
- [ ] FDF/XFDF form-value import populates the correct fields.
- [ ] XFA-only fixture: viewer banner shown; CLI flags
      `"backing": "xfa-only"`.
- [ ] Signature fixture: `office-agent pdf list-signatures` reports
      `valid: true`; incremental save preserves validity.

## Search

- [ ] Plain-text query returns expected matches on a known fixture.
- [ ] Case sensitivity toggle reduces matches as expected.
- [ ] Whole-word toggle reduces matches as expected.
- [ ] Regex query with `\bword\b` matches identically to whole-word
      mode.
- [ ] Invalid regex throws a clear error.
- [ ] Search on a scan fixture returns `[]`; viewer offers OCR.
- [ ] Performance: first hit on the 500-page fixture ≤ 200 ms p95.

## OCR

- [ ] `office-agent pdf ocr --lang deu+eng` adds a text layer to a
      scan fixture; `office-agent pdf search` then finds expected
      tokens.
- [ ] OCR output round-trips: re-parse the OCR'd file → text layer
      is non-empty.

## Accessibility

- [ ] axe-core reports zero WCAG 2.2 AA violations on `/pdf-viewer`
      with a fixture loaded.
- [ ] Lighthouse accessibility score ≥ 95 on `/pdf-viewer`.
- [ ] Reflow mode renders the tagged fixture's headings + paragraphs
      in document order.
- [ ] Keyboard reachability: every toolbar button focusable via
      `Tab`; `Ctrl+/` opens the shortcuts cheatsheet.
- [ ] Screen-reader smoke (VoiceOver scripted) reads page 1 of a
      tagged fixture in correct order.
- [ ] `prefers-reduced-motion`: animations disabled (visual diff
      shows no motion).
- [ ] `prefers-contrast: more`: contrast ≥ 7:1 verified by axe.

## Dark mode

- [ ] `Ctrl+Alt+D` toggles dark mode; `prefers-color-scheme`
      respected on first open.
- [ ] Smart-invert preserves a chromatic figure (visual snapshot
      diff on a known fixture).
- [ ] Smart-invert inverts background + text on a text-only fixture
      (visual snapshot diff).
- [ ] Per-document opt-out persists across reloads.
- [ ] Headless `office-agent pdf render` ignores dark mode (always
      original colors).

## Performance

- [ ] Cold load 50-page text PDF < 600 ms p95 (Mac M1, local file).
- [ ] Cold load 50-page mixed-content PDF < 1.5 s p95.
- [ ] Page render p95 < 150 ms.
- [ ] Scroll FPS 60 sustained on 1000-page fixture (no frame > 20 ms
      in any 5 s window).
- [ ] Memory < 600 MB on 1000-page fixture.
- [ ] Memory < 1.5 GB on 200 MB stress fixture.
- [ ] Bundle (eager) < 800 KB gzipped.

## Agent

- [ ] `PdfAgent.fromBuffer(buf).then(a => a.exportFile())` returns a
      buffer byte-equal to `buf` (no-edit roundtrip via the agent
      surface).
- [ ] `agent.applyCommand(cmd)` returns a `Mutation` with non-empty
      `diff.changes` for every typed command.
- [ ] `agent.applyCommand({ source: "agent", … })` puts the
      mutation in pending; `getPendingMutations()` returns it;
      `approveMutation(id)` moves it to approved.
- [ ] `agent.toMarkdown()` returns a page-by-page outline (one
      heading per page, then text paragraphs).
- [ ] `agent.search({ query: "…" })` returns matches with
      `pageNumber`, `start`, `end`, `match`, `preview`.
- [ ] `agent.getRange({ kind: "pdf-pages", start, end })` returns
      the typed page projection.
- [ ] **Headless invariant.** `@officeai/pdf` does not transitively
      import `react`, `react-dom`, `next`, or any DOM global.
      Enforced by `scripts/check-architecture.mjs`.

## CLI

- [ ] Every command listed in [`cli.md`](./cli.md) runs against at
      least one fixture and produces JSON validating against the
      documented schema (zod-checked in `tests/agent/pdf/`).
- [ ] Exit codes: `0` on success, `1` on bad args, `2` on engine
      error, `3` on IO error.
- [ ] `--password` works on the encrypted fixture.
- [ ] `office-agent pdf apply --commands ./batch.json` runs commands
      in order and rolls back on the first failure.
- [ ] `office-agent pdf diff` reports rotation changes / annotation
      count deltas / metadata changes correctly between two known
      revisions.

## MCP

- [ ] Every CLI command surfaced as an MCP tool by
      `packages/agent/src/pdf-mcp.ts`.
- [ ] Tool descriptors validate against MCP schema.
- [ ] Smoke test: Claude Desktop / Cursor can call `pdf_inspect` and
      `pdf_search` against a fixture.

## Realtime collaboration

- [ ] `usePublishPresence` broadcasts the current page + cursor
      position; remote `RemotePresenceList` displays it.
- [ ] `useCommandBroadcast` broadcasts P0 commands; remote agent
      receives and applies; both snapshots converge.
- [ ] `pdf-region` comment anchor remains attached after
      `pdf:set-page-rotation`.

## Build & lint gates

- [ ] `pnpm typecheck` green across all PDF packages.
- [ ] `pnpm lint:root` clean, no new warnings.
- [ ] `make audit-roundtrip` passes; PDF rows added with attribute
      counts (pages, annotations, fields, outline depth, fonts).
- [ ] `make e2e-web` includes a basic PDF flow: load fixture →
      render first page → search → annotate → save.
- [ ] `scripts/check-architecture.mjs` enforces: - `@officeai/pdf` cannot import `pdfjs-dist` or
      `@embedpdf/pdfium` directly. - Any package outside `packages/pdf-engine/src/backends/`
      cannot import `pdfjs-dist` or `@embedpdf/pdfium`. - `@officeai/pdf-engine` cannot import `react` /
      `react-dom` / `next`. - `@officeai/pdf` cannot import `react` / `react-dom` /
      `next`.

## Manual visual QA

Documented in `docs/build-log/pdf/visual-qa.md`. For each fixture:

- [ ] Open in our viewer + Chrome built-in + Adobe Acrobat Reader.
- [ ] Take a screenshot of page 1 + a representative middle page.
- [ ] Compare side-by-side; document any visible delta.
- [ ] Edit a P0 operation; reopen the output in Adobe / Preview /
      Chrome; confirm no warning dialogs.

## Wake-up report

- [ ] `NIGHT_REPORT_PDF.md` exists and mirrors the format of
      [`NIGHT_REPORT.md`](../../NIGHT_REPORT.md): TL;DR table,
      per-phase commits + try-it recipes, deferred items + why.
- [ ] `docs/build-log/pdf/10-of-10-demo.md` exists and contains a
      60-second guided demo script proving the 10/10 bar wins from
      [`prompt-pdf.md`](../../prompt-pdf.md) §"The 10/10 Bar".

## Definition of done

The night-shift PR can ship when:

1. Every checkbox above is green.
2. `make audit-roundtrip` is green and includes PDF.
3. The branch builds clean on a fresh checkout.
4. The visual QA log is complete with no unresolved deltas.

Anything else is scope; this is the bar.
