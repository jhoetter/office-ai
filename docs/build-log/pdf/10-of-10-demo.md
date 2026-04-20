# 10/10 PDF Viewer — Demo Script

A ten-minute scripted walkthrough that proves the PDF viewer hits the
quality bar set in `prompt-pdf.md`. Every step references a specific
fixture or command so the demo is reproducible.

> **Setup:** `pnpm install && pnpm --filter @officeai/web build && pnpm --filter @officeai/web start --port 3100`
> Open `http://localhost:3100`.

---

## Act 1 — "Faster than Chrome's built-in viewer" (90 s)

1. **Drag** `fixtures/pdf/large-50page.pdf` onto the home tile grid.
2. The shell's `LoadingScreen` (variant=`splash`, product=`pdf`) fades
   in within ~80 ms; the editor swaps in once the agent and engine are
   ready. Note: thumbnails populate as you scroll, not all at once —
   the page tile cache is virtualized.
3. Use `Cmd/Ctrl + F` → search for "page 25". The status bar updates
   live: `pdf.pageOf` ("Page 25 of 50").
4. Click the page-input field, type `42`, press Enter. The viewport
   jumps; the active thumbnail in the sidebar lights up.

> Talking point: we render with PDF.js by default — same engine Chrome
> uses — but we hand the canvas off to a **virtualized tile grid** with
> a HiDPI cache, so scrolling stays butter-smooth even on a Retina
> display where Chrome can hitch on text-heavy PDFs.

## Act 2 — "Better looking than Acrobat Web" (60 s)

5. Toggle smart dark mode (`pdf.darkMode` button in the toolbar).
   Pages re-tint; embedded photographs in the test PDF stay full-colour.
6. Switch view mode → `pdf.twoUp`. Two pages render side-by-side; the
   thumbnail rail and the search highlights stay in sync.
7. Switch back → `pdf.continuous`. Hit `Cmd/Ctrl + 0` → fit-page;
   then `Cmd/Ctrl + Shift + 0` → fit-width.

> Talking point: smart dark mode isn't a CSS filter on the canvas —
> the inverter walks the rendered tile, detects raster image regions,
> and skips them. Photo-heavy PDFs (annual reports, scanned books)
> stay readable instead of going negative-photo.

## Act 3 — "More capable than Preview.app" (180 s)

8. Open `fixtures/pdf/acroform-basic.pdf`. The form fields highlight
   in the canvas; the **Forms** sidebar tab lists every widget by name
   with type + required / readOnly badges.
9. Fill the form inline. Click **Form → Flatten** in the toolbar. The
   widgets disappear; the values are now baked into page content.
10. Open `fixtures/pdf/with-outline.pdf`. The **Outline** sidebar
    populates from the document's `/Outlines` tree. Click any node —
    the viewport scrolls to that page.
11. From the toolbar, **Page Ops → Rotate Pages** → `1, 3` → `90°`.
    The thumbnails reorient; the canvas re-renders.
12. **Page Ops → Reorder**. Drag thumbnail 3 above thumbnail 1. The
    document re-renumbers; the agent emits a `reorder-pages` command
    on the bus, so `Cmd/Ctrl + Z` undoes it cleanly.
13. **File → Download PDF**. Open the downloaded file in Preview.app:
    no "this file was modified, repair needed" dialog. Page rotation
    and reorder persisted.

> Talking point: every mutation goes through the same command bus the
> DOCX, XLSX and PPTX editors share. Undo / redo, multi-user awareness
> and the agent CLI all observe the same stream.

## Act 4 — "More agent-friendly than ChatPDF" (180 s)

14. In a terminal:

    ```bash
    office-agent pdf read-metadata fixtures/pdf/metadata-rich.pdf | jq .
    ```

    A versioned JSON envelope prints: `{ "version": 1, "kind": "pdf.metadata", "data": { … } }`.

15. Multi-step flow — chain three commands without any LLM-side state:

    ```bash
    office-agent pdf rotate-pages fixtures/pdf/simple-text-3page.pdf \
      --pages 1 --degrees 180 --out /tmp/r.pdf
    office-agent pdf set-metadata /tmp/r.pdf \
      --title "Demo" --author "Night Shift" --out /tmp/rm.pdf
    office-agent pdf read-metadata /tmp/rm.pdf | jq .
    ```

    Both the rotation and the metadata changes are present in the
    output. Pages were re-serialized via `sourceIndex` mapping so the
    rotated objects survive the second round-trip.

16. Show the MCP tool catalogue:

    ```bash
    office-agent mcp list-tools | jq '[.[] | select(.name | startswith("pdf_"))]'
    ```

    Sixteen `pdf_*` entries — every CLI command is reachable from any
    MCP-aware LLM client (Claude Desktop, Cursor MCP, etc.) without a
    second integration.

17. Trigger an error envelope intentionally:

    ```bash
    office-agent pdf read-page fixtures/pdf/simple-text-1page.pdf --page 99
    # → exit 1; stderr is a structured JSON envelope, no stack trace.
    ```

> Talking point: the **product** ships zero LLM logic. All AI flows
> are owned by external models calling these CLI / MCP surfaces — the
> exact split the user briefed.

## Act 5 — "Respectful of user privacy" (60 s)

18. Open dev-tools → Network. Reload the viewer with any local PDF.
    Note: zero outbound requests beyond the same-origin chunk fetch.
    PDF.js worker spins up locally, parsing happens in-memory, no
    cloud round-trip.
19. Open the Realtime panel from the shell header — the presence
    avatar uses an anonymous adjective + animal handle (no user PII)
    drawn from the `ANONYMOUS_NAME_POOL` shared with the other
    editors.

> Talking point: the viewer respects the same "your bytes never leave
> the tab" guarantee the rest of the suite makes. Even the optional
> OCR adapter runs `tesseract.js` in a worker — no remote inference.

## Act 6 — "Roundtrip integrity" (90 s)

20. Run the audit live:

    ```bash
    make audit-roundtrip-pdf
    ```

    All twelve fixtures print `✓ … (N attrs, exact match)`; the JSON
    summary is appended to `docs/build-log/roundtrip-audit-night.json`.

21. Diff a single fixture by hand:

    ```bash
    office-agent pdf rotate-pages fixtures/pdf/simple-text-3page.pdf \
      --pages 2 --degrees 90 --out /tmp/x.pdf
    office-agent pdf rotate-pages /tmp/x.pdf --pages 2 --degrees 270 \
      --out /tmp/y.pdf
    diff <(office-agent pdf read-metadata fixtures/pdf/simple-text-3page.pdf) \
         <(office-agent pdf read-metadata /tmp/y.pdf)
    # → empty (modulo the producer string, which we deliberately rewrite)
    ```

> Talking point: this is the non-negotiable acceptance bar from the
> brief — open any real-world PDF, mutate, save, reopen anywhere, no
> data loss, no repair dialog. Twelve byte-stable fixtures prove the
> envelope; thirty-one round-trip vitests prove the corners.

---

## Closing slide checklist

- [x] Spec set — clean-room, pre-build (`spec/pdf/`).
- [x] Engine layer with strategy → default + fallback.
- [x] Headless agent + command bus + undo/redo.
- [x] Page-level edits, annotations, forms, OCR adapter.
- [x] Web viewer with full toolbar, sidebar, virtualized canvas.
- [x] CLI + MCP surface so any LLM can drive the document.
- [x] i18n EN + DE.
- [x] Realtime presence (`PdfSelection`).
- [x] Comment anchors (`pdf-region`).
- [x] 12 fixtures + 31 round-trip tests + 5 agent flows.
- [x] `make audit-roundtrip-pdf` — 12/12 clean.
- [x] `pnpm -r typecheck` + `pnpm -r lint` — green.

If you can run this script end-to-end on the `night/pdf-viewer-2026-04-20`
branch in ten minutes flat, the bar is met.
