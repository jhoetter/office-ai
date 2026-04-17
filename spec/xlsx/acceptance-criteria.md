# XLSX — Acceptance Criteria

> Measurable done criteria. The XLSX format is "shipped" when every box
> below passes. These are the gates `make verify` and the browser smoke
> test enforce; the build log records pass/fail per fixture.

## Gates

### G1 — Round-trip integrity (load-bearing)

For every fixture under `fixtures/xlsx/synthetic/`:

- [ ] `parseXlsx(buf)` succeeds without error.
- [ ] `serializeXlsx(parseXlsx(buf))` succeeds without error.
- [ ] Re-parsing the serialized bytes produces a workbook **structurally
      equal** to the original (cell values, formulas, formats, sheet
      list, defined names, conditional formats, comments, hyperlinks).
- [ ] Every OOXML part **not modified** by the test edit is
      **byte-identical** between input and output (`zip.parts.get(path)
bytes equality`). Verified by `tests/roundtrip/xlsx/`.
- [ ] The serialized file opens in LibreOffice Calc headless without a
      repair dialog (`make roundtrip-libre` extends to xlsx).

### G2 — Agent API headless smoke

- [ ] `XlsxAgent.fromBuffer(buf)` succeeds.
- [ ] `getSnapshot()` returns a `XlsxSnapshot` with the expected sheet
      list and used-range cell counts.
- [ ] `getRange({ kind: "xlsx-cells", sheet: "Sheet1", range: "A1:B5" })`
      returns the projected cells.
- [ ] `applyCommand({ type: "xlsx:set-cell-value", ... })` produces a
      `Mutation` with status `"pending"` (when source `"agent"`) or
      `"approved"` (when source `"human"` / `"system"`).
- [ ] `approveMutation(id)` collapses pending into approved.
- [ ] `exportFile()` produces bytes that round-trip through G1.

### G3 — Command suite (P0)

Every P0 command listed in [`feature-scope.md`](feature-scope.md) has:

- [ ] A typed payload in `packages/xlsx/src/commands/payloads.ts`.
- [ ] A handler registered in `packages/xlsx/src/commands/registry.ts`.
- [ ] Vitest coverage for happy path + at least one edge case + the
      `precheck → { ok, reason }` path (for commands that have one).
- [ ] An inverse mutation that satisfies `apply(redo) ∘ apply(undo) ===
identity` on `before`. Asserted by a property test in
      `packages/xlsx/src/commands/handlers.test.ts`.

The 13 P0 commands:

- `xlsx:set-cell-value`
- `xlsx:set-cell-formula`
- `xlsx:set-range-values`
- `xlsx:set-cell-format`
- `xlsx:insert-row`
- `xlsx:insert-column`
- `xlsx:delete-row`
- `xlsx:delete-column`
- `xlsx:merge-cells`
- `xlsx:unmerge-cells`
- `xlsx:add-sheet`
- `xlsx:rename-sheet`
- `xlsx:add-comment`

### G4 — Formula correctness

For each function in the P0 categories (math/stats, logic, info, lookup, text):

- [ ] Vitest suite per function with at least 3 cases (basic + edge + error).
- [ ] Errors propagate per Excel semantics (`#VALUE!` for type mismatch, `#DIV/0!` for division by zero, etc.).
- [ ] Cell-cycle detection produces `#REF!` with `cycle` metadata in the diff.
- [ ] `RECALC after set-cell-formula` updates downstream `cachedValue`s
      in topological order.
- [ ] Formula text round-trips: `parse → serialize → reparse` keeps the
      formula string byte-identical (including absolute-ref markers and shared formulas).

### G5 — Performance budget

`scripts/perf-xlsx.mjs` (added in Phase 9) asserts:

- [ ] Parse a 50,000-row × 10-column synthetic fixture in **< 1.5 s** on
      Apple Silicon.
- [ ] Serialize same in **< 1.0 s**.
- [ ] Apply 1,000 `xlsx:set-cell-value` commands in **< 200 ms**.
- [ ] Recalc after a single edit on a workbook with 10,000 dependent
      formulas completes in **< 100 ms**.

These mirror the DOCX perf script in `scripts/perf-docx.mjs`.

### G6 — Renderer / browser

- [ ] `make dev` starts without error; `/editor` (or `/xlsx` route) loads.
- [ ] Drag-drop a fixture file → workbook renders, sheet tabs visible,
      first sheet scrolls smoothly.
- [ ] Clicking a cell shows it in the formula bar; typing replaces the
      value via `xlsx:set-cell-value` on the bus.
- [ ] Agent prompt panel runs a hard-coded recipe, lists pending
      mutations, "Approve" promotes them.
- [ ] Export → re-import in LibreOffice (manual smoke).
- [ ] Playwright `apps/web/e2e/xlsx.spec.ts` covers open → edit → export.

### G7 — CLI

- [ ] `pnpm --filter @officeai/agent exec office-agent xlsx inspect --file fixtures/xlsx/synthetic/01-plain-values.xlsx` prints sheet/cell counts.
- [ ] `xlsx read --file ... --sheet Sheet1 --range A1:F10 --format csv` prints CSV.
- [ ] `xlsx write --file in.xlsx --sheet Sheet1 --ref B2 --value 42 --out out.xlsx` writes; output round-trips through G1.
- [ ] `xlsx formula --file ... --sheet ... --cell G2 --formula '=SUM(A1:A10)' --out ...` writes and prints the cached value.
- [ ] `xlsx diff --before a.xlsx --after b.xlsx --format json` prints structured diff.

### G8 — MCP server

- [ ] Every `xlsx_*` tool from [`agent-commands.md`](agent-commands.md)
      is registered in `packages/agent/src/mcp.ts` with a JSON-schema
      input.
- [ ] `mcp.test.ts` round-trips at least one tool per category through
      the JSON-RPC stdio transport.

### G9 — License / architecture

- [ ] `make licenses` shows no AGPL / GPL-only / SSPL / BUSL.
- [ ] `node scripts/check-architecture.mjs`: `xlsx → core` only;
      `react/react-dom/next` forbidden in `@officeai/xlsx`.
- [ ] `pnpm typecheck` clean.
- [ ] `pnpm lint` clean.

### G10 — Docs

- [ ] All 11 spec docs in `spec/xlsx/` are present and up to date.
- [ ] `docs/build-log/xlsx.md` records non-trivial decisions, deferred
      items, validation summary.
- [ ] `README.md` status table moves XLSX to "active".
- [ ] `docs/session-summary.md` carries a recap.

## Non-criteria (intentionally not gated)

- Pixel-perfect chart rendering (we use image fallback in P0; criterion
  is "preserved through round-trip", not "rendered identically to Excel").
- Pivot table editing (out of scope; preserve only).
- Macro execution (out of scope; preserve only).
- Microsoft Office desktop manual smoke (we don't ship Office in CI;
  LibreOffice headless is the proxy).
