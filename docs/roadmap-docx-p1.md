# DOCX P1 Roadmap

> Status: 2026-04-17. P0 shipped (see `docs/session-summary.md`). This doc
> captures what we still owe to meet the [`prompt.md`](../prompt.md)
> "Office-compatible, AI-native DOCX editor" bar, plus what the public
> reference (`eigenpal/docx-js-editor`) shows in the field that we haven't
> built yet.

## Comparison snapshot

| Concern                                | Ours (P0)                                                                                    | `eigenpal/docx-js-editor`                                                                      | Verdict                                                                                           |
| -------------------------------------- | -------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------- |
| OOXML byte-preservation invariant      | First-class. SHA-256-checked per part. Untouched parts re-emitted from cached `Uint8Array`.  | Not a stated invariant; selective-save heuristics in `selectiveSave.ts`.                       | **We're ahead.** Keep this as a non-negotiable.                                                   |
| Headless agent API                     | `DocxAgent` (no DOM), exposed via `office-agent` CLI (5 subcommands).                        | `@eigenpal/docx-editor-agents` (AGPL). Browser editor not headless.                            | **We're ahead** on architecture, **behind** on CLI/MCP feature parity vs `prompt.md`.             |
| Pending / approved / working tri-state | In core; mutation diffs + approve / reject / rollback.                                       | None — there's only a single editor state.                                                     | **We're ahead.** Differentiator.                                                                  |
| Recognised OOXML coverage              | `w:p`, `w:r`, `w:hyperlink`, comments, tracked-changes wrappers, opaque blocks for the rest. | Parses tables, images, drawings, fields, headers/footers, footnotes, numbering, theme, styles. | **We're behind** on parsed coverage. P1 must cover headers/footers + tables for editable scope.   |
| Renderer                               | Single ProseMirror editor surface (`mountDocxEditor`).                                       | Dual-rendering: hidden ProseMirror + paginated `layout-painter`.                               | **We're behind** on pagination. Decide P1 vs P2.                                                  |
| Tests                                  | 64 unit + 17 integration. No E2E.                                                            | ~37 Playwright spec files + JSON-driven scenarios.                                             | **We're behind** on browser smoke tests.                                                          |
| MCP / agent embedding                  | None.                                                                                        | `packages/core/src/mcp` server + tool registry.                                                | **We're behind.** `prompt.md` §CLI explicitly calls for an agent surface; MCP is the obvious fit. |
| LibreOffice roundtrip in CI            | Manual only.                                                                                 | Not in CI either.                                                                              | Both behind; we should set the bar (it's our invariant).                                          |
| Real-world fixture corpus              | 5 synthetic + 10 reserved slots in `MANIFEST.md`.                                            | 17+ real .docx in `e2e/fixtures/`.                                                             | **We're behind** on coverage.                                                                     |
| Plugin / extension system              | Format-agnostic plugin scaffold in `packages/core` (used by agent staging).                  | Mature `PluginHost`, template plugin, extension system (26+ exts).                             | **We're behind** but plugin maturity is P2 — not on the roundtrip critical path.                  |
| Live collaboration (Yjs)               | Not started.                                                                                 | Yjs CRDT, awareness, multi-user.                                                               | **Both** unspecced for our P1; defer.                                                             |

**Bottom line:** the AI-native architecture is ours to win. The
production-readiness gaps are: (1) we don't yet roundtrip real-world Word
output, (2) we can't edit tables / images / headers, (3) the agent surface
we expose doesn't yet match what an external LLM would reach for (no MCP,
no `inspect`, no markdown read), (4) we have no E2E confidence. P1
sequences these.

---

## Themes

### Theme A — **Roundtrip surface area**

Goal: every fixture in `prompt.md` §"Fixture Corpus" round-trips clean.

| #   | Item                                                                                                                                                   | Spec ref                              | Effort | Sequencing |
| --- | ------------------------------------------------------------------------------------------------------------------------------------------------------ | ------------------------------------- | ------ | ---------- |
| A1  | Collect / generate the 10 real-world fixtures listed in `fixtures/docx/MANIFEST.md`.                                                                   | `prompt.md` §"DOCX Fixtures"          | M      | **P1.1**   |
| A2  | Add `make roundtrip-libre` that opens every fixture in headless LibreOffice, asserts no repair dialog. Wire into CI.                                   | `acceptance-criteria.md`              | S      | **P1.1**   |
| A3  | Multi-paragraph `format-range` / `delete-range` (lift the `multi-paragraph` `CommandError`).                                                           | `agent-commands.md` deviation row     | M      | **P1.1**   |
| A4  | Headers / footers: typed parse + serialize + `docx:set-header-text` / `docx:set-footer-text` commands.                                                 | `feature-scope.md` row                | M      | **P1.2**   |
| A5  | Tables: typed parse (rows / cells / `gridSpan` / `vMerge`) + `docx:insert-table` + `docx:set-cell-content` + `docx:insert-row` / `docx:insert-column`. | `analysis.md` §5, `agent-commands.md` | L      | **P1.3**   |
| A6  | Images: typed parse for inline drawings + `docx:insert-image` (mint relationship + media bytes + content type).                                        | `feature-scope.md` row                | M      | **P1.3**   |
| A7  | Numbering / lists: `docx:set-paragraph-list { numId, ilvl }` + `docx:remove-paragraph-list`. Preserve `numbering.xml`.                                 | `analysis.md` §5                      | M      | **P1.4**   |
| A8  | Hyperlink mutation: `docx:insert-hyperlink` / `docx:remove-hyperlink`.                                                                                 | `analysis.md` §5                      | S      | **P1.4**   |

### Theme B — **Agent surface**

Goal: the public agent surface matches `prompt.md` §AI-Native and §CLI to
the letter, and is reachable from any LLM (CLI, MCP, programmatic).

| #   | Item                                                                                                                                                                             | Spec ref                             | Effort | Sequencing |
| --- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------ | ------ | ---------- |
| B1  | Comment lifecycle: `docx:resolve-comment`, `docx:reply-comment`, `docx:delete-comment`. Drive `commentsExtended.xml`.                                                            | `agent-commands.md` stubs            | M      | **P1.1**   |
| B2  | Tracked changes: `docx:accept-change`, `docx:reject-change`.                                                                                                                     | `agent-commands.md` stubs            | M      | **P1.2**   |
| B3  | CLI parity with `prompt.md` §CLI: `docx inspect`, `docx read --format markdown`, `docx write --at "section:N/paragraph:M"`, `docx comment` (already works), and per-format help. | `prompt.md` lines 449–487            | S      | **P1.1**   |
| B4  | MCP server (`packages/agent/src/mcp.ts`): expose `docx_load`, `docx_save`, `docx_get_text`, `docx_apply_command`, `docx_diff`, `docx_search`. Single binary `office-agent mcp`.  | `prompt.md` §AI-Native; eigenpal-mcp | M      | **P1.1**   |
| B5  | Replace the hard-coded "`[AI] ` + comment" recipe in `apps/web` with a thin LLM caller (env-keyed, dispatches the same commands the user would).                                 | `docx.md` §"Known issues"            | S      | **P1.2**   |
| B6  | `agent.toMarkdown()` extension: section-aware, headings as `#`, lists as `-` / `1.`, tables as pipe-tables.                                                                      | `acceptance-criteria.md`             | S      | **P1.2**   |

### Theme C — **Renderer / UX**

Goal: the editor feels like Word, not like a textarea.

| #   | Item                                                                                                                           | Spec ref                  | Effort | Sequencing |
| --- | ------------------------------------------------------------------------------------------------------------------------------ | ------------------------- | ------ | ---------- |
| C1  | PM funnel: re-assert active marks across boundary edits (fixes "typing inside a partially-bold span loses the mark").          | `docx.md` §"Known issues" | S      | **P1.1**   |
| C2  | PM funnel: translate multi-block PM slices (paste of N paragraphs) into a sequence of insert-paragraph + insert-text commands. | `docx.md` §"Known issues" | M      | **P1.1**   |
| C3  | Toolbar parity: font family / size pickers, color, highlight, alignment, indentation, list controls.                           | `feature-scope.md`        | M      | **P1.2**   |
| C4  | Comments sidebar with thread view, scroll-to-highlight, resolve / reply controls (drives B1).                                  | eigenpal feature parity   | M      | **P1.2**   |
| C5  | Tracked-changes inline UI: accept/reject ribbon per change (drives B2).                                                        | eigenpal feature parity   | M      | **P1.2**   |
| C6  | Page layout / pagination preview (split visible surface into A4 pages, render headers/footers per page). Optional this round.  | eigenpal `layout-engine`  | L      | **P2**     |

### Theme D — **Quality / operational**

Goal: regressions surface in CI, not in the user's lap.

| #   | Item                                                                                                    | Spec ref                                                  | Effort | Sequencing |
| --- | ------------------------------------------------------------------------------------------------------- | --------------------------------------------------------- | ------ | ---------- |
| D1  | Playwright E2E suite for `apps/web`: open / type / Enter / format / comment / export — minimum 8 specs. | eigenpal e2e parity (37 specs there; we ship the first 8) | M      | **P1.1**   |
| D2  | Performance budgets in CI: `make perf-docx` opens / edits / serializes a 100-page doc within budget.    | `acceptance-criteria.md` perf                             | S      | **P1.2**   |
| D3  | License-graph SPDX scan in CI (today is a manual `pnpm licenses ls`).                                   | `acceptance-criteria.md`                                  | S      | **P1.2**   |
| D4  | OOXML schema validation: validate every serializer output against the ECMA-376 XSDs in CI.              | `serializer.md`                                           | M      | **P1.3**   |

---

## Sequencing & batches

P1 lands in **four batches**. Each batch is independently shippable
(green CI, docs updated, no half-built features behind a flag).

### Batch P1.1 — "Production roundtrip + agent reach" (this session, parallelizable)

Three independent workstreams, each with its own subagent. They touch
disjoint files so they merge clean.

- **W1 — Real-world fixtures + LibreOffice CI** → A1, A2, D1.
- **W2 — Range edits + PM funnel** → A3, C1, C2.
- **W3 — Agent reach** → B1, B3, B4.

### Batch P1.2 — "Editable Word features"

- **W4 — Headers/footers + tracked changes UI** → A4, B2, C5.
- **W5 — Toolbar + comments sidebar + LLM bridge** → C3, C4, B5, B6.
- **W6 — Performance + license CI** → D2, D3.

### Batch P1.3 — "Tables, images, schema validation"

- **W7 — Table model + commands** → A5.
- **W8 — Image insert + relationship machinery** → A6.
- **W9 — XSD schema validation in CI** → D4.

### Batch P1.4 — "Lists & hyperlinks"

- **W10 — Numbering / list mutation** → A7.
- **W11 — Hyperlink mutation** → A8.

After P1.4 the spec's P1 bucket is empty and we can pivot to XLSX.

---

## How we split for subagents (general rule)

We split by **file-ownership disjointness**, not by feature. Each
subagent gets a complete task brief that includes:

1. Goal, non-goals, acceptance criteria.
2. The exact files it owns (writes), and the files it may **read** but
   not write (so two agents never edit the same module).
3. The relevant spec sections it must respect.
4. Test expectations (unit + integration; whether e2e applies).
5. Build-log entry it must append.

Concretely for **batch P1.1** (which we execute now), the disjoint
ownership is:

| Subagent | Owns (writes)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                              | Reads only                                       |
| -------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ | ------------------------------------------------ |
| **W1**   | `fixtures/docx/real-world/`, `scripts/generate-real-fixtures.mjs`, `scripts/run-libreoffice-roundtrip.mjs`, `Makefile` (additive lines), `.github/workflows/ci.yml` (additive job), `apps/web/e2e/`, `apps/web/playwright.config.ts`, `tests/roundtrip/docx/real-world-roundtrip.test.ts`, append rows to `fixtures/docx/MANIFEST.md`, append section to `docs/build-log/docx.md`.                                                                                                                                                                                                                                                         | `packages/docx/src/**`, `apps/web/app/**`.       |
| **W2**   | `packages/docx/src/commands/format-range.ts`, `packages/docx/src/commands/delete-range.ts`, `packages/docx/src/commands/handlers.test.ts` (extend), `packages/docx/src/renderer/transaction-to-commands.ts`, `packages/docx/src/renderer/mount.ts`, `packages/docx/src/renderer/transaction-to-commands.test.ts` (new), append section to `docs/build-log/docx.md`.                                                                                                                                                                                                                                                                        | `packages/core/src/**`, `apps/web/**`, fixtures. |
| **W3**   | `packages/agent/src/cli.ts`, `packages/agent/src/mcp.ts` (new), `packages/agent/src/mcp.test.ts` (new), `packages/agent/src/cli.test.ts` (extend), `packages/docx/src/commands/{resolve-comment,reply-comment,delete-comment}.ts` (new), `packages/docx/src/commands/registry.ts` and `packages/docx/src/commands/index.ts` (additive), `packages/docx/src/commands/handlers.test.ts` is **avoided** — W3 puts comment-lifecycle tests in a new file `packages/docx/src/commands/comments-lifecycle.test.ts`. Append rows to `spec/docx/agent-commands.md` (status flip from stub to shipped). Append section to `docs/build-log/docx.md`. | `packages/core/src/**`, `apps/web/**`.           |

The only file that two subagents touch is `docs/build-log/docx.md` —
each appends its own section under its own H2 heading, so a
last-writer-wins merge is fine. `commands/registry.ts` is touched only
by W3 (W2 modifies existing handlers, doesn't register new ones). The
`make verify` quality gate runs after all three merge.

After P1.1 lands, batches P1.2–P1.4 follow the same pattern (3
disjoint workstreams per batch, dispatched in parallel).

---

## Out of scope this roadmap

- Yjs collaboration (defer to a Collaboration milestone after P1.4).
- Page-by-page paginated rendering (C6 — explicit P2).
- XLSX and PPTX (deferred per session-zero choice; the same architecture applies once DOCX is in P1.4).
- VBA macros and OLE — `prompt.md` §"Out of scope" makes these permanent.
