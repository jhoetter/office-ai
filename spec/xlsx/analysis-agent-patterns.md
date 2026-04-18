# Agent Integration Patterns — Analysis (clean-room)

> Reference commits (depth-1 clones, `/tmp/refs/`):
>
> - `dream-num/univer-mcp` @ `52bb1a2aa5f2341c59373a8a461b3738bca5dd00`
> - `dream-num/skills` @ `0943f3f1093a348f3665abd24272d664d67176a9`
>
> Cross-reference (this repo): `packages/agent/src/mcp.ts`,
> `packages/agent/src/cli.ts`, `spec/docx/agent-commands.md`.
>
> Method: I read the public READMEs, every `SKILL.md`, every playbook, every
> reference doc, and every example in the two repos. Concepts only — no
> code, JSON shape, identifier, prose, or tool name was copied. Where I
> describe Univer behavior I am paraphrasing observable, documented surface.

---

## 0. The big surprise: there is almost no `univer-mcp` source to read

`dream-num/univer-mcp` is a **documentation-only repository**. It ships a
README, a license, and nothing else. The actual MCP server runs as a
hosted SaaS endpoint (`mcp.univer.ai/mcp/?univer_session_id=…`) and is
proprietary. The README also describes a separate browser/Node "Univer
instance" plus a `mcp-bridge` plugin that proxies tool calls to that
instance. None of those three components are in the repo.

Practical consequence for us:

1. We can read **architectural intent** off the README (transport, auth,
   session model, multimodal output, WIP scope) but we **cannot** look at
   their tool schemas, range serializer, formula round-trip, or diff
   output. Anything I say about those is inference, not citation.
2. The `dream-num/skills` repo is the higher-signal artifact. It defines
   two skills — `agent-sheet` (a local shell-native CLI) and `sheet-git`
   (a Git-shaped review/origin layer beside it) — and the skill files
   themselves enumerate the CLI command surface, the value-type model,
   the verification discipline, and the JS-API escape hatch in detail.
3. Univer's MCP server and Univer's local CLI (`agent-sheet`) are
   **different products with different surfaces**. The MCP one is
   hosted, multimodal, browser-instance-backed. The CLI one is local,
   text-first, file-import/export-backed. I treat them separately below.

We're building neither of those. We're building a **local, in-process
MCP server** (extending the existing DOCX one in
`packages/agent/src/mcp.ts`) that talks to a `XlsxAgent` over a typed
command bus with a per-mutation diff. So the question this doc has to
answer is: _what concepts from each side should we keep, drop, or
invert?_

---

## 1. MCP tool surface (univer-mcp side)

### 1.1 What the README actually claims

The hosted MCP server exposes tools backed by Univer's plugin system,
covering at least:

- spreadsheet creation and manipulation
- formulas, conditional formatting, data validation
- charts (WIP)
- pivot tables (WIP)
- collaborative real-time editing (WIP)

It explicitly states that _some_ tools return **images** for "better
understanding", and that **plain-text mode is currently experimental and
not yet supported**. The recommended host model is therefore
multimodal.

That last constraint is striking. It implies their core read tools
(`inspect`, range-read, possibly `search`) ship a **rendered raster of
the workbook** alongside any structured payload — leveraging the fact
that there's a real Univer renderer running in the bound instance. The
agent reads the picture. Cells are not just data, they are pixels.

### 1.2 What we can infer about tool shape

Without source we can only infer: HTTP (streamable) transport with
`Bearer` API key, session affinity via `?univer_session_id=…` query
parameter (must match the value the running Univer instance was
started with), no file handle (the bound instance "is" the workbook),
tool calls are RPCs proxied by `mcp-bridge` into Univer's command
system, not in-process mutations.

### 1.3 What this means for `xlsx_*` in our build

We're the inverse design — in-process, stdio, file-handle-based,
text-only, with a typed command bus that owns the diff. We should not
copy the hosted shape, and given how thin the docs surface is we
couldn't anyway. The one thing worth absorbing: the **separation of
read from mutate**. Univer's read surface clearly leans on a renderer;
ours has to lean on structured projections (markdown, JSON, CSV/TSV)
because we have no pixels to ship.

---

## 2. The `agent-sheet` skill — the real reference

`agent-sheet` is the local CLI the `dream-num/skills` repo wraps. It's
the closest thing to "what an LLM-driven sheet agent actually does on
disk" and almost every observation in §3–§7 comes from this skill, not
from `univer-mcp`. The skill is shaped as a routing tree: a top-level
`SKILL.md`, a tree of lane playbooks (preflight, read-analyze, verify,
write-safe, file-lifecycle, script-fallback), references
(command-selection matrix, shell patterns, JS API, gotchas), and
end-to-end examples.

The implied CLI surface (from playbooks + matrix + examples):

| Group   | Commands                                                                   |
| ------- | -------------------------------------------------------------------------- |
| Init    | `init`, `file list`, `file create`, `file import`, `file open`             |
| Info    | `file info`, `inspect workbook`, `inspect sheet`, `inspect range`          |
| Read    | `read range`, `read search`                                                |
| Write   | `write cells`, `write range`, `write table`, `write fill`                  |
| Sheet   | `sheet list`, `sheet create`, `sheet rename`, `sheet copy`, `sheet delete` |
| Persist | `persist`, `file export`                                                   |
| Escape  | `script js`                                                                |

That decomposition matches our intuition for the `xlsx_*` MCP shape
almost 1:1, with a few important shape differences below.

---

## 3. Range serialization

### 3.1 Univer's flavor

`agent-sheet read range` ships with three knobs:

- **`--type {value|rawValue|formula}`** — _display value_ (formatted
  string), _raw value_ (typed underlying datum: numbers as numbers,
  dates as 1900-based serials, text as text), or _formula source_.
- **`--format {csv|tsv|json|jsonl}`** — wire shape.
- **`--to-stdout` / `--to-file --output <path>`** — sink.

This is a really clean orthogonal split: the _type_ axis decides what's
in each cell, the _format_ axis decides how the matrix is encoded, the
_sink_ axis decides where it goes. The `references/shell-patterns.md`
file makes it obvious why three encodings exist:

- **TSV** for `awk` / `sed` next-step pipelines
- **CSV** for human review and Python `csv` consumers
- **JSON / JSONL** for structured tools and search hit streams

There is no "cell-by-cell objects with refs" mode in the published
surface. The 2-D matrix is the contract. A cell's address is implicit
from its row/column position inside the requested A1 range.

### 3.2 Pagination story

There isn't one — at least not as cursors. The discipline is:

1. The caller supplies an explicit, bounded A1 range
   (e.g. `Claims!A1:H200000`), so size is the caller's problem.
2. For very large reads, the playbooks insist on `--to-file` and a
   staged artifact rather than inline output.
3. Local export hard-fails when the snapshot JSON exceeds **100 MB**.
   That's their only structural ceiling.

For us with a 50k-row workbook the lesson is: **bound the range at the
tool boundary, don't paginate inside the tool.** The agent always knows
how big a slice it wants. If a future tool needs to iterate, it can
chunk by row offset because it already speaks A1. We don't need to
invent a cursor protocol.

### 3.3 What we should adopt

The `(type × format × sink)` orthogonal split, as three arguments on
`xlsx_read_range`, is good design — clone the _shape_ (concept, not
code). 2-D array is the canonical wire form (CSV/TSV is overkill since
the MCP SDK already serializes JSON, but at minimum emit a **flat 2-D
array**, not a `{ref → value}` map — address implicit from position).
For huge ranges, refuse early on a configurable cell-count cap (default
100k) and tell the agent to narrow the range — don't invent pagination
tokens. Mirrors Univer's "100 MB hard fail" at MCP-tool granularity.

### 3.4 What we should differ on

`xlsx_search` results need a **sparse** projection
(`{sheet, ref, value, formula?}` per hit) because matches are
positional, not rectangular. Univer uses JSONL for this; in MCP a JSON
array of records is more natural than newline-delimited JSON.

---

## 4. Formula handling in tools

### 4.1 What `agent-sheet` does

Two layers:

- **Built-in writes:** `write cells` accepts a JSON map (`{"Sheet!A1": v}`),
  but the public surface does not document a separate "write formula"
  command. Inferred from the playbooks: a formula string starting with
  `=` is treated as a formula (this matches Univer's JS API where
  `setValue('=…')` is an alias for `setFormula('…')`).
- **Reads round-trip three views:**
  1. `read range` (default): displayed values
  2. `read range --type rawValue`: typed underlying values
  3. `read range --type formula`: formula source strings
  4. `inspect range`: structural summary including "formula groups"
- **Async calculation barrier (JS layer):** `setFormula(...)` does not
  return the value; the caller must `await
univerAPI.getFormula().onCalculationResultApplied()` before reading.
  500 ms quiet-then-resolve, 30 s global timeout. The verification
  playbook explicitly tells agents to verify formulas with **both**
  `read --type formula` _and_ normal `read range` so the formula and the
  computed value are checked separately.

### 4.2 What goes in, what comes back

Going **in**: a string, with a leading `=` deciding "formula vs literal".
There is no rich payload — no dependency hint, no expected result, no
units. The formula is the formula.

Coming **back from a read**: the formula string only. No cached value
is bundled with it; if you want the cached value you call `read range`
again without `--type formula`. This is _deliberately_ split, presumably
because the formula vs cached-value distinction is what the verification
discipline is built around.

### 4.3 What we should adopt

- **Two reads, not one fused payload.** `xlsx_read_range` should accept
  `value_type: "display" | "raw" | "formula"`. Don't return a fused
  `{formula, value, type}` per cell — that bloats the payload for the
  90 % of cells that aren't formulas. Let the agent ask twice for the
  rare diff case.
- **Single "string in" surface for writes.** A leading `=` is the
  signal. _But_: because we have a typed command bus and a per-mutation
  diff, we should still expose `xlsx_set_formula` as a separate tool
  from `xlsx_set_cell` for **agent ergonomics** (the LLM can disambiguate
  intent and we can validate the formula at parse time). The handler
  underneath can collapse to one command if we want. See §8 for the
  proposed shape.
- **Cached value freshness.** Our formula engine is sync (per
  `spec/xlsx/formula-engine.md` plans), so we _don't_ need an
  `awaitCalculationResultApplied()` analog. We should still document
  that `xlsx_set_formula` returns the cached value in the diff so that
  agents don't have to follow up with a `xlsx_read_range`. This is a
  real differentiator from the Univer flow.

### 4.4 What we should differ on

- Univer's formula round-trip is documented to be _async_ across a
  bridge to a renderer. Ours is in-process and the calculation graph
  is owned by us. We can therefore promise: **after `xlsx_set_formula`
  returns, the cached value in the diff is current and consistent with
  every dependent cell that was also recomputed.** That's a stronger
  guarantee than Univer's MCP can make.

---

## 5. Skill definitions — what the `dream-num/skills` repo really is

### 5.1 Form factor

These are **Cursor / Claude Code / Codex skills** — not prompt
templates, not tool aliases, not fine-tuning datasets. Each skill is a
folder: a top-level `SKILL.md` with front-matter + routing index, a
tree of lane-shaped playbooks, declarative references, end-to-end
examples, and optional helper scripts. The agent loads `SKILL.md` then
reads only the linked file relevant to the current task — the routing
tables keep token cost low.

### 5.2 What the routing tables actually look like

Two patterns dominate: **intent → command** matrices ("if you want X,
prefer Y, because Z" — `command-selection-matrix.md` is the canonical
example, ~9 intents) and **state → next-doc** routers ("if your task
looks like X, read `playbooks/Y.md` next" — the `SKILL.md` "Task
routing" table). Result: a hyperlinked, lane-based decision tree
rather than one giant prompt. Each leaf doc is ~50–250 lines,
self-contained, and ends with a "stop / escalate" block.

### 5.3 What we should adopt for our `spec/xlsx/agent-commands.md` and a sibling `xlsx-skill/`

- Our existing `spec/docx/agent-commands.md` is a **command reference**,
  not a skill. It's exhaustive, payload-typed, and explains OOXML
  impact per command. Keep it as a reference doc.
- We should ship a **separate `xlsx-skill/`** (or `agent-skills/xlsx/`)
  in the spec that **mirrors the dream-num lane shape**: preflight,
  read-analyze, write-safe, verify, file-lifecycle, escape-hatch
  (`xlsx_apply` with arbitrary command), gotchas. The `xlsx_*` MCP
  surface in §8 below is what those playbooks would route into.
- The "intent → command" matrix is the single highest-leverage
  artifact. Build it from the §8 tool list. It belongs at the top of
  the skill, before any playbook.
- **Skip** the `references/js-api-minimal.md` analog — we don't expose
  a JS escape hatch into our renderer, we expose `xlsx_apply` (typed
  command bus) which already covers the "I need to do something not
  pre-baked" case. This is actually a simpler story than Univer's.
- **Adopt** the verification-first culture wholesale. The
  `playbooks/15-verify.md` rule "do not stop at _the command succeeded_,
  prove the workbook state that matters" should be hard-coded into our
  skill's introduction. Our per-mutation diff makes this dramatically
  easier than Univer's read-back-and-compare.

---

## 6. Auth / state — how Univer scopes a session to a workbook

### 6.1 Hosted MCP

Bearer API key in `Authorization` (account-scoped). Workbook scope is
`?univer_session_id=<id>`, which must match the id the running Univer
instance was started with — distinct ids isolate instances. No file
handle anywhere; the workbook is whatever the bound instance has open.
Single workbook per session; multi-document = multiple sessions.

### 6.2 Local CLI (`agent-sheet`)

Workspace-rooted like a Git repo (`init` creates one; `file list`
detects whether the cwd already lives in one). Each workbook has an
opaque `entryId` — every read/write/inspect command takes
`--entry-id`. Multi-document = multiple entries in the same workspace.
`file import <path>` returns a fresh `entryId`; the original file is
not mutated in place; export is explicit. The skill is emphatic:
**trust the `entryId` from `file import` even if a later `file info`
returns `unitId: null`** — there's a documented metadata/runtime split.

### 6.3 Where we are

`packages/agent/src/mcp.ts` already has the right shape: a
`randomUUID()` `handle` per loaded workbook (compare: `entryId`),
process-lifetime only, a `sessionPaths` map so `docx_save` can default
to the original path, and a `lookupAgent(handle)` funnel that returns a
helpful "call `docx_load` first" error on miss. **Clone this exact
shape for XLSX**: `xlsx_load(path) → handle`,
`xlsx_save(handle, out_path?) → { wrote, bytes, revision }`, every
other tool takes `handle: string` first, multi-document = multiple
handles. The one departure from Univer: **default `xlsx_save` to the
load path** (matches DOCX, matches every editor LLM users have seen);
trades a tiny safety property for big ergonomics.

---

## 7. Pending / staged mutations and observability

### 7.1 What Univer has

`agent-sheet` itself does **not** expose an "agent → pending →
human-approval" flow at the per-tool-call level. Successful writes are
applied; `persist` is a dirty-state boundary; that's it.

The `sheet-git` sister skill is where pending/approval lives, and it's
_much_ heavier than ours: `stage` + `commit` + `proposal create` build
a Git-shaped local repo on top of the workbook; `push review` publishes
to a hosted browser review session for human approval; `proposal
comments` returns machine-readable JSON review packets with continuity
semantics (current-revision threads plus carried-forward unresolved
ones); `push origin` materializes the approved proposal; `fetch/pull
origin` handle collaborative drift; `blame --cell 'Sheet1!A1'` does
per-cell blame. Essentially Git on a workbook plus a hosted PR review
surface plus a replay engine — beautiful and **vastly more product
than we want to ship**.

### 7.2 What we have

`docx_apply_command` already exposes `auto_approve: boolean`. When
false, the mutation lands in the `agent`-source pending queue and a
downstream reviewer calls `approveMutation` / `rejectMutation`. Default
true keeps the agent fast. Per-mutation diffs come from the handlers,
not a separate `diff` step.

That's a _much_ simpler design than `sheet-git`'s and the right one for
our scope. We don't need a Git-shaped repo to express "pending review",
don't need a hosted browser approval surface (the IDE is the surface),
don't need replay/rebase. The differentiator to keep loud in our skill
docs: **every XLSX tool call can be staged, every staged mutation
surfaces a structured diff, every diff is reviewable in the same
session without a hosted detour.** Univer needs three CLIs and a
website to do that; we need one MCP tool flag.

### 7.3 Observability — structured diff after a tool call

`agent-sheet` does not return a structured diff by default — `inspect
range`, `read range --type formula`, and `inspect workbook` are the
read-back primitives the verify playbook leans on. Pattern: "write,
then read the same range back and eyeball-compare". `sheet-git` adds
`diff` and `blame` against committed history, but those are post-
commit, not per-call.

Compare to ours: every `Mutation` carries the diff inline, and
`docx_diff` already supports two-handle and handle-vs-disk modes.
Implications for XLSX:

- Every `xlsx_*` write tool MUST return `{ revision, diff }` so the
  agent never has to do a follow-up read just to confirm what changed.
- `xlsx_diff` should mirror `docx_diff` — `{before, after}` two
  handles, or `{handle, against: 'disk'}`.
- Do **not** clone the `sheet-git` surface. If a future product needs
  review/PR semantics, build it on top of our diff stream, don't fork a
  parallel CLI.

---

## 8. Recommended `xlsx_*` MCP tool surface

The list below is sized to match `docx_*` (single MCP server, one tool
per intent, all sharing a `handle: string`). Inputs are described as
field bullets, not schemas, to stay code-light. Outputs always include
`{ revision: number }` for writes and `{ diff: XlsxDiff }` for any tool
that mutates state.

### 8.1 Session lifecycle

These two are not in the user's enumerated list, but they're load-bearing
and we already have the DOCX equivalents — including them keeps the
table honest.

#### `xlsx_load`

- In: `path` (absolute or workspace-relative `.xlsx`)
- Out: `{ handle, path, summary }` where `summary` is the `xlsx_inspect`
  payload (sheets, dimensions, defined names count, comments count,
  parts list with hashes for byte-preservation tracking).

#### `xlsx_save`

- In: `handle`, optional `out_path` (defaults to load path).
- Side effect: auto-approves any pending agent mutations, then exports.
- Out: `{ wrote, bytes, revision }`.

### 8.2 Inspection and reading

#### `xlsx_inspect`

- In: `handle`, optional `sheet` (narrows scope to one sheet).
- Out: workbook-level: `{ sheets: [{ name, id, dim: 'A1:Z200', rowCount,
colCount, hidden }], definedNames, comments, parts }`. Sheet-level:
  `{ name, dim, frozenRows, frozenCols, mergedRanges, formulaCount }`.
- Mirrors Univer's `inspect workbook|sheet|range` collapsed into one
  tool. We collapse because MCP tool counts matter for context cost.

#### `xlsx_read_range`

- In: `handle`, `range` (string, A1 with sheet prefix — see §9),
  optional `value_type: "display" | "raw" | "formula"` (default
  `"display"`), optional `format: "json" | "csv" | "tsv"` (default
  `"json"`), optional `max_cells` (refuses with hint when exceeded;
  default 100000).
- Out: 2-D array under `cells`, plus `{ sheet, ref, rows, cols,
truncated: false }`. CSV/TSV variants return `text` instead.
- Address is implicit by row/col position. No per-cell ref/object
  inflation.

#### `xlsx_search`

- In: `handle`, `query` (string), optional `sheet` (limit), optional
  `case_sensitive`, optional `regex`, optional `match_entire_cell`,
  optional `value_type` (search displayed text vs. formula source), optional
  `limit` (default 500).
- Out: `{ matches: [{ sheet, ref, row, col, value, formula? }],
truncated }`.
- Different shape from `xlsx_read_range` because matches are sparse and
  positional. Mirrors `agent-sheet read search` in spirit.

### 8.3 Cell and range writes

#### `xlsx_set_cell`

- In: `handle`, `at` (`{ sheet, ref }` or `'Sheet1!A1'`), `value`
  (string|number|boolean|null), optional `format` (number-format
  string).
- Behavior: literal value only. Leading `=` is **rejected** with a
  hint to call `xlsx_set_formula`. Keeps tool intent clear for the LLM.
- Out: `{ revision, diff }`.

#### `xlsx_set_formula`

- In: `handle`, `at` (single cell), `formula` (string, with or without
  leading `=` — we normalize), optional `array_range` for dynamic-array
  formulas (`A1:C3`).
- Out: `{ revision, diff, cached_value, recalculated_cells: number }`.
- The `cached_value` and `recalculated_cells` fields are what saves the
  agent a follow-up `xlsx_read_range`. This is the §4 differentiator
  from Univer's async surface — promise the agent it's fresh.

#### `xlsx_set_range`

- In: `handle`, `range` (rectangle, `'Sheet1!A1:C10'`), `cells` (2-D
  array, **must** match range dimensions), optional `value_type`
  (`"value" | "formula" | "auto"`; `"auto"` infers per cell from leading
  `=`), optional `format` (number-format applied uniformly).
- Out: `{ revision, diff, written_cells }`.
- This is Univer's `write range` semantics: **the rectangle is the
  contract**. Mismatched dims must be rejected at the tool boundary
  with a clear message — Univer's gotchas explicitly call this out.
- We **deliberately do not** ship `write table` (Univer's A1-anchored
  table-with-header semantics). It's a fine UX but it overlaps
  `xlsx_set_range` and `xlsx_add_sheet` and we're trying to keep the
  surface small. Document the recipe in a playbook instead.

#### `xlsx_format_range`

- In: `handle`, `range`, `format` (object: `font`, `fill`, `border`,
  `align`, `numberFormat`, `wrapText` — leave-unchanged when omitted,
  matching `TextFormatPayload` discipline from `spec/docx/agent-commands.md`).
- Out: `{ revision, diff, affected_cells }`.
- This is the surface Univer has to drop into `script js` for. We get
  it as a first-class tool because our model is structurally typed.

### 8.4 Structural mutations

#### `xlsx_insert_row` / `xlsx_insert_column`

- In: `handle`, `sheet`, `at` (1-based row/col index, matching A1
  intuition), `count` (default 1).
- Out: `{ revision, diff, shifted_cells }`.

#### `xlsx_delete_row` / `xlsx_delete_column`

- In: `handle`, `sheet`, `at` (1-based), `count` (default 1).
- Out: `{ revision, diff, deleted_cells, shifted_cells }`.
- Unlike Univer's JS-API gotcha ("delete backwards in a loop"), our
  tool takes `(at, count)` and handles the index math internally. The
  agent should never have to think about iteration order.

#### `xlsx_merge` / `xlsx_unmerge`

- In: `handle`, `range`, optional `force: boolean` (for `xlsx_merge`,
  unmerges any overlapping merges first; default false → reject).
- Out: `{ revision, diff }`.
- Univer's API has the same `isForceMerge` knob. Worth keeping.

#### `xlsx_add_sheet` / `xlsx_rename_sheet`

- `xlsx_add_sheet`: In: `handle`, `name`, optional `rows`, optional
  `cols`, optional `position` (insert index). Out: `{ revision, diff,
sheet_id, name }`.
- `xlsx_rename_sheet`: In: `handle`, `from`, `to`. Out: `{ revision,
diff }`. Reject on collision.
- We deliberately omit `delete_sheet` from Phase 8. Sheet deletion is
  unrecoverable in-session and worth a separate hardening pass.

#### `xlsx_add_comment`

- In: `handle`, `at` (single cell), `text`, optional `author`.
- Out: `{ revision, diff, comment_id }`.
- Mirrors `docx:add-comment` shape from `spec/docx/agent-commands.md`.

### 8.5 Batch + diff

#### `xlsx_apply` (batch / escape hatch)

- In: `handle`, `commands: [{ type, payload }]` (typed `xlsx:*` command
  bus calls), `source: "agent" | "human" | "system"` (default `agent`),
  `agent_id` (default `officeai-mcp`), `auto_approve: boolean` (default
  true).
- Out: `{ revision, diff, mutations: [{ id, status, rejection? }] }`.
- This is the single tool that wraps our typed command bus directly,
  exactly mirroring `docx_apply_command` but plural. It's the
  intentional escape hatch for any `xlsx:*` command we ship that
  doesn't have a dedicated MCP tool. Univer's `script js` plays a
  similar role; ours is typed, theirs is `eval`.
- Pending semantics: when `auto_approve: false` and `source: "agent"`,
  mutations land in the queue. A downstream `xlsx_apply` from
  `source: "human"` (or a future `xlsx_approve` tool) settles them.
  This is exactly the DOCX shape today.

#### `xlsx_diff`

- In: either `{ before, after }` (two handles), or `{ handle, against:
"disk" }`.
- Out: structured diff: `{ sheets: [{ name, addedRows, deletedRows,
cells: [{ ref, before, after, kind: "value"|"formula"|"format" }] }],
definedNames: { added, removed, changed }, mergedRanges: { added,
removed } }`.
- Mirror of `docx_diff`. Keep the shape consistent across DOCX and
  XLSX so a single diff renderer in the IDE can handle both.

### 8.6 Tool-count budget

Total: 2 lifecycle + 3 inspection + 4 cell/range writes + 1 format + 4
row/col + 2 merge + 2 sheet + 1 comment + 1 apply + 1 diff = **21
tools**. Larger than DOCX's 7, but XLSX is a wider surface. Could
compress row+column `insert_*`/`delete_*` to two tools with a
`dimension` enum (Univer's `Enum.Dimension.ROWS` pattern); LLM
discoverability wins, keep four.

---

## 9. Range / cell address conventions

### 9.1 Univer's mix

- Built-in CLI (`agent-sheet`) uses **A1 with sheet prefix as a single
  string** for every `--range` argument: `Sheet1!A1:H80`,
  `'工作表1!A1:J3'` (must be quoted in shell because of `!`).
- Cell-only writes via `write cells` use `{ "Sheet!A1": value }` JSON —
  same `Sheet!A1` string as the key.
- The JS API (`script js`) supports **both**:
  - A1: `getRange('A1:B10')`, `getRange('A:A')`, `getRange('1:1')`
  - 0-based numeric coords: `getRange(row, col, rowCount, colCount)`
  - `setValues()` requires the range dimensions to match the data
    array — single-cell `getRange(0,0)` followed by `setValues([[…
multi-cell …]])` is a documented footgun.

R1C1 does not appear anywhere in the public surface. Neither does a
`{ sheet: 'Sheet1', ref: 'A1' }` split-object form.

### 9.2 What we should pick

**Single canonical form: A1 with sheet prefix as a string.**
`Sheet1!A1`, `Sheet1!A1:C10`, with the sheet half **required** at the
tool boundary (no "active sheet" fallback — that's a stateful landmine
in a stateless tool). Internally we parse once into
`{ sheet, start, end }` and pass that to handlers.

Reasons: (a) **one way to do it** — agent never has to pick between
`'Sheet1!A1'` and `{ sheet, ref }`, two forms doubles ambiguity and
testing; (b) **matches every external reference an LLM has seen**
(Excel, Google Sheets, every formula textbook) — fighting the
convention buys nothing; (c) **R1C1 is a non-starter** — its only edge
is relative-vs-absolute formula writing, solved at the formula-string
level (`$A$1`); (d) **numeric coords stay internal** — `insert_row` /
`delete_column` intrinsically take numbers; we use **1-based** indices
in the public API to match A1 intuition, convert to 0-based at the
handler boundary. Univer exposes 0-based in its JS API and pays for it
with documented footguns; we should not.

### 9.3 Edge-case rules

- An `at` field for single-cell tools (`xlsx_set_cell`, `xlsx_set_formula`,
  `xlsx_add_comment`) accepts the same `'Sheet1!A1'` string. No
  `{ sheet, ref }` alternative.
- A `range` field for rectangle tools (`xlsx_read_range`,
  `xlsx_set_range`, `xlsx_format_range`, `xlsx_merge`) accepts
  `'Sheet1!A1:C10'`. Single-cell ranges (`'Sheet1!A1'`) are normalized
  to `'A1:A1'` internally.
- `xlsx_insert_row`/`xlsx_delete_row` take a `sheet` field plus a
  numeric `at` (because the row index is the address — a row ref like
  `1:1` doesn't have a "before this" semantic).
- Sheet names containing `!` are not supported. Document loudly.

---

## 10. Concrete deltas vs Univer

| Question                  | Univer hosted MCP   | Univer `agent-sheet` CLI     | Our `xlsx_*`                |
| ------------------------- | ------------------- | ---------------------------- | --------------------------- |
| Transport                 | HTTP + bearer       | shell + `--entry-id`         | stdio + handle              |
| Workbook scope            | URL session id      | workspace + entryId          | in-process handle           |
| File handle               | none                | yes, on `import`             | yes, on `xlsx_load`         |
| Default output            | multimodal (image)  | display values               | structured JSON             |
| Range encoding            | unknown             | 2-D, CSV/TSV/JSON/JSONL      | 2-D JSON (CSV/TSV optional) |
| Pagination                | unknown             | bounded A1 + 100 MB cap      | bounded A1 + 100k-cell cap  |
| Formula round-trip        | unknown, async      | three-axis value/raw/formula | three-axis (same)           |
| Cached value freshness    | unknown             | async, manual await          | sync, returned in diff      |
| Format / borders / freeze | first-class implied | `script js` only             | first-class tool            |
| Approval flow             | none documented     | none in CLI; `sheet-git` Git | per-tool `auto_approve`     |
| Diff after write          | none returned       | none; verify by re-read      | inline in every write       |
| Address                   | unknown             | A1 + `!` sheet prefix        | A1 + `!`, required, 1-based |

---

## 11. Open questions

Flagged but not locked here: (1) per-`xlsx:*` JSON-schema validation at
the tool boundary vs trusting the typed handler — DOCX trusts the
handler, probably keep that; (2) whether to ship `xlsx_get_text` (whole
workbook → markdown) as a parallel to `docx_get_text` — likely yes,
useful for "summarize this workbook" prompts but not in the enumerated
list; (3) comment threading — `xlsx_add_comment` is included, but
resolve / reply / delete (parallels of DOCX `docx:resolve-comment`
etc.) are not — defer or ship in Phase 8 consciously; (4) conditional
formatting and data validation — Univer ships these, we assume out of
scope per `spec/xlsx/feature-scope.md`; (5) `xlsx_apply` vs dedicated
tools — every dedicated write tool is also reachable through
`xlsx_apply` with the right command type, but the dedicated tools are
for ergonomics/discoverability and `xlsx_apply` is for arbitrary types
and batch atomicity. Keep both.

---

## 12. Bottom line

- Univer's hosted MCP server (`dream-num/univer-mcp`) is documentation
  only; the actual tool surface is opaque, multimodal, and built around
  a remote browser-resident instance. We can lift architectural intent
  but no concrete shape.
- Univer's local CLI skill (`dream-num/skills/agent-sheet`) is the high-
  signal reference. It's not MCP, but its command decomposition,
  read-axis split (`value` × `raw` × `formula`), explicit `entryId`
  handle model, A1-with-sheet-prefix addressing, write-then-verify
  discipline, and "smallest matching command" routing matrix are all
  directly applicable.
- Their separate `sheet-git` skill is a Git-shaped collaborative
  approval surface — vastly heavier than our per-mutation `auto_approve`
  flag. Our flow is a real differentiator and we should advertise it.
- The proposed 21-tool `xlsx_*` MCP surface in §8 mirrors our existing
  `docx_*` surface for handle/save/inspect/diff/apply, picks up the
  `(value × raw × formula)` × `(json × csv × tsv)` orthogonal split
  from Univer for `xlsx_read_range`, exposes formatting / merging /
  row+col ops as first-class tools (Univer's biggest pain point —
  `script js` only — becomes our biggest win), and standardizes on
  A1-with-required-sheet-prefix addressing with 1-based numeric
  indices to match A1 intuition.
