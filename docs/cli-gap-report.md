# `office-agent` CLI Gap Report

This report captures what's missing from the `office-agent` CLI surface
when used as the **sole** interface for an AI to do CRUD on DOCX, XLSX,
and PPTX. It's grounded in three end-to-end smoke runs:

- [`scratch/cli-smoke/contract.docx`](../scratch/cli-smoke/contract.docx) — employment contract built via [`scratch/cli-smoke/build-contract.sh`](../scratch/cli-smoke/build-contract.sh)
- [`scratch/cli-smoke/salaries.xlsx`](../scratch/cli-smoke/salaries.xlsx) — 5-employee salary calculator built via [`scratch/cli-smoke/build-salaries.sh`](../scratch/cli-smoke/build-salaries.sh)
- [`scratch/cli-smoke/onboarding.pptx`](../scratch/cli-smoke/onboarding.pptx) — 4-slide onboarding deck built via [`scratch/cli-smoke/build-onboarding.sh`](../scratch/cli-smoke/build-onboarding.sh)

Gaps are ordered roughly by impact, not by format. Each gap includes
**Problem**, **Proposed fix**, and **Files to touch**.

---

## §G0 — `*-apply` corrupts documents on ≥3 commands (CRITICAL)

**Problem.** `oa docx apply -c file.json`, `oa xlsx apply-file -c file.json`,
and `oa pptx apply -c file.json` all share this pattern:

```ts
const muts = await agent.applyCommands(cmds);
agent.getPendingMutations().forEach((m) => agent.approveMutation(m.id));
```

`getPending()` in `CommandBus` returns the live `this.pending` array
(no copy). `approveMutation` calls `this.pending.splice(idx, 1)` and
then `recomputeWorking()` which **reassigns** `this.pending` to a new
array. The `forEach` is iterating the _old_ reference whose head was
spliced — so iteration indices skip every other element, leaving half
the mutations stuck in `pending`. Each `recomputeWorking()` call then
re-applies those stuck mutations against the evolving approved
snapshot, accumulating duplicates.

**Repro** (probe captured during the smoke run):

```json
[
  { "type": "docx:insert-text", "payload": { "at": { "paragraph": 0, "offset": 0 }, "text": "A" } },
  { "type": "docx:insert-text", "payload": { "at": { "paragraph": 1, "offset": 0 }, "text": "B" } },
  { "type": "docx:insert-text", "payload": { "at": { "paragraph": 2, "offset": 0 }, "text": "C" } },
  { "type": "docx:insert-text", "payload": { "at": { "paragraph": 3, "offset": 0 }, "text": "D" } }
]
```

against `fixtures/docx/synthetic/01-plain-paragraphs.docx` produces:

| paragraph | expected             | observed                 |
| --------- | -------------------- | ------------------------ |
| 0         | `APlain`             | `APlain`                 |
| 1         | `BFirst body para…`  | `BBBFirst body para…` ❌ |
| 2         | `CSecond body para…` | `CSecond body para…`     |
| 3         | `DThird body para…`  | `DThird body para…`      |

**Proposed fix.** Snapshot the IDs before approving, e.g.

```ts
const ids = agent.getPendingMutations().map((m) => m.id);
for (const id of ids) agent.approveMutation(id);
```

The smoke runs work around this by issuing one mutation per CLI
invocation; that's slow and unidiomatic for batch agent use.

**Files to touch.**

- `packages/agent/src/cli.ts` (`docx apply` action ~L807-L824, `runWrite` ~L989-L991)
- `packages/agent/src/cli-xlsx.ts` (`xlsx apply-file` action ~L613-L632, `runXlsxWrite` ~L220-L240)
- `packages/agent/src/pptx-cli.ts` (`dispatchAndWrite` ~L1370-L1400)
- (Optional) `packages/core/src/commands/bus.ts` — make `getPending()` return a frozen copy so this footgun isn't reachable.

---

## §G1 — No `create` / blank-document constructor (cross-format)

**Problem.** Every write subcommand requires `--file <existing>` and
none of the agent packages expose an `Agent.empty()` constructor. The
smoke runs had to bootstrap each scenario by `cp`ing the smallest
synthetic fixture (`01-plain-paragraphs.docx`,
`01-single-sheet-numbers.xlsx`, `01-blank.pptx`) into `scratch/`. An AI
agent without filesystem access to the fixture tree literally cannot
start a new document.

**Proposed fix.**

- Add `DocxAgent.empty(opts?)`, `XlsxAgent.empty(opts?)`,
  `PptxAgent.empty(opts?)` returning a minimal valid OOXML container
  (1 paragraph / 1 sheet "Sheet1" / 1 blank slide on layout 1).
- Surface as `oa docx create --out <path>`, `oa xlsx create --out <path>`,
  `oa pptx create --out <path>` (optionally `--from-template <path>`).

**Files to touch.**

- `packages/docx/src/agent/agent.ts` — add `static empty()`.
- `packages/xlsx/src/agent/agent.ts` — add `static empty()`.
- `packages/pptx/src/agent/agent.ts` — add `static empty()`.
- `packages/agent/src/cli.ts`, `cli-xlsx.ts`, `pptx-cli.ts` — register
  `create` subcommand on each group.

---

## §G2 — DOCX: missing `insert-paragraph`, `delete-range`, `replace-text`

**Problem.** `docx:insert-paragraph` and `docx:delete-range` are
registered bus commands
([`packages/docx/src/commands/registry.ts`](../packages/docx/src/commands/registry.ts))
with full handler implementations
([`packages/docx/src/commands/insert-paragraph.ts`](../packages/docx/src/commands/insert-paragraph.ts),
[`packages/docx/src/commands/delete-range.ts`](../packages/docx/src/commands/delete-range.ts))
but neither is exposed as a typed `oa docx …` subcommand. The contract
build had to drop into `oa docx apply` with hand-rolled JSON for
**every** new paragraph and **every** text replacement. There is also
no `oa docx replace-text` convenience — replacing existing text needs
delete-range + insert-text, which is two round-trips per edit.

**Proposed fix.**

- `oa docx insert-paragraph --file … --at <selector> [--style <styleId>]`.
- `oa docx delete-range --file … --range <selector>` (reuse the
  `paragraph:0/text:0..5` selector grammar already present elsewhere).
- `oa docx replace-text --file … --paragraph-id <id> --text <text>`
  that compiles to `delete-range` over the entire paragraph plus
  `insert-text` at offset 0.

**Files to touch.**

- `packages/agent/src/cli.ts` — add three new `docx.command(...)` blocks
  near the existing `write` / `style` registrations (~L240-L320). Reuse
  `runWrite` so they auto-approve.

---

## §G3 — DOCX: `inspect`/`read` cannot expose table ids or run/text offsets

**Problem.** Two related discovery gaps surfaced:

1. **Table ids are unreachable.** `oa docx insert-table` mints a fresh
   table NodeId at runtime (`mintNodeId()`), but neither
   `oa docx insert-table`'s output nor `oa docx inspect` nor
   `oa docx read --format json` exposes table ids. `oa docx set-cell-text`
   strictly requires `--table-id`. The contract scenario therefore
   could not populate its compensation table cells via the CLI.
   See the empty `|  |  |` rows in
   [`scratch/cli-smoke/contract.docx`](../scratch/cli-smoke/contract.docx)'s
   `## 3. Compensation` section.
2. **Run / text offsets are opaque.** `oa docx read --format json`
   returns flat paragraph text; selectors like
   `paragraph:0/run:0/text:0` require knowing the run/text split,
   which is invisible without reading source XML.

**Proposed fix.**

- Make `runWrite` echo any newly minted node ids in its JSON envelope
  for `docx:insert-table`, `docx:insert-image`, `docx:insert-row`,
  `docx:insert-column` (return `mutation.diff.changes[].nodeId`). Best
  achieved by returning the post-mutation snapshot's last-touched node.
- Add `--with-tables` to `docx read --format json` so tables become
  first-class entries (id, rows, cols, per-cell text).
- Add `--with-runs` to `docx inspect` (or `docx read`) printing per-run
  text offsets so selectors are addressable without reading XML.

**Files to touch.**

- `packages/agent/src/cli.ts` — `runWrite` ~L984-L998 and the
  `inspect`/`read` actions ~L60-L160.
- `packages/docx/src/projection/markdown.ts` (or sibling) — extend the
  JSON projection to include tables.

---

## §G4 — XLSX: missing typed subcommands & `delete-sheet` handler

**Problem.** Several XLSX bus commands have no typed CLI surface:

| Bus command             | Handler at                                                    | CLI exposure                                                             |
| ----------------------- | ------------------------------------------------------------- | ------------------------------------------------------------------------ |
| `xlsx:set-column-width` | [`registry.ts:14`](../packages/xlsx/src/commands/registry.ts) | none — must use `oa xlsx apply --type xlsx:set-column-width --payload …` |
| `xlsx:set-row-height`   | [`registry.ts:16`](../packages/xlsx/src/commands/registry.ts) | none — same workaround                                                   |
| `xlsx:delete-sheet`     | **not registered** at all                                     | none                                                                     |

The salary build had to widen column A via `oa xlsx apply --type
xlsx:set-column-width --payload '{"sheetName":"Salaries","column":1,"width":18}'`
which is the kind of friction the typed surface is meant to remove.

There's also no fill-down helper: 5 employees needed 10 separate
`set-formula` calls (`D2..D6`, `E2..E6`). Even though `set-range`
accepts a 2-D matrix of literals, it doesn't accept formulas, so an
AI must emit one CLI call per row.

Finally, `clear-range` would be the natural counterpart to `set-range`
for wiping a region — currently you have to set every cell to `""`.

**Proposed fix.**

- Add typed `oa xlsx set-column-width --sheet … --column <n> --width <pts>`
  and `oa xlsx set-row-height --sheet … --row <n> --height <pts>`.
- Implement and register `xlsx:delete-sheet` in the bus, then expose
  `oa xlsx delete-sheet --sheet <name>`.
- Add `oa xlsx fill-formula --sheet … --range D2:D6 --formula "=B{row}*C{row}/100"`
  expanding `{row}` / `{col}` per cell in **one** round-trip.
- Add `oa xlsx clear-range --sheet … --range A1:E5`.

**Files to touch.**

- `packages/xlsx/src/commands/` — new `delete-sheet.ts` handler;
  register in `registry.ts`.
- `packages/xlsx/src/commands/payloads.ts` — `DeleteSheetPayload`.
- `packages/agent/src/cli-xlsx.ts` — add four new `xlsx.command(...)`
  blocks, plus the `fill-formula` expander helper that compiles N
  `xlsx:set-formula` payloads in a single dispatch.

---

## §G5 — XLSX: `--value` requires JSON-encoded literals

**Problem.** `oa xlsx set-cell --ref A7 --value Total` is rejected
because `Total` is not valid JSON. Strings have to be passed as
`'"Total"'`, with shell-aware quoting. For LLM-driven tool calls this
is a sharp edge — every string value needs an extra quoting round.

**Proposed fix.**

- Accept bare strings/numbers/booleans for `--value`. Heuristic:
  try `JSON.parse(raw)` first; on `SyntaxError` fall back to
  `String(raw)`. Add a `--value-raw <json>` escape hatch for the few
  cases that need typed nulls or error literals.
- Same treatment for any future `xlsx fill-formula --formula <expr>`
  if it grows a `--placeholder-value` flag.

**Files to touch.**

- `packages/agent/src/cli-xlsx.ts` — `set-cell` action (~L300-L340)
  and the `apply` shim that accepts `--payload <json>` (~L580).

---

## §G6 — PPTX: EMU coordinates, role-blind shapes, env-var determinism

**Problem.** Three PPTX papercuts hit the onboarding build:

1. **EMU only.** Every `--x/--y/--width/--height` expects English
   Metric Units (1" = 914400 EMU). The build script hand-converts
   inches → EMU on every line, e.g. `914400` for 1″. LLMs must do
   unit math or trust that `914400 ≈ 1"`.
2. **`pptx add-slide` returns no shape ids.** When `--layout` clones a
   layout's title/body placeholders into the new slide, the freshly
   minted shape NodeIds are not in the JSON envelope. To target them
   for `set-text`, the AI must re-run `pptx inspect`. The smoke run
   sidestepped this by `add-text-box`'ing freestanding shapes for
   every title — losing layout-driven theming.
3. **`OFFICEAI_DETERMINISTIC_IDS=1` is an env var, not a CLI flag.** It
   has to be re-exported on every shell prefix. The smoke `build-onboarding.sh`
   prepends `OFFICEAI_DETERMINISTIC_IDS=1` to a single string then
   `eval`s — fragile and easy to forget.

**Proposed fix.**

- Promote determinism to `--deterministic-ids` on the `pptx` parent
  command (or on every write subcommand). Internally still flip the
  same env-gate or pass through to `loadAgent`.
- Add `--unit emu|in|cm|pt` (default `emu`) to every `--x/--y/--width/
--height`-bearing PPTX subcommand. Convert in the action.
- Have `pptx add-slide` echo the new slide's shape ids in its JSON
  output (`{ wrote, mutations: [...], slide: { index, shapes: [{id, role, type}] } }`).
- Add `oa pptx set-title --slide <n> --text <text>` and
  `oa pptx set-body --slide <n> --text <text>` that resolve the
  layout's title/body placeholder by role and dispatch
  `pptx:set-text-content` against it.

**Files to touch.**

- `packages/agent/src/pptx-cli.ts` — root command options (~L80-L120),
  `add-slide` action (~L880-L920), and a new `set-title` / `set-body`
  block. Add a small `parseUnit(opts.unit, value)` helper.
- `packages/pptx/src/agent/agent.ts` — expose a `findPlaceholderByRole(slideIndex, role)`
  method so the CLI doesn't have to walk the snapshot.

---

## §G7 — Cross-format: `--from-stdin` for batch payloads

**Problem.** `oa <fmt> apply -c <path>` and `oa xlsx apply --payload <json>`
both require a temp file or a giant inline `--payload` string. An LLM
that wants to stream a multi-command batch ends up writing temp files
just to satisfy the CLI. That's an avoidable roundtrip when the agent
already has the JSON in memory.

**Proposed fix.**

- Accept `--from-stdin` (or `-c -`) on every `apply` / `apply-file`
  subcommand and on the typed write subcommands that take a `--payload`.
- When `--from-stdin` is set, read JSON from `process.stdin` (after
  draining; `await readStdin(io.stdin)`).

**Files to touch.**

- `packages/agent/src/cli.ts` — `docx apply` action.
- `packages/agent/src/cli-xlsx.ts` — `xlsx apply` and `xlsx apply-file`.
- `packages/agent/src/pptx-cli.ts` — `pptx apply`.

---

## Bonus findings (worth tracking but not full §G entries)

- **Rejected mutations exit 0.** When `oa docx footer --part word/footer1.xml`
  rejects with `unknown-target` (no footer part exists), the CLI still
  exits 0 and rewrites the input file with reshuffled NodeIds. Scripts
  cannot detect failures cheaply. Fix: exit non-zero whenever
  `mutation.status === "rejected"` (or any in `mutations[]` is
  rejected). Files: `packages/agent/src/cli.ts` `runWrite` (~L989),
  same for xlsx / pptx wrappers.
- **NodeIds reshuffle on every load.** Every `loadAgent` call mints
  fresh ids unless `OFFICEAI_DETERMINISTIC_IDS=1` is set, but only
  PPTX honours the env var. DOCX and XLSX always reshuffle, so any
  `--paragraph-id` / `--table-id` discovered by `read --format json`
  is stale by the time the next CLI call resolves it. Fix: respect
  `OFFICEAI_DETERMINISTIC_IDS` (and the proposed `--deterministic-ids`
  flag) in `loadDocxAgent` / `loadXlsxAgent` too. Files:
  `packages/agent/src/loader.ts` (or wherever `loadAgent` lives) and
  the per-format `parse*` entry points.
- **No header/footer creation.** `oa docx footer --part word/footer1.xml`
  fails when the part doesn't exist; there is no CLI command that
  _creates_ a header or footer part. Already noted as a deferred item
  in [`docs/session-summary.md`](session-summary.md), but worth
  bundling into the CLI surface once the engine ships.

---

## Suggested sequencing

1. **Fix §G0 first.** It's a correctness bug that silently produces
   wrong files; everything else can be worked around.
2. **§G1** unblocks the "AI-only" workflow.
3. **§G2 + §G3 + §G4** close the typed-subcommand parity gaps; do
   them together to keep the help text consistent.
4. **§G5 + §G6 + §G7** are ergonomics — schedule after the parity
   work so we can ship a single CLI-UX pass.
