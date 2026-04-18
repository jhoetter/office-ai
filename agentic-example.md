# OfficeAI Agentic Challenge — build three real Office docs with only the CLI

> **Hand this file to a fresh AI agent (Cursor, Claude Code, Codex, …)** and let it
> attempt the task end-to-end. Do **not** pre-flatten the CLI surface in your
> prompt — the goal is to see what the agent reaches for, what it gets stuck on,
> and which CLI ergonomics still feel sharp. It is a _naive challenge_: the agent
> must discover the tools on its own.

---

## Setup (one-time, ~30s)

You are working inside this monorepo. The only tool you may use to author content
is the `office-agent` CLI shipped at `packages/agent/`. Run these once:

```bash
pnpm install
pnpm --filter @officeai/agent build
export OA="node $(pwd)/packages/agent/dist/cli.js --deterministic-ids"
mkdir -p /tmp/officeai-demo
```

`--deterministic-ids` is recommended for chained scripted runs so that node IDs
returned by one command stay valid on the next invocation.

**Ground rules** (do not violate, otherwise the demo is meaningless):

1. **No hand-editing of OOXML.** No `unzip`, no `sed` against `.xml`, no
   external `python-docx` / `openpyxl` / `python-pptx`.
2. **Only `$OA` subcommands** for any document mutation. You may use `cat`,
   `jq`, `node`, etc. for plumbing.
3. **No GUI Office apps.** The whole point is to prove an AI can ship from a
   shell.
4. If something is missing from the CLI, **work around it once**, then write
   it down for the report instead of forking the codebase.

You can discover the surface with:

```bash
$OA --help
$OA docx --help
$OA xlsx --help
$OA pptx --help
$OA <group> <subcommand> --help
```

Every write subcommand emits a JSON envelope on stdout (use `--pretty` while
iterating). Exit code `0` = success, `2` = at least one mutation rejected (the
JSON tells you why), `64` = bad CLI invocation.

---

## The challenge

Produce **three** files at `/tmp/officeai-demo/`. Each one targets a realistic
authoring scenario; together they exercise the bulk of the CLI surface.

### 1. `employee-contract.docx` — Employment Agreement

Compose a short employment contract. Required content:

- A **bold, centered title** on the first line: `EMPLOYMENT AGREEMENT`.
- At least **3 distinct paragraphs**: one identifying the parties (use placeholder
  names, e.g. _Acme Corp_ and _Jane Doe_), one defining the role + start date,
  one covering compensation (USD figure of your choice).
- A **2×3 table** somewhere in the body with columns
  `Term | Value | Notes` and 2 data rows of your choice (e.g. _Probation_,
  _Vacation Days_).
- One **inline comment** anchored on the compensation paragraph with text
  _"Confirm with Finance before sending."_
- A trailing **signature paragraph** containing the literal text
  `Signed: ______________________   Date: __________`.

### 2. `salary-calc.xlsx` — Salary Calculator

Single-sheet workbook named `Payroll`:

- Row 1 is a **bold header**: `Name | Role | Base Salary | Bonus % | Total Comp`.
- Rows 2–6 hold **5 employees** of your choice (any plausible names, roles, and
  base salaries between 60k and 200k; bonus percentages between 5 and 25).
- Column **E (`Total Comp`) for each employee row must be a live formula**:
  `=C{row} * (1 + D{row}/100)`. No hard-coded totals.
- Row 8 contains the literal text `TOTAL` in column A and `=SUM(E2:E6)` in
  column E.
- Apply a **number format** of `"$#,##0"` to columns C and E.
- Set column widths so nothing is clipped (your call on the values).

### 3. `onboarding.pptx` — New-Hire Onboarding Deck

A **4-slide** presentation:

1. **Title slide** — text box(es) containing the title `Welcome to Acme` and a
   subtitle `Your first 30 days`.
2. **Agenda slide** — a list of at least 4 bullet points: _Day 1_, _Week 1_,
   _Month 1_, _Quarter 1_.
3. **Tooling slide** — three text boxes laid out as a row, each ~3 inches wide,
   labelled e.g. `Email`, `Slack`, `Source control`. Coordinate values must be
   expressed in inches (`--unit in`), not raw EMUs.
4. **Closing slide** — a single centered text box with `Questions? jane@acme`.

The first slide must be slide index `0`; reorder if needed.

---

## Acceptance criteria

A solution passes when **all** of the following are true. Verify them yourself
before declaring done; do not ask the human.

```bash
# All three files exist and parse cleanly.
test -s /tmp/officeai-demo/employee-contract.docx
test -s /tmp/officeai-demo/salary-calc.xlsx
test -s /tmp/officeai-demo/onboarding.pptx

# DOCX: title is the first paragraph, table is present, exactly one comment.
$OA docx inspect --file /tmp/officeai-demo/employee-contract.docx --pretty | head
$OA docx read    --file /tmp/officeai-demo/employee-contract.docx --format text \
  | head -20

# XLSX: header row, 5 employees, totals are formulas, SUM in row 8.
$OA xlsx read --file /tmp/officeai-demo/salary-calc.xlsx --sheet Payroll \
  --range A1:E8 --format json --pretty

# PPTX: 4 slides, slide 0 is the title slide.
$OA pptx inspect --file /tmp/officeai-demo/onboarding.pptx --pretty
$OA pptx read    --file /tmp/officeai-demo/onboarding.pptx --format text
```

For the XLSX verification, every row 2–6 in the JSON output must show a `formula`
field (not just a `value`) on cells in column E. For the PPTX verification, the
`slides` array must have length 4.

If you have any office suite installed locally and want a sanity check, opening
the files in Word / Excel / PowerPoint should not raise a "file is corrupt"
dialog. (This is optional — the JSON checks above are authoritative.)

---

## Hints (use only if stuck — try discovery first)

- The CLI is grouped: `$OA docx …`, `$OA xlsx …`, `$OA pptx …`. Each group has
  a `create --out <path>` subcommand for greenfield documents.
- Mutation commands accept either `--out <path>` (write to a new file) or no
  `--out` (rewrite in place).
- Most write subcommands have a typed equivalent (e.g. `docx write`,
  `xlsx set-cell`, `pptx set-text`) and a generic `apply` escape hatch that
  takes JSON. Use `apply --from-stdin` to pipe a batch of commands.
- `docx insert-text` / `docx write` need a **position selector** like
  `paragraph:0/run:0/text:0`. `$OA docx read --format json --pretty` and
  `$OA docx inspect --with-runs --pretty` show you the available IDs and
  offsets.
- `pptx add-text-box --unit in --x 1 --y 1 --width 4 --height 1 --text Hello`
  is the easy way to drop text on a slide created via `pptx create`. The
  `pptx set-title` / `pptx set-body` shortcuts only work on decks that ship
  real placeholder shapes (most of our fixtures don't).
- For XLSX, `xlsx fill-formula --range D2:D6 --formula '=B{row}*C{row}/100'`
  expands `{row}` / `{col}` per cell so you don't have to dispatch one
  command per row.
- A non-zero exit code with a `"rejection"` block in the JSON usually means
  your selector or shape ID is wrong, not that the CLI is broken — re-read
  the snapshot first.

---

## Deliverable

When you finish, write `/tmp/officeai-demo/REPORT.md` covering:

1. **What worked.** Which subcommands you used, in what order, and roughly how
   many shell calls each document took.
2. **What surprised you.** Anything counter-intuitive in command names,
   payload shapes, output formats, or discovery flow.
3. **Gaps you hit.** Real or perceived holes — missing subcommands, confusing
   error messages, options you wished existed. For each, note the workaround
   you used (if any).
4. **Time-to-first-doc.** Roughly how long from `$OA --help` to a passing
   verification block for each of the three files.
5. **One concrete suggestion** for the next CLI iteration that would have
   saved you the most time.

Keep the report under ~300 lines. The point is not to document every command —
it's to surface the friction the _next_ agent hitting this CLI will feel.

Good luck.
