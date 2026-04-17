# Build Office-Compatible, AI-Native Document Editors: DOCX, XLSX, PPTX

## Mission

You are a senior software architect and engineer. You will autonomously complete a full build of three browser-embeddable, AI-native document editors — for DOCX (Word), XLSX (Excel), and PPTX (PowerPoint) — in one continuous session.

Work in this exact sequence, without skipping ahead:

1. Spec DOCX → Build DOCX → Validate DOCX
2. Spec XLSX → Build XLSX → Validate XLSX
3. Spec PPTX → Build PPTX → Validate PPTX

Do not start Step 2 until Step 1 is fully validated. Do not start building until the spec for that format is complete.

---

## Non-Negotiable Quality Bar

**Office Roundtrip Integrity.** Every editor must satisfy this bar before moving on:

- Open a real-world OOXML file produced by actual Microsoft Office
- Perform a trivial edit (change one word / one cell value / one text box)
- Save back to the original format
- Reopen in LibreOffice headless (CI) and visually/structurally compare
- All OOXML parts the editor did not touch must be byte-preserved
- No corruption, no silent data loss, no structural damage

This is the only acceptance criterion that cannot be traded away. Everything else is scope.

---

## Legal Constraint (Clean-Room Approach)

You will analyze reference repositories to extract concepts, patterns, and architectural decisions. You will then build a fresh implementation from a specification you derive — not a fork, not a dependency, not a copy.

**Allowed:** Study public code, extract architecture concepts, describe behavior and algorithms at the conceptual level, implement independently from first principles + OOXML spec. You are also absolutely allowed of course to use open source libraries for building, e.g. MIT or Apache-licensed etc that we consider unproblematic to build on.

**Not allowed:** Copy code verbatim, lightly rename identifiers, use any AGPL-licensed component as a runtime dependency, import reference repos as packages.

**Runtime dependencies permitted** (MIT / Apache 2.0 / BSD only):

- ProseMirror (MIT) — rich text editing substrate for DOCX
- Y.js (MIT) — CRDT collaboration primitives
- JSZip (MIT) — OOXML container (zip) handling
- SheetJS Community Edition (Apache 2.0) — XLSX parse/serialize fallback
- PptxGenJS (MIT) — PPTX serialization helper
- Zod (MIT) — runtime schema validation
- Any other MIT/Apache/BSD library if justified in the spec

---

## Reference Repositories

Study these before speccing each format. Read architecture docs, main source files, and tests. Do not copy. Understand.

### DOCX

- https://github.com/eigenpal/docx-js-editor (MIT) — primary reference: ProseMirror-based, direct OOXML manipulation, tracked changes, comments, Yjs collaboration
- https://github.com/eigenpal/docx-editor (MIT) — same team, slightly different cut; study both
- https://github.com/dolanmiu/docx (MIT) — mature DOCX read/write; study serialization edge cases

### XLSX + Unified Architecture

- https://github.com/dream-num/univer (Apache 2.0 OSS core) — primary reference for multi-format unified architecture, plugin system, command/mutation pattern, agent integration
- https://github.com/dream-num/univer-mcp — agent integration pattern
- https://github.com/dream-num/skills — agent skill definitions
- https://github.com/SheetJS/sheetjs (Apache 2.0) — XLSX parse/serialize, cell format handling

### PPTX

- https://github.com/gitbrent/PptxGenJS (MIT) — PPTX generation and serialization
- https://github.com/pipipi-pikachu/PPTist (AGPL — study architecture only, zero code copying)

### Format Standards (canonical truth, always prefer over implementations)

- OOXML spec: https://ecma-international.org/publications-and-standards/standards/ecma-376/
- Office Open XML reference: http://officeopenxml.com/
- Microsoft format docs: https://learn.microsoft.com/en-us/openspecs/office_standards/

---

## Project Structure

Create this monorepo from the start:

```
/
  packages/
    core/           # shared: document model, command bus, plugin system, OOXML utils
    docx/           # DOCX editor: parser, model, renderer, serializer, agent API
    xlsx/           # XLSX editor: parser, model, renderer, formula engine, serializer, agent API
    pptx/           # PPTX editor: parser, model, renderer, serializer, agent API
    agent/          # shared agent CLI + programmatic API (across all three formats)
  spec/             # living specification (produced in spec phase, updated during build)
    docx/
    xlsx/
    pptx/
    agent/
    shared/
  fixtures/         # real-world OOXML test files (one folder per format)
  tests/
    roundtrip/      # roundtrip integrity tests per format
    agent/          # agent API tests
  docs/
    build-log/      # decisions, discoveries, deviations from spec
```

---

## Phase Structure (Repeat for Each Format)

### Step A: Analyze

Before writing a single line of spec or code, deeply study the reference repos for this format. Specifically answer:

1. How does the reference implementation model the document in memory? What are the core data types?
2. How does it parse OOXML into that model? What's the parsing strategy?
3. How does it serialize back to OOXML? How does it preserve untouched parts?
4. What's the mutation/command pattern? How is an edit represented?
5. How does it handle the hard parts (tables for DOCX, formulas for XLSX, slide layout for PPTX)?
6. What does it get wrong or sacrifice that we should improve?
7. What's missing from the 80% scope we need?

Write analysis notes in `/spec/{format}/analysis.md`. These notes inform the spec but are not the spec.

---

### Step B: Spec

Produce the specification for this format. The spec is the contract for the build. It must be complete enough that someone could implement it independently.

**Required spec documents for each format:**

#### `/spec/shared/` (produce once, before DOCX)

- `document-model.md` — the shared abstraction: what is a "document", "node", "command", "mutation", "snapshot", "diff" in our system; must work for all three formats
- `command-bus.md` — the shared command/mutation pattern; every edit (human or agent) flows through this; commands are serializable JSON; they compose into diffs; diffs are reviewable
- `plugin-system.md` — how features are registered; how a plugin declares what commands it handles; how UI and logic stay separate
- `ooxml-utils.md` — shared utilities: zip container, XML parse/write, namespace handling, relationship graph
- `agent-api.md` — the unified agent API: what operations every agent needs across all formats; the contract between AI and document

#### `/spec/docx/` (or xlsx/ or pptx/)

- `feature-scope.md` — exactly what is IN the 80% and what is explicitly OUT; no ambiguity
- `document-model.md` — the in-memory representation specific to this format; data types, node types, property schemas
- `ooxml-mapping.md` — for every in-scope feature: which OOXML elements map to which model nodes, and back; include the OOXML tag names and attribute names; this is the most important spec doc
- `parser.md` — algorithm for OOXML → internal model; how unknown/unsupported elements are preserved as opaque blobs
- `serializer.md` — algorithm for internal model → OOXML; how opaque blobs are re-emitted untouched; how roundtrip integrity is maintained
- `renderer.md` — how the model is rendered in the browser; what technology (ProseMirror / Canvas / SVG); how rendering state stays in sync with model
- `formula-engine.md` (XLSX only) — formula evaluation: which functions, evaluation order, cell dependency graph, error handling, Excel-compatible behavior
- `agent-commands.md` — the complete list of agent-callable commands for this format; for each command: name, parameters, effect on model, OOXML impact, example
- `edge-cases.md` — known hard cases and how we handle them; what we degrade gracefully; what triggers the "open in Office" fallback
- `acceptance-criteria.md` — measurable done criteria: which fixture files must roundtrip clean, which agent commands must work, which user flows must work

**Spec quality bar:** A spec document is done when:

- It is self-contained (doesn't assume knowledge from the reference repos)
- It is precise (data types have explicit shapes, algorithms have pseudocode or step-by-step prose)
- It is honest (scope exclusions are explicit, uncertainties are flagged)
- It is actionable (someone could implement from it without asking clarifying questions)

Do NOT begin building until the spec passes this bar.

---

### Step C: Build

Implement the format based on the spec. Follow this sub-order:

1. **OOXML I/O first** — parser and serializer before any UI; validate against fixtures immediately
2. **Document model** — the in-memory representation; pure TypeScript, no DOM, no React
3. **Command bus** — the mutation layer; pure TypeScript; headless-testable
4. **Agent API** — expose commands programmatically; test this before building the UI
5. **Renderer** — connect model to browser UI; wire human interactions to the command bus
6. **Polish** — error handling, performance, graceful degradation, "open in Office" fallback trigger

**Build discipline:**

- Write roundtrip tests before implementing each parser/serializer feature
- Every command must be testable headlessly (no browser required)
- Every PR-equivalent commit must not reduce roundtrip pass rate
- Keep `/docs/build-log/{format}.md` updated with non-trivial decisions

---

### Step D: Validate

Before declaring a format complete and moving to the next:

Run the full validation suite:

- [ ] Roundtrip tests pass on all fixtures in `/fixtures/{format}/`
- [ ] Agent API tests pass: all commands work headlessly
- [ ] Performance: large files render and edit smoothly (100-page DOCX / 50k-row XLSX / 50-slide PPTX)
- [ ] Manual spot-check: open a produced file in LibreOffice, confirm visual fidelity
- [ ] License check: no AGPL or proprietary runtime dependency
- [ ] Spec and build log are up to date

Only after all boxes are checked: move to the next format.

---

## The 80% Scope per Format

### DOCX — In Scope

- Open any real-world .docx; edit; save; reopen in Word without fidelity loss
- Text: bold, italic, underline, strikethrough, font family/size/color, highlight
- Paragraphs: alignment, indentation, spacing, heading styles (H1–H6), lists (bullet, numbered, nested)
- Tables: create/edit, merge/split cells, borders, background color, column widths
- Images: inline and floating, resize, reposition; preserve image quality
- Headers and footers (including different first page)
- Page breaks, section breaks (preserve but not edit section properties)
- Comments: add, reply, resolve, delete; preserve existing threading
- Tracked changes (suggestion mode): wrap edits in revision markup; accept/reject
- Hyperlinks
- Preserve all OOXML parts not explicitly modified (styles, fonts, numbering, relationships, custom XML, etc.)

### DOCX — Explicitly Out of Scope

- Mail merge / field codes (preserve, do not edit)
- Complex cross-references and footnotes (preserve, do not edit)
- Embedded objects (OLE, charts linked to Excel — preserve as blobs)
- VBA macros (preserve, do not execute)
- Custom XML data binding

---

### XLSX — In Scope

- Open any real-world .xlsx; edit; save; reopen in Excel without fidelity loss
- Cell editing: values, formulas, deletion
- **Formula engine:** implement the ~150 highest-priority functions (see priority list below)
- Cell formats: number, currency, percentage, date, time, text; custom format strings
- Font, fill, borders, alignment per cell
- Merged cells
- Multiple sheets; sheet rename; sheet tab color
- Freeze panes
- Filters (apply; preserve existing; reset)
- Sorting (by column)
- Comments / notes
- Hyperlinks
- Conditional formatting: highlight rules (greater than, less than, equal, between, text contains, date occurring); color scales; data bars (basic)
- Charts: line, bar (clustered/stacked), column (clustered/stacked), pie, area — render existing; create basic new ones
- Preserve all OOXML parts not explicitly modified (pivot cache, VBA, custom XML, printer settings, etc.)

**Formula priority list (implement in this order, stop at 150 or time runs out):**

Math/Stats: SUM, AVERAGE, COUNT, COUNTA, COUNTBLANK, MIN, MAX, SUMIF, SUMIFS, COUNTIF, COUNTIFS, AVERAGEIF, AVERAGEIFS, ROUND, ROUNDUP, ROUNDDOWN, INT, ABS, MOD, POWER, SQRT, CEILING, FLOOR, RAND, RANDBETWEEN, LARGE, SMALL, RANK, MEDIAN, STDEV, VAR, PRODUCT, SUMPRODUCT

Logic: IF, IFS, AND, OR, NOT, XOR, IFERROR, IFNA, SWITCH, TRUE, FALSE

Lookup: VLOOKUP, HLOOKUP, INDEX, MATCH, XLOOKUP (if time), CHOOSE, OFFSET, INDIRECT, ROW, ROWS, COLUMN, COLUMNS

Text: CONCATENATE, CONCAT, TEXTJOIN, LEFT, RIGHT, MID, LEN, TRIM, UPPER, LOWER, PROPER, FIND, SEARCH, SUBSTITUTE, REPLACE, REPT, TEXT, VALUE, NUMBERVALUE, CHAR, CODE, EXACT, T

Date/Time: TODAY, NOW, DATE, TIME, YEAR, MONTH, DAY, HOUR, MINUTE, SECOND, WEEKDAY, WEEKNUM, EOMONTH, EDATE, DATEDIF, NETWORKDAYS, WORKDAY, DATEVALUE, TIMEVALUE

Finance: PMT, PV, FV, RATE, NPER, NPV, IRR, SLN

Info: ISBLANK, ISNUMBER, ISTEXT, ISERROR, ISNA, ISODD, ISEVEN, ISODD, TYPE, N, NA

Array (modern Excel): UNIQUE, SORT, FILTER, SEQUENCE — implement if time, otherwise defer

### XLSX — Explicitly Out of Scope

- Power Query / Power Pivot
- VBA macro execution (preserve, do not run)
- What-If Analysis tools (Goal Seek, Solver, Scenario Manager)
- Advanced pivot table editing (preserve existing; create new: defer)
- Sparklines (preserve; create new: defer)
- Slicers (preserve)
- External data connections (preserve)

---

### PPTX — In Scope

- Open any real-world .pptx; edit; save; reopen in PowerPoint without fidelity loss
- Render all slides faithfully (what you see in PowerPoint should look right in our renderer)
- Edit text in text boxes and shapes (content, font, size, color, bold/italic/underline)
- Reposition shapes, images, text boxes (drag/move, property-set position)
- Resize shapes and images
- Add new slides (from a small set of blank/title/content layout templates)
- Duplicate existing slides
- Delete slides
- Reorder slides (drag in slide panel)
- Insert images into slides
- Preserve ALL elements we don't edit: SmartArt, animations, transitions, embedded charts, tables, videos, audio — roundtrip byte-clean

### PPTX — Explicitly Out of Scope

- SmartArt editing (render + preserve only)
- Animation and transition editing (preserve only)
- Chart editing within slides (preserve only)
- Master slide / layout editing (preserve only)
- Notes pages editing (preserve only)
- Table editing within slides (render + basic text edit; full edit deferred)

---

## The AI-Native Design (Most Important Section)

This is the core differentiator. Every editor must be designed from the ground up so that AI agents are first-class users — not an afterthought bolted on top.

### Core Principle: Everything Is a Command

No direct model mutation is ever allowed. Every change — whether made by a human clicking a button, or by an AI agent calling an API — flows through the **command bus**. This is not optional architecture; it is the invariant that makes everything else possible.

A command is:

```typescript
interface Command<T extends string, P> {
  type: T; // e.g. "docx:insert-text", "xlsx:set-cell-value"
  payload: P; // fully typed, serializable to JSON
  source: "human" | "agent" | "system";
  agentId?: string; // which agent, if source === 'agent'
  timestamp: number;
  sessionId: string;
}
```

A mutation is the result of applying a command to a document snapshot:

```typescript
interface Mutation {
  command: Command;
  before: DocumentSnapshot; // or just the affected subtree for efficiency
  after: DocumentSnapshot;
  diff: DocumentDiff; // structured diff of what changed
}
```

### The Agent API

Every format exposes the same top-level agent interface pattern:

```typescript
interface DocumentAgent {
  // Read
  getSnapshot(): DocumentSnapshot;
  getRange(range: RangeSpec): RangeSnapshot; // cells for xlsx, selection for docx/pptx
  search(query: SearchSpec): SearchResult[];

  // Write (every write goes through command bus, returns a Mutation)
  applyCommand(command: Command): Promise<Mutation>;
  applyCommands(commands: Command[]): Promise<Mutation[]>; // atomic batch

  // Diff & Review
  getDiff(fromSnapshot: DocumentSnapshot, toSnapshot: DocumentSnapshot): DocumentDiff;
  getPendingMutations(): Mutation[]; // agent changes not yet approved by human
  approveMutation(mutationId: string): void;
  rejectMutation(mutationId: string): void;
  rollback(mutationId: string): void;

  // I/O
  importFile(buffer: ArrayBuffer, format: "docx" | "xlsx" | "pptx"): Promise<void>;
  exportFile(format: "docx" | "xlsx" | "pptx"): Promise<ArrayBuffer>;
}
```

This interface must work **headlessly** — with zero DOM, zero React, zero browser. An AI agent running in Node.js on a server must be able to:

- Load a document from a buffer
- Read data from it
- Apply commands to it
- Export the result
- Get a structured diff of what changed

The browser editor is just a UI skin on top of the same agent interface.

### Format-Specific Agent Commands

Each format must expose a complete set of typed commands. Below is the minimum required set. During the spec phase, expand and fully type each one.

#### DOCX Agent Commands

```
docx:insert-text        { at: Position, text: string }
docx:delete-range       { range: Selection }
docx:format-range       { range: Selection, format: TextFormat }
docx:insert-paragraph   { at: Position, style?: ParagraphStyle }
docx:insert-table       { at: Position, rows: number, cols: number }
docx:set-cell-content   { tableId: string, row: number, col: number, content: Node[] }
docx:insert-image       { at: Position, data: ArrayBuffer, width: number, height: number }
docx:add-comment        { range: Selection, text: string, author: string }
docx:resolve-comment    { commentId: string }
docx:accept-change      { changeId: string }
docx:reject-change      { changeId: string }
docx:set-paragraph-style { at: Position, style: string }
```

#### XLSX Agent Commands

```
xlsx:set-cell-value     { sheet: string, ref: CellRef, value: CellValue }
xlsx:set-cell-formula   { sheet: string, ref: CellRef, formula: string }
xlsx:set-cell-format    { sheet: string, ref: RangeRef, format: CellFormat }
xlsx:insert-row         { sheet: string, at: number, count: number }
xlsx:insert-column      { sheet: string, at: number, count: number }
xlsx:delete-row         { sheet: string, at: number, count: number }
xlsx:delete-column      { sheet: string, at: number, count: number }
xlsx:merge-cells        { sheet: string, range: RangeRef }
xlsx:unmerge-cells      { sheet: string, range: RangeRef }
xlsx:add-sheet          { name: string, at?: number }
xlsx:rename-sheet       { name: string, newName: string }
xlsx:set-range-values   { sheet: string, range: RangeRef, values: CellValue[][] }
xlsx:add-comment        { sheet: string, ref: CellRef, text: string, author: string }
xlsx:set-conditional-format { sheet: string, range: RangeRef, rule: ConditionalFormatRule }
```

#### PPTX Agent Commands

```
pptx:add-slide          { at?: number, layoutId?: string }
pptx:delete-slide       { slideIndex: number }
pptx:duplicate-slide    { slideIndex: number }
pptx:move-slide         { from: number, to: number }
pptx:set-text           { slideIndex: number, shapeId: string, text: string }
pptx:format-text        { slideIndex: number, shapeId: string, range: TextRange, format: TextFormat }
pptx:set-position       { slideIndex: number, shapeId: string, x: number, y: number }
pptx:set-size           { slideIndex: number, shapeId: string, width: number, height: number }
pptx:insert-image       { slideIndex: number, data: ArrayBuffer, x: number, y: number, width: number, height: number }
pptx:add-text-box       { slideIndex: number, text: string, x: number, y: number, width: number, height: number }
```

### The Human Review Flow

When source is 'agent', mutations are staged — not immediately applied to the "approved" document state. They go into a pending queue:

```
Document State:
  ├── approved:  DocumentSnapshot  (what the human has okayed)
  ├── pending:   Mutation[]        (agent proposals, not yet approved)
  └── working:   DocumentSnapshot  (approved + pending = what the browser shows)
```

The UI renders the working state. Pending agent mutations are visually marked (like tracked changes in Word). The human can:

- **Approve all** → pending mutations move to approved history
- **Approve one** → that mutation moves; others stay pending
- **Reject one** → removed from pending; document reverts that change
- **Rollback to snapshot** → approved history is rewound to a previous point

This model must be built into the core command bus — not per-format, not per-component.

### The CLI / Programmatic Interface

Produce a CLI tool (`office-agent`) that wraps the headless agent API:

```bash
# XLSX examples
office-agent xlsx inspect --file report.xlsx
# → sheets, used range, formula groups, table definitions

office-agent xlsx read --file report.xlsx --sheet "Sales" --range "A1:F100" --format csv
# → CSV to stdout

office-agent xlsx write --file report.xlsx --sheet "Summary" --data ./summary.csv
# → writes, outputs mutation diff as JSON

office-agent xlsx formula --file report.xlsx --sheet "Sales" --cell "G2" --formula "=SUMIF(C:C,\"DE\",F:F)"
# → sets formula, outputs result + diff

office-agent xlsx diff --before report_v1.xlsx --after report_v2.xlsx --format json
# → structured diff of what changed between two versions

# DOCX examples
office-agent docx read --file contract.docx --format markdown
# → converts doc to markdown for AI consumption

office-agent docx write --file contract.docx --at "section:2/paragraph:3" --text "Updated clause text"
# → applies targeted edit, outputs diff

office-agent docx comment --file contract.docx --range "section:2/paragraph:3" --text "Review this clause" --author "AI Agent"

# PPTX examples
office-agent pptx inspect --file deck.pptx
# → slide count, shapes per slide, text content summary

office-agent pptx read --file deck.pptx --slide 3 --format json
# → structured representation of slide 3

office-agent pptx write --file deck.pptx --slide 3 --shape "title" --text "New Title"
```

The CLI is the primary interface for AI agents running in server-side pipelines. Design it to be pipeable, scriptable, and composable with standard UNIX tools (stdin/stdout/stderr, exit codes, JSON output).

---

## Import / Export Requirements

This is critical. Every editor must:

### Import

- Accept any valid OOXML file produced by Microsoft Office 2010, 2013, 2016, 2019, 2021, 365 (desktop), 365 (web)
- Accept files produced by LibreOffice (may have minor OOXML deviations — handle gracefully)
- Accept files produced by Google Docs/Sheets/Slides export
- On import failure (corrupted file, unsupported OOXML variant): surface a clear error, never crash silently

### Export

- Produce valid OOXML conformant to ECMA-376
- Files must open without repair dialogs in Microsoft Office
- Files must not trigger "this file was created by an older version" false warnings
- Files must preserve all relationships, styles, and metadata present in the original
- Export must be available both from browser (download) and from Node.js (Buffer)

### Import/Export API

```typescript
// Browser
const editor = await DocxEditor.fromBuffer(arrayBuffer)
const outputBuffer = await editor.export('docx')
const blob = new Blob([outputBuffer], { type: 'application/vnd.openxmlformats-...' })

// Node.js (headless)
const agent = await DocxAgent.fromBuffer(fs.readFileSync('document.docx'))
await agent.applyCommand({ type: 'docx:insert-text', ... })
const output = await agent.exportFile('docx')
fs.writeFileSync('document-edited.docx', output)

// XLSX
const agent = await XlsxAgent.fromBuffer(fs.readFileSync('data.xlsx'))
const values = await agent.getRange({ sheet: 'Sheet1', range: 'A1:Z100' })
await agent.applyCommand({ type: 'xlsx:set-cell-value', sheet: 'Sheet1', ref: 'B2', value: 42 })
const output = await agent.exportFile('xlsx')
```

---

## Fixture Corpus

Before building each format, collect real-world test files. This is not optional.

### DOCX Fixtures (collect before building DOCX)

- 5 German business letters (Geschäftsbriefe) — headers, addresses, simple formatting
- 5 contracts (Verträge) — complex formatting, numbered lists, headers/footers, tracked changes
- 3 invoices (Rechnungen) — tables, addresses, totals
- 3 reports (Berichte) — headings, images, multi-page, table of contents
- 2 CVs / resumes — tables, columns, images
- 2 "template" documents with lots of {placeholder} style markers
- 2 files exported from Google Docs
- 2 files created by LibreOffice

### XLSX Fixtures (collect before building XLSX)

- 5 budget/calculation sheets (Kalkulationen) — many formulas, named ranges, conditional formatting
- 5 data tables (Datentabellen) — filters, large datasets (5k+ rows), multiple sheets
- 3 invoices in Excel format — templates with formulas
- 3 project plans (Projektpläne) — dates, conditonal formatting, status tracking
- 2 financial models — PMT, NPV, IRR functions, data validation
- 2 pivot-heavy files — to verify we roundtrip pivot definitions correctly even if we don't edit them
- 2 chart-heavy files — to verify chart preservation
- 2 files exported from Google Sheets

### PPTX Fixtures (collect before building PPTX)

- 5 corporate slide decks — title slides, content slides, mixed layouts
- 3 sales presentations — images, charts, tables
- 3 technical/product presentations — code screenshots, diagrams
- 2 presentations with SmartArt — to verify we preserve without breaking
- 2 presentations with animations — to verify we preserve without breaking
- 2 files exported from Google Slides

If you cannot access real files during the build, generate realistic synthetic fixtures using python-pptx, openpyxl, and python-docx — but flag these as synthetic and plan to replace with real-world files before production.

---

## Architecture Principles (Non-Negotiable)

1. **Headless-first.** The core of every editor (parser, model, command bus, serializer) runs in Node.js with zero DOM. The browser is just a rendering surface. This is what makes the agent API real.

2. **Commands are the only mutation path.** Direct model mutation is never allowed outside of the parser. Everything else goes through the command bus. This is the invariant that enables diffs, review, rollback, and multi-agent coordination.

3. **OOXML is the source of truth, not the model.** The in-memory model is a working surface. The OOXML file is what's saved. When in doubt, keep the OOXML closer to the original.

4. **Opaque blob preservation.** Any OOXML element the editor doesn't understand is parsed into an opaque blob and re-emitted verbatim on export. We never silently drop unknown content.

5. **Separation of concerns.** Model ↔ Command Bus ↔ Renderer ↔ Agent API. Each layer has one job. No shortcuts that fuse layers together.

6. **Format agnosticism in the core.** The command bus, plugin system, and agent API are designed in `packages/core` and are format-agnostic. Each format package implements the format-specific parts. The agent CLI works the same way regardless of format.

7. **Progressive rendering.** Large files must not block the UI. Parse and render incrementally. Virtual scrolling for XLSX. Lazy page rendering for DOCX. Lazy slide thumbnail rendering for PPTX.

8. **Fail loudly.** Import failures, formula errors, and roundtrip anomalies surface as structured errors with useful messages — never as silent data corruption.

---

## Output at the End of Each Format

When you complete a format (Spec → Build → Validate), produce:

1. **`/spec/{format}/`** — all spec documents, complete and up-to-date
2. **`/packages/{format}/`** — the implementation
3. **`/tests/roundtrip/{format}/`** — passing roundtrip test suite
4. **`/docs/build-log/{format}.md`** — decisions, deviations from spec, known issues
5. **A summary comment** in the session describing: what was built, what passes, what's deferred, what was harder than expected

---

## Start Instructions

1. Read this entire prompt twice.
2. Set up the monorepo structure.
3. Collect or generate the DOCX fixture corpus.
4. Begin the DOCX analysis phase — study the reference repos for DOCX.
5. Produce the DOCX spec.
6. Build the DOCX editor.
7. Validate DOCX against fixtures.
8. Move to XLSX. Repeat.
9. Move to PPTX. Repeat.

Before starting, confirm:

- You understand the clean-room constraint and will not copy code
- You understand the roundtrip integrity bar and will not move forward without passing it
- You understand the headless-first / agent-first design requirement
- You understand the product sequence: DOCX complete before XLSX starts, XLSX complete before PPTX starts

Ask no clarifying questions. Begin.
