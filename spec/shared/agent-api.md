# Shared Agent API

> The unified surface every format exposes to AI agents and the CLI.
> Headless. No DOM. No React. No browser.

## Interface

```typescript
import type { Command, Mutation, DocumentSnapshot, DocumentDiff } from "@officeai/core";

export interface DocumentAgent<TSnapshot extends DocumentSnapshot = DocumentSnapshot> {
  // ── Read ─────────────────────────────────────────────────────────────
  /** Current working snapshot (approved + pending). */
  getSnapshot(): TSnapshot;
  /** Format-specific projection. For DOCX: paragraphs in a range. For XLSX: cells. */
  getRange(range: RangeSpec): RangeSnapshot;
  /** Plain-text or markdown digest convenient for prompts. */
  toMarkdown(): string;
  /** Find content matching a query. */
  search(query: SearchSpec): SearchResult[];

  // ── Write (every write goes through the bus, returns a Mutation) ─────
  applyCommand(command: Command): Promise<Mutation<TSnapshot>>;
  /** Atomic batch — all-or-nothing. */
  applyCommands(commands: ReadonlyArray<Command>): Promise<ReadonlyArray<Mutation<TSnapshot>>>;

  // ── Diff & Review ────────────────────────────────────────────────────
  getDiff(from: TSnapshot, to: TSnapshot): DocumentDiff;
  getPendingMutations(): ReadonlyArray<Mutation<TSnapshot>>;
  approveMutation(id: string): void;
  rejectMutation(id: string): void;
  /** Rewind the approved history to a previous revision. Drops pending. */
  rollback(toRevision: number): void;

  // ── I/O ──────────────────────────────────────────────────────────────
  /** Replace the document with one parsed from a buffer. */
  importFile(buffer: ArrayBuffer): Promise<void>;
  /** Serialize the working snapshot to OOXML bytes. */
  exportFile(): Promise<ArrayBuffer>;

  // ── Lifecycle ────────────────────────────────────────────────────────
  /** Subscribe to working-snapshot changes. Returns unsubscribe. */
  subscribe(listener: (snapshot: TSnapshot, mutation: Mutation<TSnapshot>) => void): () => void;
}

export type RangeSpec =
  | { kind: "docx-paragraphs"; start: number; end: number }
  | { kind: "xlsx-cells"; sheet: string; range: string }
  | { kind: "pptx-shape"; slideIndex: number; shapeId: string };

export type RangeSnapshot = unknown;

export interface SearchSpec { query: string; caseSensitive?: boolean; regex?: boolean; }
export interface SearchResult { path: ReadonlyArray<string | number>; preview: string; }
```

## Construction

Each format exposes a factory:

```typescript
import { DocxAgent } from "@officeai/docx/agent";

const agent = await DocxAgent.fromBuffer(fs.readFileSync("contract.docx"));
const out = await agent.applyCommand({
  type: "docx:insert-text",
  payload: { at: { paragraph: 0, run: 0, offset: 0 }, text: "DRAFT — " },
  source: "agent",
  agentId: "claude-1",
  timestamp: Date.now(),
  sessionId: "s-1",
});
console.log(out.diff);
const buffer = await agent.exportFile();
fs.writeFileSync("contract-edited.docx", Buffer.from(buffer));
```

## Headless invariant

`@officeai/docx/agent` (and its peers) **must not import**:

- `react`, `react-dom`
- `prosemirror-view` (the renderer; ProseMirror **state** types are fine —
  the headless code does not construct an `EditorView`)
- `next`, anything from `apps/web`
- DOM globals (`document`, `window`)

This is enforced via package-boundary lint and verified by a Node-only
agent test that `import`s the package and runs a round-trip.

## CLI as a thin shell

The `office-agent` CLI never re-implements logic; it instantiates the
appropriate `DocumentAgent`, dispatches commands, and prints
JSON/markdown/CSV. See [`spec/agent/cli.md`](../agent/cli.md).
