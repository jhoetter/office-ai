# Shared Command Bus

> The single mutation path for every editor. Required reading.

## Invariant

**No code anywhere in this monorepo mutates a document snapshot directly.**
Every change — a keystroke from a human, a tool call from an agent, a CLI
invocation — must produce a `Command`, dispatch it through a `CommandBus`,
and use the resulting `Mutation`.

The parser is the only exception: it produces the _initial_ snapshot. After
that, the bus is the law.

## Types

```typescript
export interface Command<TType extends string = string, TPayload = unknown> {
  /** Format-prefixed kebab string, e.g. "docx:insert-text". */
  readonly type: TType;
  /** Fully typed, JSON-serializable. No functions, no class instances. */
  readonly payload: TPayload;
  readonly source: "human" | "agent" | "system";
  /** Set when source === "agent". Identifies which agent issued this. */
  readonly agentId?: string;
  /** Wall-clock at dispatch time (ms since epoch). */
  readonly timestamp: number;
  /** Logical session id; survives reloads if persisted. */
  readonly sessionId: string;
}

export interface Mutation<TSnapshot = unknown> {
  readonly id: string;
  readonly command: Command;
  readonly before: TSnapshot;
  readonly after: TSnapshot;
  readonly diff: DocumentDiff;
  /** Set by the bus; "approved" until reviewed if source==='agent'. */
  status: "approved" | "pending" | "rejected" | "rolled-back";
}

export interface CommandHandler<TPayload, TSnapshot> {
  /** Handler-declared command type, e.g. "docx:insert-text". */
  readonly type: string;
  /** Pure function: snapshot + payload -> next snapshot + diff. */
  apply(
    snapshot: TSnapshot,
    payload: TPayload,
    ctx: HandlerContext
  ): {
    next: TSnapshot;
    diff: DocumentDiff;
  };
}

export interface HandlerContext {
  readonly mintNodeId: () => NodeId;
  readonly now: () => number;
}
```

## Dispatch contract

```typescript
class CommandBus<TSnapshot> {
  register(handler: CommandHandler<unknown, TSnapshot>): void;
  /** Dispatches one command. Throws if no handler is registered for type. */
  dispatch(command: Command): Promise<Mutation<TSnapshot>>;
  /** Atomically apply a list. If any handler throws, none are applied. */
  dispatchAll(commands: ReadonlyArray<Command>): Promise<ReadonlyArray<Mutation<TSnapshot>>>;
  getSnapshot(): TSnapshot;
  /** Subscribe to the working snapshot. Fires after every mutation. */
  subscribe(listener: (snapshot: TSnapshot, mutation: Mutation<TSnapshot>) => void): () => void;
}
```

## Staging tri-state (per [`prompt.md`](../../prompt.md) lines 437–451)

```
Document State:
  approved : DocumentSnapshot      // what the human has okayed
  pending  : Mutation[]             // agent proposals, not yet approved
  working  : DocumentSnapshot       // approved + pending = what the UI shows
```

Rules:

- A command with `source === "agent"` is applied to the **working snapshot**
  and produces a `Mutation` with `status === "pending"`. The approved
  snapshot is unchanged.
- A command with `source === "human"` or `source === "system"` is applied
  to the working snapshot and is **immediately approved** (it both updates
  approved and the pending stack collapses against it).
- `approveMutation(id)` collapses one pending mutation into approved (and
  rebases the remaining pending stack on the new approved snapshot).
- `rejectMutation(id)` removes a pending mutation; the working snapshot is
  recomputed from approved + remaining pending.
- `rollback(toRevision)` rewinds approved; pending stack is dropped.

## Undo / Redo (Phase 13h)

The bus exposes a per-author Undo / Redo trail layered on top of the
approved history. The contract mirrors Excel's behaviour exactly:

- `canUndo(): boolean` — true iff at least one approved mutation
  exists in the history.
- `canRedo(): boolean` — true iff `redoStack` is non-empty.
- `undo(): Mutation | null` — pops the most recent approved mutation,
  flips its `status` to `"undone"`, restores `approved = mutation.before`,
  and pushes the mutation onto `redoStack`. `pending` mutations are
  re-applied atop the rolled-back approved snapshot so the working
  view stays consistent. Subscribers fire with the new working snapshot.
  Returns `null` (no-op) when the history has no approved entries.
- `redo(): Mutation | null` — pops `redoStack` and **re-runs the
  handler** against the current `approved` snapshot rather than blindly
  restoring the recorded `after`. This keeps redo correct after the
  pending stack has rebased the world out from under the undone
  mutation. The fresh mutation is appended to `history` with
  `status: "approved"`. Returns `null` when nothing to redo.
- Any new authored mutation (human / system / agent-then-approved)
  **clears** `redoStack`. Branching the timeline kills the redo trail;
  this is the same rule Word, Excel, PowerPoint, VS Code and every
  other tree-undo system uses.

The `MutationStatus` union therefore gains `"undone"` alongside
`"approved" | "pending" | "rejected" | "rolled-back"`.

`XlsxAgent` and `DocxAgent` proxy `canUndo`, `canRedo`, `undo`, and
`redo` directly to the bus. The MCP server exposes `xlsx_undo` and
`xlsx_redo` tools that return `{ did_undo, can_undo, can_redo,
undone: { id, type } | null }` for symmetry. The `oa xlsx apply-file`
CLI accepts an optional `--undo <n>` flag that peels back the last `n`
approved mutations after a batch apply, useful for batch-then-revert
diffing workflows.

## MutationStore

```typescript
class MutationStore<TSnapshot> {
  approved: TSnapshot;
  pending: Mutation<TSnapshot>[];
  /** Derived: approved with all pending mutations replayed atop. */
  working(): TSnapshot;
  approve(id: string): void;
  reject(id: string): void;
  /** Replay-on-rebase: pending mutations re-execute on top of the new approved. */
  collapse(into: TSnapshot): void;
}
```

## Determinism guarantees

- A handler's `apply` must be a **pure function** of `(snapshot, payload, ctx)`.
  No I/O, no random except via `ctx`. This is what makes diffs reproducible.
- `mintNodeId` is the single source of randomness. Tests inject a
  deterministic minter.
- Handlers must produce a `DocumentDiff` whose `changes` array is sorted by
  `path` for stable serialization.

## Why this matters

- **Agent review flow** is impossible without a mutation that has a structured `diff`.
- **Multi-agent coordination** is impossible without a serializable `Command`.
- **Headless testing** is impossible without pure handlers.
- **Replay & rollback** are impossible without immutable snapshots.

This is the load-bearing column of the system. Build it first, build it well.
