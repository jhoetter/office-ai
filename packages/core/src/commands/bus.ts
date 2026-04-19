import { v4 as uuidv4 } from "uuid";
import type { DocumentDiff, DocumentSnapshot } from "../types/document.js";
import { defaultIdMinter, type IdMinter } from "../util/ids.js";
import {
  CommandError,
  type Command,
  type CommandHandler,
  type CommandLite,
  type HandlerContext,
  type Mutation,
} from "./types.js";

export interface CommandBusOptions {
  readonly sessionId?: string;
  readonly mintNodeId?: IdMinter;
  readonly now?: () => number;
}

export type Listener<TSnapshot> = (snapshot: TSnapshot, mutation: Mutation<TSnapshot>) => void;

/**
 * The command bus + mutation store. Holds the **approved** snapshot,
 * a **pending** stack of agent mutations, and produces the **working**
 * snapshot on demand.
 *
 * See spec/shared/command-bus.md.
 */
export class CommandBus<TSnapshot extends DocumentSnapshot = DocumentSnapshot> {
  private readonly handlers = new Map<string, CommandHandler<unknown, TSnapshot>>();
  private readonly history: Mutation<TSnapshot>[] = [];
  /**
   * Stack of mutations that have been undone and are still eligible
   * for redo. Cleared the moment any new approved mutation lands —
   * matches Excel's "branching kills the redo trail" behaviour.
   */
  private redoStack: Mutation<TSnapshot>[] = [];
  private pending: Mutation<TSnapshot>[] = [];
  private approved: TSnapshot;
  private working: TSnapshot;
  private readonly listeners = new Set<Listener<TSnapshot>>();
  private readonly sessionId: string;
  private readonly ctx: HandlerContext;

  constructor(initial: TSnapshot, opts: CommandBusOptions = {}) {
    this.approved = initial;
    this.working = initial;
    this.sessionId = opts.sessionId ?? uuidv4();
    this.ctx = {
      mintNodeId: opts.mintNodeId ?? defaultIdMinter,
      now: opts.now ?? (() => Date.now()),
    };
  }

  register<TPayload>(handler: CommandHandler<TPayload, TSnapshot>): void {
    this.handlers.set(handler.type, handler as CommandHandler<unknown, TSnapshot>);
  }

  registerAll(handlers: ReadonlyArray<CommandHandler<unknown, TSnapshot>>): void {
    for (const h of handlers) this.register(h);
  }

  hasHandler(type: string): boolean {
    return this.handlers.has(type);
  }

  getApproved(): TSnapshot {
    return this.approved;
  }

  getWorking(): TSnapshot {
    return this.working;
  }

  /** Convenience alias matching the agent API. */
  getSnapshot(): TSnapshot {
    return this.working;
  }

  getPending(): ReadonlyArray<Mutation<TSnapshot>> {
    // Return a defensive copy: callers (notably the CLI) loop with
    // `getPending().forEach(m => approveMutation(m.id))`, which mutates
    // `this.pending` mid-iteration. Handing out the live array caused
    // every other mutation to be skipped and re-applied during
    // `recomputeWorking`, producing duplicated content. See §G0 of
    // docs/cli-gap-report.md for the full trace.
    return [...this.pending];
  }

  getHistory(): ReadonlyArray<Mutation<TSnapshot>> {
    return this.history;
  }

  subscribe(listener: Listener<TSnapshot>): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  /** Issue one command. */
  async dispatch(command: Command | CommandLite): Promise<Mutation<TSnapshot>> {
    const cmd = this.normalize(command);
    return this.applyOne(cmd);
  }

  /** Atomic batch — if any handler throws, none are applied. */
  async dispatchAll(
    commands: ReadonlyArray<Command | CommandLite>
  ): Promise<ReadonlyArray<Mutation<TSnapshot>>> {
    const normalized = commands.map((c) => this.normalize(c));
    const snapshotOfApproved = this.approved;
    const snapshotOfPending = [...this.pending];
    const snapshotOfWorking = this.working;
    const results: Mutation<TSnapshot>[] = [];
    try {
      for (const c of normalized) {
        results.push(this.applyOneSync(c));
      }
      for (const m of results) this.notify(m);
      return results;
    } catch (err) {
      this.approved = snapshotOfApproved;
      this.pending = snapshotOfPending;
      this.working = snapshotOfWorking;
      throw err;
    }
  }

  approveMutation(id: string): void {
    const idx = this.pending.findIndex((m) => m.id === id);
    if (idx === -1) return;
    const m = this.pending[idx];
    m.status = "approved";
    this.approved = m.after;
    this.pending.splice(idx, 1);
    this.history.push(m);
    this.redoStack = [];
    this.recomputeWorking();
  }

  rejectMutation(id: string): void {
    const idx = this.pending.findIndex((m) => m.id === id);
    if (idx === -1) return;
    const m = this.pending[idx];
    m.status = "rejected";
    this.pending.splice(idx, 1);
    this.history.push(m);
    this.recomputeWorking();
  }

  /**
   * Returns true iff there is at least one approved mutation in the
   * history that can be undone (P13).
   */
  canUndo(): boolean {
    for (let i = this.history.length - 1; i >= 0; i--) {
      if (this.history[i]!.status === "approved") return true;
    }
    return false;
  }

  /** True iff `redoStack` has at least one entry. */
  canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  /**
   * Undo the most recent approved mutation by restoring its `before`
   * snapshot. Pending agent mutations are re-applied on top so the
   * `working` view stays consistent.
   *
   * Returns the undone mutation, or `null` if there's nothing to undo.
   */
  undo(): Mutation<TSnapshot> | null {
    let idx = -1;
    for (let i = this.history.length - 1; i >= 0; i--) {
      if (this.history[i]!.status === "approved") {
        idx = i;
        break;
      }
    }
    if (idx === -1) return null;
    const m = this.history[idx]!;
    m.status = "undone";
    this.approved = m.before as TSnapshot;
    this.redoStack.push(m);
    this.recomputeWorking();
    this.notify(m);
    return m;
  }

  /**
   * Re-apply the most recently undone mutation. Returns the
   * re-applied mutation, or `null` if there's nothing to redo.
   *
   * Implementation note: we re-run the handler against the current
   * approved snapshot rather than naively restoring the old `after`,
   * so structurally-rebased pending mutations don't desync.
   */
  redo(): Mutation<TSnapshot> | null {
    const m = this.redoStack.pop();
    if (!m) return null;
    const handler = this.handlers.get(m.command.type);
    if (!handler) {
      // Handler vanished (re-registration mid-session?). Best-effort:
      // restore the recorded `after` and keep going.
      m.status = "approved";
      this.approved = m.after as TSnapshot;
      this.history.push(m);
      this.recomputeWorking();
      this.notify(m);
      return m;
    }
    try {
      const out = handler.apply(this.approved, m.command.payload, this.ctx);
      const re: Mutation<TSnapshot> = {
        ...m,
        before: this.approved,
        after: out.next,
        diff: out.diff,
        status: "approved",
      };
      this.approved = out.next;
      this.history.push(re);
      this.recomputeWorking();
      this.notify(re);
      return re;
    } catch (err) {
      const rejected: Mutation<TSnapshot> = {
        ...m,
        status: "rejected",
        rejection: {
          code: err instanceof CommandError ? err.code : "handler-threw",
          message: err instanceof Error ? err.message : String(err),
        },
      };
      this.history.push(rejected);
      this.notify(rejected);
      return rejected;
    }
  }

  /** Roll back approved history; pending stack is dropped. */
  rollback(toRevision: number): void {
    if (toRevision >= this.approved.revision) return;
    let target = this.approved;
    while (this.history.length > 0 && target.revision > toRevision) {
      const last = this.history[this.history.length - 1];
      if (last.status !== "approved") {
        this.history.pop();
        continue;
      }
      target = last.before as TSnapshot;
      last.status = "rolled-back";
      this.history.pop();
    }
    this.approved = target;
    this.pending = [];
    this.working = this.approved;
  }

  // ── internals ────────────────────────────────────────────────────────────
  private applyOne(cmd: Command): Mutation<TSnapshot> {
    const m = this.applyOneSync(cmd);
    this.notify(m);
    return m;
  }

  private applyOneSync(cmd: Command): Mutation<TSnapshot> {
    const handler = this.handlers.get(cmd.type);
    if (!handler) {
      throw new CommandError("no-handler", `No handler registered for command type "${cmd.type}"`);
    }
    const before = this.working;
    let next: TSnapshot;
    let diff: DocumentDiff;
    try {
      const out = handler.apply(before, cmd.payload, this.ctx);
      next = out.next;
      diff = out.diff;
    } catch (err) {
      const m: Mutation<TSnapshot> = {
        id: uuidv4(),
        command: cmd,
        before,
        after: before,
        diff: {
          format: before.format,
          fromRevision: before.revision,
          toRevision: before.revision,
          changes: [],
        },
        status: "rejected",
        rejection: {
          code: err instanceof CommandError ? err.code : "handler-threw",
          message: err instanceof Error ? err.message : String(err),
        },
      };
      this.history.push(m);
      return m;
    }

    const m: Mutation<TSnapshot> = {
      id: uuidv4(),
      command: cmd,
      before,
      after: next,
      diff,
      status: cmd.source === "agent" ? "pending" : "approved",
    };

    if (cmd.source === "agent") {
      this.pending.push(m);
    } else {
      this.approved = next;
      this.history.push(m);
      // Any new authored mutation kills the redo trail.
      this.redoStack = [];
    }
    this.working = next;
    return m;
  }

  /**
   * Re-apply every pending agent mutation on top of the current
   * `approved` snapshot. Called after `approveMutation`,
   * `rejectMutation`, `undo`, `redo`, and `rollback` — anything
   * that shifts the floor under the pending stack.
   *
   * Rebase invariants:
   *   - A pending mutation whose handler now throws is flipped to
   *     `"rejected"` with a `rebase-failed` rejection. Callers
   *     (the editor's `subscribe` listener) get notified so the UI
   *     can surface a toast — the previous behaviour silently
   *     dropped the mutation, producing the "agent suggestion just
   *     vanished after Cmd+Z" class of bug.
   *   - A mutation whose handler still applies cleanly keeps its
   *     `id` and `command`; the `before` / `after` / `diff` fields
   *     are refreshed against the new floor.
   *   - The `working` snapshot reflects the rebased stack so the
   *     editor's render loop sees a consistent doc on the next
   *     subscribe tick.
   */
  private recomputeWorking(): void {
    let snapshot = this.approved;
    const newPending: Mutation<TSnapshot>[] = [];
    const rejected: Mutation<TSnapshot>[] = [];
    for (const m of this.pending) {
      const handler = this.handlers.get(m.command.type);
      if (!handler) {
        newPending.push(m);
        continue;
      }
      try {
        const out = handler.apply(snapshot, m.command.payload, this.ctx);
        const re: Mutation<TSnapshot> = {
          ...m,
          before: snapshot,
          after: out.next,
          diff: out.diff,
        };
        snapshot = out.next;
        newPending.push(re);
      } catch (err) {
        // Capture full rejection details (mirrors the synchronous
        // reject path in `applyOneSync`). The well-known code
        // `rebase-failed` lets editor subscribers pattern-match
        // and surface a toast — see spec/shared/agent-api.md.
        const rej: Mutation<TSnapshot> = {
          ...m,
          status: "rejected",
          rejection: {
            code: err instanceof CommandError ? err.code : "rebase-failed",
            message: err instanceof Error ? err.message : String(err),
          },
        };
        rejected.push(rej);
      }
    }
    this.pending = newPending;
    this.working = snapshot;
    // Move rebase-failed mutations into history so they're
    // auditable (matches `rejectMutation`'s bookkeeping). The
    // previous code kept them in `pending` with `status:
    // "rejected"`, where they accumulated forever and skewed
    // every subsequent rebase pass. Pushing to history makes the
    // status transition observable AND terminal.
    for (const m of rejected) this.history.push(m);
    // Notify AFTER the bus state is fully reconciled so subscribers
    // observing `getWorking()` / `getPending()` see the post-rebase
    // world, not a half-updated one. Notifications fire one per
    // rejected mutation so each can be surfaced individually.
    for (const m of rejected) this.notify(m);
  }

  private normalize(c: Command | CommandLite): Command {
    if ("source" in c && "timestamp" in c && "sessionId" in c) return c as Command;
    const lite = c as CommandLite;
    return {
      type: lite.type,
      payload: lite.payload,
      source: lite.source ?? "system",
      ...(lite.agentId ? { agentId: lite.agentId } : {}),
      timestamp: this.ctx.now(),
      sessionId: this.sessionId,
    };
  }

  private notify(m: Mutation<TSnapshot>): void {
    for (const l of this.listeners) {
      try {
        l(this.working, m);
      } catch {
        /* ignore listener errors */
      }
    }
  }
}
