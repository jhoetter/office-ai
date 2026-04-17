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
    return this.pending;
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
    }
    this.working = next;
    return m;
  }

  private recomputeWorking(): void {
    let snapshot = this.approved;
    const newPending: Mutation<TSnapshot>[] = [];
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
      } catch {
        newPending.push({ ...m, status: "rejected" });
      }
    }
    this.pending = newPending;
    this.working = snapshot;
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
