import type { DocumentDiff, NodeId } from "../types/document.js";
import type { IdMinter } from "../util/ids.js";

export type CommandSource = "human" | "agent" | "system";

export interface Command<TType extends string = string, TPayload = unknown> {
  readonly type: TType;
  readonly payload: TPayload;
  readonly source: CommandSource;
  readonly agentId?: string;
  readonly timestamp: number;
  readonly sessionId: string;
}

export interface CommandLite<TType extends string = string, TPayload = unknown> {
  readonly type: TType;
  readonly payload: TPayload;
  readonly source?: CommandSource;
  readonly agentId?: string;
}

export type MutationStatus =
  | "approved"
  | "pending"
  | "rejected"
  | "rolled-back";

export interface Mutation<TSnapshot = unknown> {
  readonly id: string;
  readonly command: Command;
  readonly before: TSnapshot;
  readonly after: TSnapshot;
  readonly diff: DocumentDiff;
  status: MutationStatus;
  /** Optional rejection reason — populated when status === 'rejected'. */
  rejection?: { code: string; message: string };
}

export interface HandlerContext {
  readonly mintNodeId: IdMinter;
  readonly now: () => number;
}

export interface CommandHandler<TPayload = unknown, TSnapshot = unknown> {
  readonly type: string;
  apply(
    snapshot: TSnapshot,
    payload: TPayload,
    ctx: HandlerContext,
  ): { next: TSnapshot; diff: DocumentDiff };
}

export class CommandError extends Error {
  readonly code: string;
  constructor(code: string, message: string, opts?: { cause?: unknown }) {
    super(message);
    this.name = "CommandError";
    this.code = code;
    if (opts?.cause !== undefined) (this as unknown as { cause: unknown }).cause = opts.cause;
  }
}

export class NotImplementedError extends CommandError {
  constructor(commandType: string, opts?: { reason?: string }) {
    super("not-implemented", `${commandType} is not implemented yet${opts?.reason ? `: ${opts.reason}` : ""}`);
    this.name = "NotImplementedError";
  }
}

export type { NodeId };
