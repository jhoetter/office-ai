import { v4 as uuidv4 } from "uuid";
import { defaultIdMinter, type IdMinter } from "../util/ids.js";
import type { DocumentDiff, DocumentFormat, DocumentSnapshot } from "../types/document.js";
import { emptyDiff } from "../types/document.js";
import type { CommandHandler, CommandSource, HandlerContext, Mutation } from "./types.js";
import { CommandBus } from "./bus.js";
import { CommandError } from "./types.js";

export type CommandSurface = "mcp" | "web" | "cli" | "internal";
export type CommandPolicyMode = "dry_run" | "auto_apply" | "pending";
export type CommandLifecycleStage =
  | "created"
  | "validated"
  | "previewed"
  | "applied"
  | "queued"
  | "reviewed"
  | "exported"
  | "failed";

export type CommandDiagnosticLevel = "info" | "warning" | "error" | "destructive";

export interface CommandDiagnostic {
  readonly level: CommandDiagnosticLevel;
  readonly code: string;
  readonly message: string;
  readonly path?: ReadonlyArray<string | number>;
}

export interface CommandTarget {
  readonly sessionId: string;
  readonly documentId: string;
  readonly revision: number;
  readonly anchor?: unknown;
}

export interface CommandEnvelope<TPayload = unknown> {
  readonly id: string;
  readonly format: DocumentFormat;
  readonly operation: string;
  readonly arguments: TPayload;
  readonly target: CommandTarget;
  readonly source: {
    readonly surface: CommandSurface;
    readonly actorId?: string;
  };
  readonly policy: {
    readonly mode: CommandPolicyMode;
    readonly requiresReview: boolean;
  };
  readonly createdAt: number;
}

export interface CreateCommandEnvelopeInput<TPayload = unknown> {
  readonly id?: string;
  readonly format: DocumentFormat;
  readonly operation: string;
  readonly arguments: TPayload;
  readonly target: CommandTarget;
  readonly source: CommandEnvelope["source"];
  readonly policy?: Partial<CommandEnvelope["policy"]>;
  readonly createdAt?: number;
}

export interface CommandLifecycleResult<TSnapshot extends DocumentSnapshot = DocumentSnapshot> {
  readonly envelope: CommandEnvelope;
  readonly stage: CommandLifecycleStage;
  readonly ok: boolean;
  readonly diagnostics: ReadonlyArray<CommandDiagnostic>;
  readonly diff?: DocumentDiff;
  readonly mutation?: Mutation<TSnapshot>;
}

export interface PreviewCommandOptions {
  readonly now?: () => number;
  readonly mintNodeId?: IdMinter;
}

export function createCommandEnvelope<TPayload>(
  input: CreateCommandEnvelopeInput<TPayload>
): CommandEnvelope<TPayload> {
  const mode = input.policy?.mode ?? "pending";
  return {
    id: input.id ?? uuidv4(),
    format: input.format,
    operation: input.operation,
    arguments: input.arguments,
    target: input.target,
    source: input.source,
    policy: {
      mode,
      requiresReview: input.policy?.requiresReview ?? mode !== "dry_run",
    },
    createdAt: input.createdAt ?? Date.now(),
  };
}

export function validateCommandEnvelope<TSnapshot extends DocumentSnapshot>(
  envelope: CommandEnvelope,
  snapshot: TSnapshot
): CommandLifecycleResult<TSnapshot> {
  const diagnostics: CommandDiagnostic[] = [];
  if (envelope.format !== snapshot.format) {
    diagnostics.push({
      level: "error",
      code: "format-mismatch",
      message: `Command targets ${envelope.format}, but the current document is ${snapshot.format}.`,
    });
  }
  if (envelope.target.revision !== snapshot.revision) {
    diagnostics.push({
      level: "error",
      code: "stale-revision",
      message: `Command targets revision ${envelope.target.revision}, but the current document is revision ${snapshot.revision}.`,
    });
  }
  const ok = !hasBlockingDiagnostics(diagnostics);
  return {
    envelope,
    stage: ok ? "validated" : "failed",
    ok,
    diagnostics,
  };
}

export function previewCommandEnvelope<TPayload, TSnapshot extends DocumentSnapshot>(
  envelope: CommandEnvelope<TPayload>,
  snapshot: TSnapshot,
  handler: CommandHandler<TPayload, TSnapshot>,
  opts: PreviewCommandOptions = {}
): CommandLifecycleResult<TSnapshot> {
  const validation = validateCommandEnvelope(envelope, snapshot);
  if (!validation.ok) return validation;
  try {
    const out = handler.apply(snapshot, envelope.arguments, handlerContext(opts));
    return {
      envelope,
      stage: "previewed",
      ok: true,
      diagnostics: [],
      diff: out.diff,
    };
  } catch (err) {
    return failedLifecycle(envelope, err, snapshot);
  }
}

export async function applyCommandEnvelope<TPayload, TSnapshot extends DocumentSnapshot>(
  bus: CommandBus<TSnapshot>,
  envelope: CommandEnvelope<TPayload>
): Promise<CommandLifecycleResult<TSnapshot>> {
  const validation = validateCommandEnvelope(envelope, bus.getSnapshot());
  if (!validation.ok) return validation;
  if (envelope.policy.mode === "dry_run") {
    return {
      envelope,
      stage: "failed",
      ok: false,
      diagnostics: [
        {
          level: "error",
          code: "dry-run-apply",
          message: "Dry-run commands must be previewed; apply requires policy.mode auto_apply or pending.",
        },
      ],
    };
  }

  try {
    const mutation = await bus.dispatch({
      type: envelope.operation,
      payload: envelope.arguments,
      source: commandSourceForEnvelope(envelope),
      ...(envelope.source.actorId ? { agentId: envelope.source.actorId } : {}),
    });
    if (envelope.policy.mode === "auto_apply" && mutation.status === "pending") {
      bus.approveMutation(mutation.id);
    }
    const finalStatus = mutation.status;
    const ok = finalStatus !== "rejected";
    return {
      envelope,
      stage: finalStatus === "pending" ? "queued" : ok ? "applied" : "failed",
      ok,
      diagnostics: mutation.rejection
        ? [
            {
              level: "error",
              code: mutation.rejection.code,
              message: mutation.rejection.message,
            },
          ]
        : [],
      diff: mutation.diff,
      mutation,
    };
  } catch (err) {
    return failedLifecycle(envelope, err, bus.getSnapshot());
  }
}

export function hasBlockingDiagnostics(diagnostics: ReadonlyArray<CommandDiagnostic>): boolean {
  return diagnostics.some((d) => d.level === "error" || d.level === "destructive");
}

function commandSourceForEnvelope(envelope: CommandEnvelope): CommandSource {
  if (envelope.policy.mode === "pending") return "agent";
  if (envelope.source.surface === "internal") return "system";
  if (envelope.source.surface === "mcp") return "agent";
  return "human";
}

function handlerContext(opts: PreviewCommandOptions): HandlerContext {
  return {
    mintNodeId: opts.mintNodeId ?? defaultIdMinter,
    now: opts.now ?? (() => Date.now()),
  };
}

function failedLifecycle<TSnapshot extends DocumentSnapshot>(
  envelope: CommandEnvelope,
  err: unknown,
  snapshot: TSnapshot
): CommandLifecycleResult<TSnapshot> {
  const code = err instanceof CommandError ? err.code : "handler-threw";
  const message = err instanceof Error ? err.message : String(err);
  return {
    envelope,
    stage: "failed",
    ok: false,
    diagnostics: [{ level: "error", code, message }],
    diff: emptyDiff(snapshot.format, snapshot.revision, snapshot.revision),
  };
}
