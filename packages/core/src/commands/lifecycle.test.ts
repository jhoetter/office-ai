import { describe, expect, it } from "vitest";
import type { DocumentSnapshot } from "../types/document.js";
import { CommandBus } from "./bus.js";
import { CommandError, type CommandHandler } from "./types.js";
import {
  applyCommandEnvelope,
  createCommandEnvelope,
  previewCommandEnvelope,
  validateCommandEnvelope,
} from "./lifecycle.js";
import { resolveReviewPolicy } from "./review-policy.js";

interface ToySnapshot extends DocumentSnapshot<{ value: number }> {
  format: "docx";
}

const initial: ToySnapshot = {
  format: "docx",
  revision: 0,
  root: { value: 0 },
  partHashes: {},
};

const incHandler: CommandHandler<{ by: number }, ToySnapshot> = {
  type: "docx:inc",
  apply(snapshot, payload) {
    return {
      next: {
        ...snapshot,
        revision: snapshot.revision + 1,
        root: { value: snapshot.root.value + payload.by },
      },
      diff: {
        format: "docx",
        fromRevision: snapshot.revision,
        toRevision: snapshot.revision + 1,
        changes: [
          {
            kind: "node-updated",
            nodeId: "root",
            path: ["root", "value"],
            field: "value",
            summary: `+${payload.by}`,
          },
        ],
      },
    };
  },
};

const anchoredHandler: CommandHandler<{ anchor: string; by: number }, ToySnapshot> = {
  type: "docx:anchored-inc",
  apply(snapshot, payload) {
    if (payload.anchor !== "root") {
      throw new CommandError("invalid-anchor", `Unknown anchor ${payload.anchor}`);
    }
    return incHandler.apply(snapshot, { by: payload.by });
  },
};

function envelope(overrides: Partial<Parameters<typeof createCommandEnvelope>[0]> = {}) {
  return createCommandEnvelope({
    id: "cmd-1",
    format: "docx",
    operation: "docx:inc",
    arguments: { by: 2 },
    target: { sessionId: "s-1", documentId: "d-1", revision: 0 },
    source: { surface: "mcp", actorId: "agent-1" },
    policy: { mode: "auto_apply", requiresReview: false },
    createdAt: 123,
    ...overrides,
  });
}

describe("command lifecycle contract", () => {
  it("creates an explicit command envelope", () => {
    const env = envelope();

    expect(env.id).toBe("cmd-1");
    expect(env.format).toBe("docx");
    expect(env.operation).toBe("docx:inc");
    expect(env.arguments).toEqual({ by: 2 });
    expect(env.source).toEqual({ surface: "mcp", actorId: "agent-1" });
    expect(env.policy).toEqual({ mode: "auto_apply", requiresReview: false });
    expect(env.target.revision).toBe(0);
  });

  it("previews without mutating the command bus", () => {
    const bus = new CommandBus<ToySnapshot>(initial);
    bus.register(incHandler);

    const result = previewCommandEnvelope(envelope(), bus.getSnapshot(), incHandler);

    expect(result.ok).toBe(true);
    expect(result.stage).toBe("previewed");
    expect(result.diff?.toRevision).toBe(1);
    expect(bus.getSnapshot().revision).toBe(0);
    expect(bus.getSnapshot().root.value).toBe(0);
    expect(bus.getHistory()).toHaveLength(0);
  });

  it("applies auto-apply commands and records the mutation", async () => {
    const bus = new CommandBus<ToySnapshot>(initial);
    bus.register(incHandler);

    const result = await applyCommandEnvelope(bus, envelope());

    expect(result.ok).toBe(true);
    expect(result.stage).toBe("applied");
    expect(result.mutation?.command.source).toBe("agent");
    expect(result.mutation?.status).toBe("approved");
    expect(bus.getApproved().root.value).toBe(2);
    expect(bus.getPending()).toHaveLength(0);
  });

  it("queues pending commands without approving them", async () => {
    const bus = new CommandBus<ToySnapshot>(initial);
    bus.register(incHandler);

    const result = await applyCommandEnvelope(
      bus,
      envelope({ policy: { mode: "pending", requiresReview: true } })
    );

    expect(result.ok).toBe(true);
    expect(result.stage).toBe("queued");
    expect(result.mutation?.status).toBe("pending");
    expect(bus.getApproved().root.value).toBe(0);
    expect(bus.getWorking().root.value).toBe(2);
    expect(bus.getPending()).toHaveLength(1);
  });

  it("fails stale revisions before preview or apply", async () => {
    const stale = envelope({ target: { sessionId: "s-1", documentId: "d-1", revision: 9 } });
    const bus = new CommandBus<ToySnapshot>(initial);
    bus.register(incHandler);

    const validation = validateCommandEnvelope(stale, bus.getSnapshot());
    const applied = await applyCommandEnvelope(bus, stale);

    expect(validation.ok).toBe(false);
    expect(validation.diagnostics[0]?.code).toBe("stale-revision");
    expect(applied.ok).toBe(false);
    expect(applied.diagnostics[0]?.code).toBe("stale-revision");
    expect(bus.getSnapshot().revision).toBe(0);
  });

  it("keeps dry-run and apply as separate lifecycle paths", async () => {
    const bus = new CommandBus<ToySnapshot>(initial);
    bus.register(incHandler);

    const result = await applyCommandEnvelope(
      bus,
      envelope({ policy: { mode: "dry_run", requiresReview: false } })
    );

    expect(result.ok).toBe(false);
    expect(result.diagnostics[0]?.code).toBe("dry-run-apply");
    expect(bus.getHistory()).toHaveLength(0);
    expect(bus.getSnapshot().revision).toBe(0);
  });

  it("returns handler diagnostics for invalid anchors", () => {
    const result = previewCommandEnvelope(
      envelope({
        operation: "docx:anchored-inc",
        arguments: { anchor: "missing", by: 2 },
      }),
      initial,
      anchoredHandler
    );

    expect(result.ok).toBe(false);
    expect(result.stage).toBe("failed");
    expect(result.diagnostics[0]).toMatchObject({
      level: "error",
      code: "invalid-anchor",
    });
    expect(result.diff?.fromRevision).toBe(0);
    expect(result.diff?.toRevision).toBe(0);
  });
});

describe("review policy", () => {
  it("allows explicitly safe auto-apply commands", () => {
    const policy = resolveReviewPolicy({
      operation: "docx:insert-text",
      requestedMode: "auto_apply",
      requestedRequiresReview: false,
      actionRequiresReview: false,
      sourceSurface: "mcp",
    });

    expect(policy.mode).toBe("auto_apply");
    expect(policy.requiresReview).toBe(false);
    expect(policy.diagnostics).toEqual([]);
  });

  it("downgrades catalogue-required review from auto_apply to pending", () => {
    const policy = resolveReviewPolicy({
      operation: "pptx:set-slide-size",
      requestedMode: "auto_apply",
      requestedRequiresReview: false,
      actionRequiresReview: true,
      sourceSurface: "mcp",
    });

    expect(policy.mode).toBe("pending");
    expect(policy.requiresReview).toBe(true);
    expect(policy.diagnostics.map((d) => d.code)).toEqual([
      "catalog-review-required",
      "review-opt-out-ignored",
      "auto-apply-downgraded-to-pending",
    ]);
  });

  it("requires review for destructive operations even without catalogue metadata", () => {
    const policy = resolveReviewPolicy({
      operation: "pdf:delete-pages",
      requestedMode: "auto_apply",
      requestedRequiresReview: false,
      sourceSurface: "mcp",
    });

    expect(policy.mode).toBe("pending");
    expect(policy.requiresReview).toBe(true);
    expect(policy.diagnostics.map((d) => d.code)).toEqual([
      "destructive-command-review-required",
      "review-opt-out-ignored",
      "auto-apply-downgraded-to-pending",
    ]);
  });
});
