/**
 * POST /api/sessions/:documentId/changes/:mutationId
 *
 * Metadata-level review for persisted pending changes. This does not replay a
 * live command diff after restart; it records the human decision and preserves
 * the audit trail in the same local store MCP writes.
 */

import { randomUUID } from "node:crypto";
import { NextResponse } from "next/server";
import { createLocalSessionStore, SessionStoreCorruptError } from "@officeai/agent/session-store";
import { toWebDocumentDetailEntry, toWebSessionEntry } from "@/lib/sessions/web-sessions";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type ReviewDecision = "approved" | "rejected";

export async function POST(
  request: Request,
  context: { readonly params: Promise<{ readonly documentId: string; readonly mutationId: string }> }
): Promise<NextResponse> {
  try {
    const { documentId, mutationId } = await context.params;
    const rawBody = await request.json().catch(() => ({}));
    const body =
      rawBody && typeof rawBody === "object" && !Array.isArray(rawBody)
        ? (rawBody as Record<string, unknown>)
        : {};
    const decision = reviewDecision(body.decision);
    if (!decision) {
      return badRequest("invalid-review-decision", "decision must be approve, approved, reject or rejected.");
    }

    const store = createLocalSessionStore();
    const document = await store.getDocument(documentId);
    if (!document) {
      return NextResponse.json(
        {
          schema: "office-ai/web-sessions-error@1",
          code: "document-not-found",
          message: `Document ${documentId} was not found in the local workspace.`,
        },
        { status: 404 }
      );
    }
    const target = document.pendingChanges.find((change) => change.id === mutationId);
    if (!target) {
      return NextResponse.json(
        {
          schema: "office-ai/web-sessions-error@1",
          code: "change-not-found",
          message: `Change ${mutationId} was not found for document ${documentId}.`,
        },
        { status: 404 }
      );
    }
    if (target.status !== "pending") {
      return NextResponse.json(
        {
          schema: "office-ai/web-sessions-error@1",
          code: "change-already-reviewed",
          message: `Change ${mutationId} is already ${target.status}.`,
        },
        { status: 409 }
      );
    }

    const now = new Date().toISOString();
    const reason = typeof body.reason === "string" && body.reason.trim() ? body.reason.trim() : undefined;
    const diagnostics = [
      {
        level: "info" as const,
        code: decision === "approved" ? "change-approved" : "change-rejected",
        message:
          decision === "approved"
            ? `Approved change ${mutationId}.`
            : reason
              ? `Rejected change ${mutationId}: ${reason}`
              : `Rejected change ${mutationId}.`,
      },
    ];
    const reviewed = {
      ...target,
      status: decision,
      ...(decision === "rejected"
        ? { rejection: { code: "human-rejected", message: reason ?? "Rejected in web review." } }
        : {}),
    };
    const updatedDocument = await store.putDocument({
      id: document.id,
      sessionId: document.sessionId,
      format: document.format,
      name: document.name,
      status: document.status,
      ...(document.sourcePath ? { sourcePath: document.sourcePath } : {}),
      createdAt: document.createdAt,
      updatedAt: now,
      revision: document.revision,
      diagnostics: [...document.diagnostics, ...diagnostics],
      exportHistory: document.exportHistory,
      pendingChanges: document.pendingChanges.map((change) => (change.id === mutationId ? reviewed : change)),
      commandLog: [
        ...document.commandLog,
        {
          id: `log_${randomUUID()}`,
          operation: target.operation,
          status: decision,
          stage: "reviewed",
          source: "web",
          ...(target.actorId ? { actorId: target.actorId } : {}),
          recordedAt: now,
          diagnostics,
        },
      ],
    });
    const session = await store.getSession(document.sessionId);

    return NextResponse.json({
      schema: "office-ai/web-change-review@1",
      decision,
      mutationId,
      session: toWebSessionEntry(session),
      document: toWebDocumentDetailEntry(updatedDocument),
    });
  } catch (err) {
    const message = err instanceof Error ? err.message : String(err);
    if (err instanceof SessionStoreCorruptError) {
      return NextResponse.json(
        {
          schema: "office-ai/web-sessions-error@1",
          code: "corrupt-session-store",
          message,
        },
        { status: 409 }
      );
    }
    return NextResponse.json(
      {
        schema: "office-ai/web-sessions-error@1",
        code: "change-review-error",
        message,
      },
      { status: 500 }
    );
  }
}

function badRequest(code: string, message: string): NextResponse {
  return NextResponse.json(
    {
      schema: "office-ai/web-sessions-error@1",
      code,
      message,
    },
    { status: 400 }
  );
}

function reviewDecision(value: unknown): ReviewDecision | null {
  if (value === "approve" || value === "approved") return "approved";
  if (value === "reject" || value === "rejected") return "rejected";
  return null;
}
