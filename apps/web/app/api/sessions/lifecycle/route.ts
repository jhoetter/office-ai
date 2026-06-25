/**
 * POST /api/sessions/lifecycle
 *
 * Session-level rename/delete/duplicate commands for the local OfficeAI
 * workspace. Document detail routes remain document-id based.
 */

import { NextResponse } from "next/server";
import {
  createLocalSessionStore,
  SessionStoreCorruptError,
  SessionStoreStorageError,
} from "@officeai/agent/session-store";
import { toWebDocumentDetailEntry, toWebSessionEntry } from "@/lib/sessions/web-sessions";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

type LifecycleAction = "rename" | "delete" | "duplicate";

export async function POST(request: Request): Promise<NextResponse> {
  try {
    const body = await readBody(request);
    if (!body) return badRequest("invalid-body", "Expected a JSON object.");
    const action = typeof body.action === "string" ? body.action : "";
    if (!isLifecycleAction(action)) {
      return badRequest("unsupported-action", "action must be rename, delete or duplicate.");
    }
    const sessionId = typeof body.sessionId === "string" ? body.sessionId.trim() : "";
    if (!sessionId) return badRequest("missing-session-id", "sessionId is required.");

    const store = createLocalSessionStore();
    if (action === "rename") {
      const title = typeof body.title === "string" ? body.title.trim() : "";
      if (!title) return badRequest("missing-title", "title is required for rename.");
      const session = await store.renameSession(sessionId, title);
      return NextResponse.json({
        schema: "office-ai/web-session-lifecycle@1",
        action,
        session: toWebSessionEntry(session),
        documents: [],
      });
    }

    if (action === "delete") {
      await store.deleteSession(sessionId);
      return NextResponse.json({
        schema: "office-ai/web-session-lifecycle@1",
        action,
        deletedSessionId: sessionId,
        documents: [],
      });
    }

    const title = typeof body.title === "string" ? body.title.trim() : undefined;
    const result = await store.duplicateSession(sessionId, { title });
    return NextResponse.json(
      {
        schema: "office-ai/web-session-lifecycle@1",
        action,
        session: toWebSessionEntry(result.session),
        documents: result.documents.map(toWebDocumentDetailEntry),
      },
      { status: 201 }
    );
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
    if (err instanceof SessionStoreStorageError && err.causeCode === "ENOENT") {
      return NextResponse.json(
        {
          schema: "office-ai/web-sessions-error@1",
          code: "session-not-found",
          message,
        },
        { status: 404 }
      );
    }
    return NextResponse.json(
      {
        schema: "office-ai/web-sessions-error@1",
        code: "session-lifecycle-error",
        message,
      },
      { status: 500 }
    );
  }
}

async function readBody(request: Request): Promise<Record<string, unknown> | null> {
  const raw = await request.json().catch(() => null);
  return raw && typeof raw === "object" && !Array.isArray(raw) ? (raw as Record<string, unknown>) : null;
}

function isLifecycleAction(value: string): value is LifecycleAction {
  return value === "rename" || value === "delete" || value === "duplicate";
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
