/**
 * GET /api/sessions
 *
 * Read-only inspector over the local OfficeAI data-dir. MCP tools and the web
 * editor use the same store, but this route deliberately returns path-free
 * metadata so local file system details do not leak into browser payloads.
 */

import { NextResponse } from "next/server";
import { createLocalSessionStore, SessionStoreCorruptError } from "@officeai/agent/session-store";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(): Promise<NextResponse> {
  try {
    const store = createLocalSessionStore();
    const [sessions, documents] = await Promise.all([store.listSessions(), store.listDocuments()]);
    return NextResponse.json({
      schema: "office-ai/web-sessions@1",
      sessions: sessions.map((session) => ({
        sessionId: session.id,
        title: session.title,
        createdAt: session.createdAt,
        updatedAt: session.updatedAt,
        documentCount: session.documentIds.length,
      })),
      documents: documents.map((document) => ({
        documentId: document.id,
        sessionId: document.sessionId,
        format: document.format,
        name: document.name,
        status: document.status,
        createdAt: document.createdAt,
        updatedAt: document.updatedAt,
        revision: document.revision,
        diagnostics: document.diagnostics,
        exportCount: document.exportHistory.length,
        pendingChangeCount: document.pendingChanges.length,
        commandLogCount: document.commandLog.length,
        artifacts: {
          hasOriginal: Boolean(document.artifacts.originalPath),
          hasWorking: Boolean(document.artifacts.workingPath),
        },
      })),
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
        code: "session-store-error",
        message,
      },
      { status: 500 }
    );
  }
}
