/**
 * GET /api/sessions
 *
 * Read-only inspector over the local OfficeAI data-dir. MCP tools and the web
 * editor use the same store, but this route deliberately returns path-free
 * metadata so local file system details do not leak into browser payloads.
 */

import { NextResponse } from "next/server";
import {
  createLocalSessionStore,
  SessionStoreCorruptError,
  SessionStoreStorageError,
} from "@officeai/agent/session-store";
import { toWebDocumentEntry, toWebSessionEntry } from "@/lib/sessions/web-sessions";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(): Promise<NextResponse> {
  try {
    const store = createLocalSessionStore();
    const [sessions, documents] = await Promise.all([store.listSessions(), store.listDocuments()]);
    return NextResponse.json({
      schema: "office-ai/web-sessions@1",
      sessions: sessions.map(toWebSessionEntry),
      documents: documents.map(toWebDocumentEntry),
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
    if (err instanceof SessionStoreStorageError) {
      return NextResponse.json(
        {
          schema: "office-ai/web-sessions-error@1",
          code: "session-store-storage-error",
          message,
          diagnostic: err.diagnostic,
        },
        { status: 500 }
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
