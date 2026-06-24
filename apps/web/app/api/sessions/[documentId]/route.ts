/**
 * GET /api/sessions/:documentId
 *
 * Path-free document detail view over the local OfficeAI data-dir. This exposes
 * reviewable workspace metadata, not local artifact paths or raw command diffs.
 */

import { NextResponse } from "next/server";
import { createLocalSessionStore, SessionStoreCorruptError } from "@officeai/agent/session-store";
import { toWebDocumentDetailEntry, toWebSessionEntry } from "@/lib/sessions/web-sessions";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function GET(
  _request: Request,
  context: { readonly params: Promise<{ readonly documentId: string }> }
): Promise<NextResponse> {
  try {
    const { documentId } = await context.params;
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

    const session = await store.getSession(document.sessionId);
    return NextResponse.json({
      schema: "office-ai/web-document@1",
      session: toWebSessionEntry(session),
      document: toWebDocumentDetailEntry(document),
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
