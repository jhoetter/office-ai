/**
 * POST /api/sessions/:documentId/export
 *
 * Download the persisted working artifact for a local workspace document and
 * record the export in the same store MCP uses.
 */

import { randomUUID } from "node:crypto";
import { NextResponse } from "next/server";
import { createLocalSessionStore, SessionStoreCorruptError } from "@officeai/agent/session-store";
import { contentDisposition, ensureExtension, mimeForFormat } from "@/lib/sessions/server-documents";

export const runtime = "nodejs";
export const dynamic = "force-dynamic";

export async function POST(
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

    const bytes = await store.readWorkingBytes(document);
    const now = new Date().toISOString();
    const filename = ensureExtension(document.name, document.format);
    const diagnostics = [
      { level: "info" as const, code: "exported", message: `Exported ${filename} as ${document.format}.` },
    ];
    await store.putDocument({
      id: document.id,
      sessionId: document.sessionId,
      format: document.format,
      name: filename,
      status: document.status,
      ...(document.sourcePath ? { sourcePath: document.sourcePath } : {}),
      createdAt: document.createdAt,
      updatedAt: now,
      revision: document.revision,
      diagnostics: [...document.diagnostics, ...diagnostics],
      exportHistory: [
        ...document.exportHistory,
        {
          path: `browser-download:${filename}`,
          bytes: bytes.byteLength,
          exportedAt: now,
        },
      ],
      pendingChanges: document.pendingChanges,
      commandLog: [
        ...document.commandLog,
        {
          id: `log_${randomUUID()}`,
          operation: "export_document",
          status: "applied",
          stage: "exported",
          source: "web",
          recordedAt: now,
          diagnostics,
        },
      ],
    });

    return new NextResponse(Buffer.from(bytes), {
      status: 200,
      headers: {
        "content-type": mimeForFormat(document.format),
        "content-disposition": contentDisposition(filename),
        "x-officeai-document-id": document.id,
        "x-officeai-exported-at": now,
      },
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
        code: "session-export-error",
        message,
      },
      { status: 500 }
    );
  }
}
