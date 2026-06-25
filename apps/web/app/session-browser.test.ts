import { describe, expect, it } from "vitest";
import {
  documentsForSession,
  sessionBrowserCounts,
  type WebSessionsPayload,
} from "@/lib/sessions/web-sessions";
import { formatParityDiagnostics, formatParityFor } from "@/lib/sessions/format-parity";

const payload: WebSessionsPayload = {
  schema: "office-ai/web-sessions@1",
  sessions: [
    {
      sessionId: "session_a",
      title: "A",
      createdAt: "2026-06-24T10:00:00.000Z",
      updatedAt: "2026-06-24T10:02:00.000Z",
      documentCount: 2,
    },
    {
      sessionId: "session_b",
      title: "B",
      createdAt: "2026-06-24T11:00:00.000Z",
      updatedAt: "2026-06-24T11:01:00.000Z",
      documentCount: 1,
    },
  ],
  documents: [
    {
      documentId: "doc_1",
      sessionId: "session_a",
      format: "docx",
      name: "one.docx",
      status: "ready",
      createdAt: "2026-06-24T10:00:00.000Z",
      updatedAt: "2026-06-24T10:02:00.000Z",
      revision: 1,
      diagnostics: [{ level: "info", code: "imported", message: "Imported." }],
      exportCount: 0,
      pendingChangeCount: 2,
      commandLogCount: 3,
      artifacts: { hasOriginal: true, hasWorking: true },
    },
    {
      documentId: "doc_2",
      sessionId: "session_a",
      format: "xlsx",
      name: "two.xlsx",
      status: "ready",
      createdAt: "2026-06-24T10:00:00.000Z",
      updatedAt: "2026-06-24T10:03:00.000Z",
      revision: 2,
      diagnostics: [],
      exportCount: 1,
      pendingChangeCount: 0,
      commandLogCount: 1,
      artifacts: { hasOriginal: false, hasWorking: true },
    },
    {
      documentId: "doc_3",
      sessionId: "session_b",
      format: "pdf",
      name: "three.pdf",
      status: "error",
      createdAt: "2026-06-24T11:00:00.000Z",
      updatedAt: "2026-06-24T11:01:00.000Z",
      revision: 0,
      diagnostics: [{ level: "error", code: "parse-error", message: "Broken PDF." }],
      exportCount: 0,
      pendingChangeCount: 1,
      commandLogCount: 1,
      artifacts: { hasOriginal: true, hasWorking: false },
    },
  ],
};

describe("session browser model", () => {
  it("groups documents by session", () => {
    expect(documentsForSession(payload, "session_a").map((document) => document.documentId)).toEqual([
      "doc_1",
      "doc_2",
    ]);
    expect(documentsForSession(payload, "session_b").map((document) => document.documentId)).toEqual([
      "doc_3",
    ]);
  });

  it("summarizes sessions, documents, pending changes and diagnostics", () => {
    expect(sessionBrowserCounts(payload)).toEqual({
      sessions: 2,
      documents: 3,
      pending: 3,
      diagnostics: 2,
    });
  });

  it("defines honest web parity diagnostics for each supported format", () => {
    expect(formatParityFor("docx").rows.map((row) => row.label)).toEqual([
      "Import",
      "Read view",
      "Editable structures",
      "Review and diff",
      "Export",
    ]);
    expect(formatParityDiagnostics("docx")[0].code).toBe("web-parity-docx-partial-edit");
    expect(formatParityDiagnostics("xlsx")[0].code).toBe("web-parity-xlsx-partial-edit");
    expect(formatParityDiagnostics("pptx")[0].code).toBe("web-parity-pptx-partial-edit");
    expect(formatParityDiagnostics("pdf")[0]).toMatchObject({
      level: "warning",
      code: "web-parity-pdf-review-only",
    });
  });
});
