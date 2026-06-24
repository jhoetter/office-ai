import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { createLocalSessionStore } from "@officeai/agent/session-store";
import { DocxAgent } from "@officeai/docx";
import { PdfAgent } from "@officeai/pdf";
import { PptxAgent } from "@officeai/pptx";
import { XlsxAgent } from "@officeai/xlsx";
import { GET as GET_DOCUMENT } from "./[documentId]/route";
import { POST as REVIEW_CHANGE } from "./[documentId]/changes/[mutationId]/route";
import { POST as EXPORT_DOCUMENT } from "./[documentId]/export/route";
import { GET as PROJECT_DOCUMENT } from "./[documentId]/projection/route";
import { POST as CREATE_DOCUMENT } from "./create/route";
import { POST as IMPORT_DOCUMENT } from "./import/route";
import { GET } from "./route";
import type { WebOfficeFormat } from "@/lib/sessions/web-sessions";

let previousOfficeAiDataDir: string | undefined;
let dataDir: string;

beforeEach(() => {
  previousOfficeAiDataDir = process.env.OFFICEAI_DATA_DIR;
  dataDir = mkdtempSync(join(tmpdir(), "officeai-web-data-"));
  process.env.OFFICEAI_DATA_DIR = dataDir;
});

afterEach(() => {
  if (previousOfficeAiDataDir === undefined) {
    delete process.env.OFFICEAI_DATA_DIR;
  } else {
    process.env.OFFICEAI_DATA_DIR = previousOfficeAiDataDir;
  }
  rmSync(dataDir, { recursive: true, force: true });
});

describe("GET /api/sessions", () => {
  it("returns path-free session and document metadata from the local data-dir", async () => {
    const store = createLocalSessionStore();
    await store.putSession({
      id: "session_1",
      title: "Web visible",
      createdAt: "2026-06-24T10:00:00.000Z",
      updatedAt: "2026-06-24T10:02:00.000Z",
      documentIds: ["doc_1"],
    });
    await store.putDocument(
      {
        id: "doc_1",
        sessionId: "session_1",
        format: "pdf",
        name: "statement.pdf",
        status: "ready",
        sourcePath: "/very/local/statement.pdf",
        createdAt: "2026-06-24T10:00:00.000Z",
        updatedAt: "2026-06-24T10:02:00.000Z",
        revision: 4,
        diagnostics: [{ level: "info", code: "imported", message: "Imported statement.pdf as pdf." }],
        exportHistory: [
          { path: "/very/local/export.pdf", bytes: 42, exportedAt: "2026-06-24T10:02:00.000Z" },
        ],
        pendingChanges: [
          {
            id: "mut_1",
            operation: "pdf:rotate-pages",
            status: "pending",
            source: "agent",
          },
        ],
        commandLog: [
          {
            id: "log_1",
            operation: "pdf:rotate-pages",
            status: "pending",
            stage: "queued",
            source: "agent",
            recordedAt: "2026-06-24T10:01:00.000Z",
          },
        ],
      },
      { originalBytes: Buffer.from("%PDF-original"), workingBytes: Buffer.from("%PDF-working") }
    );

    const response = await GET();
    expect(response.status).toBe(200);
    const payload = (await response.json()) as {
      schema: string;
      sessions: Array<{ sessionId: string; documentCount: number }>;
      documents: Array<{
        documentId: string;
        exportCount: number;
        pendingChangeCount: number;
        commandLogCount: number;
        artifacts: { hasOriginal: boolean; hasWorking: boolean };
      }>;
    };

    expect(payload.schema).toBe("office-ai/web-sessions@1");
    expect(payload.sessions).toEqual([
      {
        sessionId: "session_1",
        title: "Web visible",
        createdAt: "2026-06-24T10:00:00.000Z",
        updatedAt: "2026-06-24T10:02:00.000Z",
        documentCount: 1,
      },
    ]);
    expect(payload.documents[0]).toMatchObject({
      documentId: "doc_1",
      exportCount: 1,
      pendingChangeCount: 1,
      commandLogCount: 1,
      artifacts: { hasOriginal: true, hasWorking: true },
    });
    expect(JSON.stringify(payload)).not.toContain("/very/local");
    expect(JSON.stringify(payload)).not.toContain(dataDir);
  });

  it("reports corrupt store metadata clearly", async () => {
    const store = createLocalSessionStore();
    await store.putSession({
      id: "session_1",
      title: "Corrupt",
      createdAt: "2026-06-24T10:00:00.000Z",
      updatedAt: "2026-06-24T10:00:00.000Z",
      documentIds: [],
    });
    writeFileSync(join(dataDir, "sessions", "session_1", "session.json"), "{not json");

    const response = await GET();
    const payload = (await response.json()) as { code: string; message: string };
    expect(response.status).toBe(409);
    expect(payload.code).toBe("corrupt-session-store");
    expect(payload.message).toContain("Corrupt metadata file");
  });

  it("reports storage adapter diagnostics clearly", async () => {
    rmSync(dataDir, { recursive: true, force: true });
    writeFileSync(dataDir, "not a directory");

    const response = await GET();
    const payload = (await response.json()) as {
      code: string;
      diagnostic: { level: string; code: string; message: string };
    };
    expect(response.status).toBe(500);
    expect(payload.code).toBe("session-store-storage-error");
    expect(payload.diagnostic).toMatchObject({
      level: "error",
      code: "storage-ensure-dir",
    });
  });
});

describe("GET /api/sessions/:documentId", () => {
  it("returns path-free document detail metadata from the local data-dir", async () => {
    const store = createLocalSessionStore();
    await store.putSession({
      id: "session_1",
      title: "Detail visible",
      createdAt: "2026-06-24T10:00:00.000Z",
      updatedAt: "2026-06-24T10:05:00.000Z",
      documentIds: ["doc_1"],
    });
    await store.putDocument(
      {
        id: "doc_1",
        sessionId: "session_1",
        format: "docx",
        name: "proposal.docx",
        status: "ready",
        sourcePath: "/very/local/proposal.docx",
        createdAt: "2026-06-24T10:00:00.000Z",
        updatedAt: "2026-06-24T10:05:00.000Z",
        revision: 7,
        diagnostics: [{ level: "warning", code: "needs-review", message: "Layout needs review." }],
        exportHistory: [
          { path: "/very/local/proposal-export.docx", bytes: 4096, exportedAt: "2026-06-24T10:05:00.000Z" },
        ],
        pendingChanges: [
          {
            id: "mut_1",
            operation: "docx.replace-text",
            status: "pending",
            source: "agent",
            actorId: "assistant",
            timestamp: 1782295440000,
            diff: { path: "/very/local/diff.json", secret: "raw-diff" },
          },
        ],
        commandLog: [
          {
            id: "log_1",
            commandId: "cmd_1",
            operation: "docx.replace-text",
            status: "pending",
            stage: "previewed",
            source: "agent",
            actorId: "assistant",
            recordedAt: "2026-06-24T10:04:00.000Z",
            diff: { path: "/very/local/command-diff.json", secret: "raw-command-diff" },
            diagnostics: [{ level: "info", code: "preview-ready", message: "Preview is ready." }],
          },
        ],
      },
      { originalBytes: Buffer.from("docx-original"), workingBytes: Buffer.from("docx-working") }
    );

    const response = await GET_DOCUMENT(new Request("http://localhost/api/sessions/doc_1"), {
      params: Promise.resolve({ documentId: "doc_1" }),
    });
    expect(response.status).toBe(200);
    const payload = (await response.json()) as {
      schema: string;
      session: { sessionId: string; title: string };
      document: {
        documentId: string;
        exportCount: number;
        exports: Array<{ bytes: number; exportedAt: string }>;
        pendingChanges: Array<{ operation: string; hasDiff: boolean }>;
        commandLog: Array<{ stage: string; hasDiff: boolean; diagnostics: Array<{ code: string }> }>;
      };
    };

    expect(payload.schema).toBe("office-ai/web-document@1");
    expect(payload.session).toMatchObject({ sessionId: "session_1", title: "Detail visible" });
    expect(payload.document).toMatchObject({
      documentId: "doc_1",
      exportCount: 1,
      exports: [{ bytes: 4096, exportedAt: "2026-06-24T10:05:00.000Z" }],
      pendingChanges: [{ operation: "docx.replace-text", hasDiff: true }],
      commandLog: [{ stage: "previewed", hasDiff: true, diagnostics: [{ code: "preview-ready" }] }],
    });
    expect(JSON.stringify(payload)).not.toContain("/very/local");
    expect(JSON.stringify(payload)).not.toContain(dataDir);
    expect(JSON.stringify(payload)).not.toContain("raw-diff");
    expect(JSON.stringify(payload)).not.toContain("raw-command-diff");
  });

  it("reports a missing document as not found", async () => {
    const response = await GET_DOCUMENT(new Request("http://localhost/api/sessions/missing"), {
      params: Promise.resolve({ documentId: "missing" }),
    });
    const payload = (await response.json()) as { code: string; message: string };
    expect(response.status).toBe(404);
    expect(payload.code).toBe("document-not-found");
    expect(payload.message).toContain("missing");
  });
});

describe("POST /api/sessions/:documentId/changes/:mutationId", () => {
  it("approves a persisted MCP-style pending change without returning raw diffs or changing bytes", async () => {
    const store = createLocalSessionStore();
    const workingBytes = await seedReviewDocument(store);

    const response = await REVIEW_CHANGE(
      new Request("http://localhost/api/sessions/doc_review/changes/mut_1", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ decision: "approve" }),
      }),
      { params: Promise.resolve({ documentId: "doc_review", mutationId: "mut_1" }) }
    );
    expect(response.status).toBe(200);
    const payload = (await response.json()) as {
      schema: string;
      decision: string;
      document: {
        pendingChangeCount: number;
        pendingChanges: Array<{ id: string; status: string; diffSummary?: string }>;
        commandLog: Array<{ operation: string; status: string; stage: string; source: string }>;
        diagnostics: Array<{ code: string }>;
      };
    };

    expect(payload.schema).toBe("office-ai/web-change-review@1");
    expect(payload.decision).toBe("approved");
    expect(payload.document.pendingChangeCount).toBe(0);
    expect(payload.document.pendingChanges[0]).toMatchObject({
      id: "mut_1",
      status: "approved",
      diffSummary: "Structured diff available",
    });
    expect(payload.document.commandLog.at(-1)).toMatchObject({
      operation: "docx.replace-text",
      status: "approved",
      stage: "reviewed",
      source: "web",
    });
    expect(payload.document.diagnostics.at(-1)).toMatchObject({ code: "change-approved" });
    expect(JSON.stringify(payload)).not.toContain("raw-diff");
    expect(JSON.stringify(payload)).not.toContain(dataDir);

    const stored = await store.getDocument("doc_review");
    if (!stored) throw new Error("Expected review fixture to be persisted.");
    expect(stored.pendingChanges[0]?.status).toBe("approved");
    expect(Buffer.from(await store.readWorkingBytes(stored)).toString("utf8")).toBe(
      workingBytes.toString("utf8")
    );
  });

  it("rejects a pending change, preserves the rejection in the audit trail and leaves bytes unchanged", async () => {
    const store = createLocalSessionStore();
    const workingBytes = await seedReviewDocument(store);

    const response = await REVIEW_CHANGE(
      new Request("http://localhost/api/sessions/doc_review/changes/mut_1", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ decision: "reject", reason: "wrong customer name" }),
      }),
      { params: Promise.resolve({ documentId: "doc_review", mutationId: "mut_1" }) }
    );
    expect(response.status).toBe(200);
    const payload = (await response.json()) as {
      decision: string;
      document: {
        pendingChangeCount: number;
        pendingChanges: Array<{ status: string; rejection?: { code: string; message: string } }>;
        commandLog: Array<{ status: string; stage: string }>;
        diagnostics: Array<{ code: string; message: string }>;
      };
    };

    expect(payload.decision).toBe("rejected");
    expect(payload.document.pendingChangeCount).toBe(0);
    expect(payload.document.pendingChanges[0]).toMatchObject({
      status: "rejected",
      rejection: { code: "human-rejected", message: "wrong customer name" },
    });
    expect(payload.document.commandLog.at(-1)).toMatchObject({ status: "rejected", stage: "reviewed" });
    expect(payload.document.diagnostics.at(-1)).toMatchObject({
      code: "change-rejected",
      message: "Rejected change mut_1: wrong customer name",
    });

    const stored = await store.getDocument("doc_review");
    if (!stored) throw new Error("Expected review fixture to be persisted.");
    expect(stored.pendingChanges[0]).toMatchObject({
      status: "rejected",
      rejection: { code: "human-rejected", message: "wrong customer name" },
    });
    expect(Buffer.from(await store.readWorkingBytes(stored)).toString("utf8")).toBe(
      workingBytes.toString("utf8")
    );
  });

  it("refuses to review the same change twice", async () => {
    const store = createLocalSessionStore();
    await seedReviewDocument(store);

    const first = await REVIEW_CHANGE(
      new Request("http://localhost/api/sessions/doc_review/changes/mut_1", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ decision: "approve" }),
      }),
      { params: Promise.resolve({ documentId: "doc_review", mutationId: "mut_1" }) }
    );
    expect(first.status).toBe(200);

    const second = await REVIEW_CHANGE(
      new Request("http://localhost/api/sessions/doc_review/changes/mut_1", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ decision: "reject" }),
      }),
      { params: Promise.resolve({ documentId: "doc_review", mutationId: "mut_1" }) }
    );
    const payload = (await second.json()) as { code: string; message: string };
    expect(second.status).toBe(409);
    expect(payload.code).toBe("change-already-reviewed");
    expect(payload.message).toContain("approved");
  });

  it("undoes a persisted web review decision back to pending", async () => {
    const store = createLocalSessionStore();
    const workingBytes = await seedReviewDocument(store);

    const rejected = await REVIEW_CHANGE(
      new Request("http://localhost/api/sessions/doc_review/changes/mut_1", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ decision: "reject", reason: "wrong customer name" }),
      }),
      { params: Promise.resolve({ documentId: "doc_review", mutationId: "mut_1" }) }
    );
    expect(rejected.status).toBe(200);

    const undone = await REVIEW_CHANGE(
      new Request("http://localhost/api/sessions/doc_review/changes/mut_1", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ decision: "undo" }),
      }),
      { params: Promise.resolve({ documentId: "doc_review", mutationId: "mut_1" }) }
    );
    expect(undone.status).toBe(200);
    const payload = (await undone.json()) as {
      decision: string;
      document: {
        pendingChangeCount: number;
        pendingChanges: Array<{ status: string; rejection?: { code: string; message: string } }>;
        commandLog: Array<{ status: string; stage: string }>;
        diagnostics: Array<{ code: string; message: string }>;
      };
    };

    expect(payload.decision).toBe("pending");
    expect(payload.document.pendingChangeCount).toBe(1);
    expect(payload.document.pendingChanges[0]).toMatchObject({ status: "pending" });
    expect(payload.document.pendingChanges[0]?.rejection).toBeUndefined();
    expect(payload.document.commandLog.at(-1)).toMatchObject({
      status: "pending",
      stage: "review-undone",
    });
    expect(payload.document.diagnostics.at(-1)).toMatchObject({
      code: "change-review-undone",
      message: "Moved change mut_1 back to pending review.",
    });

    const stored = await store.getDocument("doc_review");
    if (!stored) throw new Error("Expected review fixture to be persisted.");
    expect(stored.pendingChanges[0]).toMatchObject({ status: "pending" });
    expect(stored.pendingChanges[0]?.rejection).toBeUndefined();
    expect(Buffer.from(await store.readWorkingBytes(stored)).toString("utf8")).toBe(
      workingBytes.toString("utf8")
    );
  });

  it("refuses to undo a change that is already pending", async () => {
    const store = createLocalSessionStore();
    await seedReviewDocument(store);

    const response = await REVIEW_CHANGE(
      new Request("http://localhost/api/sessions/doc_review/changes/mut_1", {
        method: "POST",
        headers: { "content-type": "application/json" },
        body: JSON.stringify({ decision: "undo" }),
      }),
      { params: Promise.resolve({ documentId: "doc_review", mutationId: "mut_1" }) }
    );
    const payload = (await response.json()) as { code: string; message: string };
    expect(response.status).toBe(409);
    expect(payload.code).toBe("change-review-not-reviewed");
    expect(payload.message).toContain("pending");
  });
});

describe("POST /api/sessions/import", () => {
  it("imports an uploaded document into the same local data-dir without returning local paths", async () => {
    const bytes = readFileSync(
      new URL("../../../../../fixtures/docx/synthetic/01-plain-paragraphs.docx", import.meta.url)
    );
    const form = new FormData();
    form.set(
      "file",
      new File([bytes], "upload.docx", {
        type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      })
    );

    const response = await IMPORT_DOCUMENT(
      new Request("http://localhost/api/sessions/import", { method: "POST", body: form })
    );
    expect(response.status).toBe(201);
    const payload = (await response.json()) as {
      schema: string;
      session: { sessionId: string; documentCount: number };
      document: {
        documentId: string;
        format: string;
        name: string;
        artifacts: { hasOriginal: boolean; hasWorking: boolean };
        commandLog: Array<{ operation: string; stage: string; source: string }>;
      };
    };

    expect(payload.schema).toBe("office-ai/web-import@1");
    expect(payload.session.documentCount).toBe(1);
    expect(payload.document).toMatchObject({
      format: "docx",
      name: "upload.docx",
      artifacts: { hasOriginal: true, hasWorking: true },
      commandLog: [{ operation: "import_document", stage: "imported", source: "web" }],
    });

    const store = createLocalSessionStore();
    const documents = await store.listDocuments();
    expect(documents).toHaveLength(1);
    expect(documents[0]?.id).toBe(payload.document.documentId);
    expect(JSON.stringify(payload)).not.toContain(dataDir);
    expect(JSON.stringify(payload)).not.toContain("/fixtures/");
  });

  it("rejects unsupported upload extensions", async () => {
    const form = new FormData();
    form.set("file", new File([Buffer.from("plain")], "notes.txt", { type: "text/plain" }));

    const response = await IMPORT_DOCUMENT(
      new Request("http://localhost/api/sessions/import", { method: "POST", body: form })
    );
    const payload = (await response.json()) as { code: string; message: string };
    expect(response.status).toBe(400);
    expect(payload.code).toBe("unsupported-format");
    expect(payload.message).toContain("notes.txt");
  });
});

describe("POST /api/sessions/create and /api/sessions/:documentId/export", () => {
  for (const format of ["docx", "xlsx", "pptx", "pdf"] as const) {
    it(`creates and exports a blank ${format} document`, async () => {
      const createResponse = await CREATE_DOCUMENT(
        new Request("http://localhost/api/sessions/create", {
          method: "POST",
          headers: { "content-type": "application/json" },
          body: JSON.stringify({ format, name: `blank.${format}` }),
        })
      );
      expect(createResponse.status).toBe(201);
      const created = (await createResponse.json()) as {
        schema: string;
        document: {
          documentId: string;
          format: WebOfficeFormat;
          name: string;
          artifacts: { hasOriginal: boolean; hasWorking: boolean };
        };
      };
      expect(created.schema).toBe("office-ai/web-create@1");
      expect(created.document).toMatchObject({
        format,
        name: `blank.${format}`,
        artifacts: { hasOriginal: false, hasWorking: true },
      });

      const exportResponse = await EXPORT_DOCUMENT(
        new Request(`http://localhost/api/sessions/${created.document.documentId}/export`, {
          method: "POST",
        }),
        { params: Promise.resolve({ documentId: created.document.documentId }) }
      );
      expect(exportResponse.status).toBe(200);
      expect(exportResponse.headers.get("content-disposition")).toContain(`blank.${format}`);
      const bytes = new Uint8Array(await exportResponse.arrayBuffer());
      expect(bytes.byteLength).toBeGreaterThan(0);
      await reopenExportedBytes(format, bytes);

      const store = createLocalSessionStore();
      const stored = await store.getDocument(created.document.documentId);
      expect(stored?.exportHistory).toHaveLength(1);
      expect(stored?.commandLog.at(-1)).toMatchObject({
        operation: "export_document",
        stage: "exported",
        source: "web",
      });
      expect(JSON.stringify(created)).not.toContain(dataDir);
    });
  }
});

describe("GET /api/sessions/:documentId/projection", () => {
  it("returns a path-free projection for a large persisted document", async () => {
    const bytes = readFileSync(
      new URL("../../../../../fixtures/docx/synthetic/05-long-body.docx", import.meta.url)
    );
    const agent = await DocxAgent.fromBuffer(bytes);
    const store = createLocalSessionStore();
    await store.putSession({
      id: "session_projection",
      title: "Projection",
      createdAt: "2026-06-24T13:00:00.000Z",
      updatedAt: "2026-06-24T13:00:00.000Z",
      documentIds: ["doc_projection"],
    });
    await store.putDocument(
      {
        id: "doc_projection",
        sessionId: "session_projection",
        format: "docx",
        name: "long.docx",
        status: "ready",
        sourcePath: "/very/local/long.docx",
        createdAt: "2026-06-24T13:00:00.000Z",
        updatedAt: "2026-06-24T13:00:00.000Z",
        revision: agent.getSnapshot().revision,
        diagnostics: [],
        exportHistory: [
          { path: "/very/local/long-export.docx", bytes: 1234, exportedAt: "2026-06-24T13:01:00.000Z" },
        ],
        pendingChanges: [],
        commandLog: [],
      },
      { originalBytes: bytes, workingBytes: bytes }
    );

    const response = await PROJECT_DOCUMENT(
      new Request(
        `http://localhost/api/sessions/doc_projection/projection?projection=text&revision=${agent.getSnapshot().revision}`
      ),
      { params: Promise.resolve({ documentId: "doc_projection" }) }
    );
    expect(response.status).toBe(200);
    const payload = (await response.json()) as {
      schema: string;
      documentId: string;
      revision: number;
      projection: string;
      content: string;
      document: { documentId: string; exportHistory: Array<{ bytes: number; exportedAt: string }> };
    };

    expect(payload.schema).toBe("office-ai/document-projection@1");
    expect(payload.documentId).toBe("doc_projection");
    expect(payload.document.documentId).toBe("doc_projection");
    expect(payload.revision).toBe(agent.getSnapshot().revision);
    expect(payload.projection).toBe("text");
    expect(payload.content.length).toBeGreaterThan(100);
    expect(payload.document.exportHistory).toEqual([{ bytes: 1234, exportedAt: "2026-06-24T13:01:00.000Z" }]);
    expect(JSON.stringify(payload)).not.toContain("/very/local");
    expect(JSON.stringify(payload)).not.toContain(dataDir);
  });

  it("rejects stale revisions and invalid window parameters", async () => {
    const store = createLocalSessionStore();
    const agent = await DocxAgent.empty();
    const bytes = new Uint8Array(await agent.exportFile());
    await store.putSession({
      id: "session_projection",
      title: "Projection",
      createdAt: "2026-06-24T13:00:00.000Z",
      updatedAt: "2026-06-24T13:00:00.000Z",
      documentIds: ["doc_projection"],
    });
    await store.putDocument(
      {
        id: "doc_projection",
        sessionId: "session_projection",
        format: "docx",
        name: "blank.docx",
        status: "ready",
        createdAt: "2026-06-24T13:00:00.000Z",
        updatedAt: "2026-06-24T13:00:00.000Z",
        revision: agent.getSnapshot().revision,
        diagnostics: [],
        exportHistory: [],
        pendingChanges: [],
        commandLog: [],
      },
      { workingBytes: bytes }
    );

    const stale = await PROJECT_DOCUMENT(
      new Request("http://localhost/api/sessions/doc_projection/projection?revision=999"),
      { params: Promise.resolve({ documentId: "doc_projection" }) }
    );
    expect(stale.status).toBe(409);
    expect(((await stale.json()) as { code: string }).code).toBe("stale-revision");

    const invalidPage = await PROJECT_DOCUMENT(
      new Request("http://localhost/api/sessions/doc_projection/projection?projection=page&page=0"),
      { params: Promise.resolve({ documentId: "doc_projection" }) }
    );
    expect(invalidPage.status).toBe(400);
    expect(((await invalidPage.json()) as { code: string }).code).toBe("invalid-page");

    const missingAnchor = await PROJECT_DOCUMENT(
      new Request("http://localhost/api/sessions/doc_projection/projection?anchor=paragraph:99"),
      { params: Promise.resolve({ documentId: "doc_projection" }) }
    );
    expect(missingAnchor.status).toBe(404);
    expect(((await missingAnchor.json()) as { code: string }).code).toBe("anchor-not-found");
  });
});

async function reopenExportedBytes(format: WebOfficeFormat, bytes: Uint8Array): Promise<void> {
  switch (format) {
    case "docx":
      await DocxAgent.fromBuffer(bytes);
      return;
    case "xlsx":
      await XlsxAgent.fromBuffer(bytes);
      return;
    case "pptx":
      await PptxAgent.fromBuffer(bytes);
      return;
    case "pdf":
      await PdfAgent.fromBuffer(bytes);
      return;
  }
}

async function seedReviewDocument(store: ReturnType<typeof createLocalSessionStore>): Promise<Buffer> {
  const workingBytes = Buffer.from("review-working");
  await store.putSession({
    id: "session_review",
    title: "Review",
    createdAt: "2026-06-24T12:00:00.000Z",
    updatedAt: "2026-06-24T12:00:00.000Z",
    documentIds: ["doc_review"],
  });
  await store.putDocument(
    {
      id: "doc_review",
      sessionId: "session_review",
      format: "docx",
      name: "review.docx",
      status: "ready",
      sourcePath: "/very/local/review.docx",
      createdAt: "2026-06-24T12:00:00.000Z",
      updatedAt: "2026-06-24T12:00:00.000Z",
      revision: 3,
      diagnostics: [],
      exportHistory: [],
      pendingChanges: [
        {
          id: "mut_1",
          operation: "docx.replace-text",
          status: "pending",
          source: "agent",
          actorId: "mcp-test",
          timestamp: 1782302400000,
          diff: { secret: "raw-diff", ops: [{ replace: "Acme" }] },
        },
      ],
      commandLog: [
        {
          id: "log_preview",
          commandId: "cmd_review",
          operation: "docx.replace-text",
          status: "pending",
          stage: "previewed",
          source: "mcp",
          actorId: "mcp-test",
          recordedAt: "2026-06-24T12:00:00.000Z",
          diff: { secret: "raw-command-diff" },
          diagnostics: [{ level: "info", code: "preview-ready", message: "Preview is ready." }],
        },
      ],
    },
    { originalBytes: Buffer.from("review-original"), workingBytes }
  );
  return workingBytes;
}
