import { mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { dirname, join, resolve, resolve as resolvePath } from "node:path";
import { fileURLToPath } from "node:url";
import { afterEach, beforeEach, describe, expect, it } from "vitest";
import { Document, HeadingLevel, Packer, Paragraph, TextRun } from "docx";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import { InMemoryTransport } from "@modelcontextprotocol/sdk/inMemory.js";
import { XlsxAgent } from "@officeai/xlsx";
import { createMcpServer, __resetMcpSessionsForTests } from "./mcp.js";
import { runCli } from "./cli.js";
import { fixturePath, requiredMatrixFixture, type FixtureFormat } from "../../../tests/fixture-matrix.js";

const here = dirname(fileURLToPath(import.meta.url));
const xlsxFixtures = resolvePath(here, "../../../fixtures/xlsx/synthetic");
let previousOfficeAiDataDir: string | undefined;
let activeOfficeAiDataDir: string | undefined;

beforeEach(() => {
  previousOfficeAiDataDir = process.env.OFFICEAI_DATA_DIR;
  activeOfficeAiDataDir = mkdtempSync(join(tmpdir(), "officeai-mcp-data-"));
  process.env.OFFICEAI_DATA_DIR = activeOfficeAiDataDir;
  __resetMcpSessionsForTests();
});

afterEach(() => {
  __resetMcpSessionsForTests();
  if (previousOfficeAiDataDir === undefined) {
    delete process.env.OFFICEAI_DATA_DIR;
  } else {
    process.env.OFFICEAI_DATA_DIR = previousOfficeAiDataDir;
  }
  if (activeOfficeAiDataDir) {
    rmSync(activeOfficeAiDataDir, { recursive: true, force: true });
    activeOfficeAiDataDir = undefined;
  }
});

function copyXlsxFixture(name: string, dest: string): string {
  const target = join(dest, name);
  writeFileSync(target, readFileSync(resolvePath(xlsxFixtures, name)));
  return target;
}

async function makeFixture(): Promise<string> {
  const dir = mkdtempSync(join(tmpdir(), "mcp-fixture-"));
  const doc = new Document({
    creator: "test",
    sections: [
      {
        properties: {},
        children: [
          new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun("Hello")] }),
          new Paragraph({ children: [new TextRun("first paragraph body")] }),
          new Paragraph({ children: [new TextRun("second paragraph body")] }),
        ],
      },
    ],
  });
  const buf = await Packer.toBuffer(doc);
  const inputPath = join(dir, "in.docx");
  writeFileSync(inputPath, buf);
  return inputPath;
}

async function makeClient(): Promise<Client> {
  const server = createMcpServer();
  const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
  const client = new Client({ name: "officeai-test", version: "0.0.0" }, { capabilities: {} });
  await Promise.all([client.connect(clientTransport), server.connect(serverTransport)]);
  return client;
}

function structured(result: unknown): Record<string, unknown> {
  if (!result || typeof result !== "object") throw new Error("expected structured tool result");
  const r = result as { structuredContent?: unknown; content?: unknown; isError?: boolean };
  if (r.isError) {
    throw new Error(`tool returned error: ${JSON.stringify(r.content)}`);
  }
  if (r.structuredContent && typeof r.structuredContent === "object") {
    return r.structuredContent as Record<string, unknown>;
  }
  if (Array.isArray(r.content) && r.content.length > 0) {
    const first = r.content[0] as { text?: string };
    if (first.text) return JSON.parse(first.text);
  }
  throw new Error("tool result had neither structuredContent nor parseable text");
}

class CapturedStream {
  chunks: string[] = [];
  write(s: string | Uint8Array): boolean {
    this.chunks.push(typeof s === "string" ? s : Buffer.from(s).toString("utf8"));
    return true;
  }
  text(): string {
    return this.chunks.join("");
  }
}

function makeIO() {
  const stdout = new CapturedStream();
  const stderr = new CapturedStream();
  return {
    io: {
      stdout: stdout as unknown as NodeJS.WritableStream,
      stderr: stderr as unknown as NodeJS.WritableStream,
    },
    stdout,
    stderr,
  };
}

function matrixPath(format: FixtureFormat): string {
  return fixturePath(
    requiredMatrixFixture(format, {
      complexity: "simple",
      expectedBehavior: "import",
    })
  );
}

describe("OfficeAI MCP server", () => {
  beforeEach(() => __resetMcpSessionsForTests());
  afterEach(() => __resetMcpSessionsForTests());

  it("registers the hand-rolled docx_*, xlsx_*, pptx_* and pdf_* tools", async () => {
    const client = await makeClient();
    const list = await client.listTools();
    const names = new Set(list.tools.map((t) => t.name));
    // Hand-rolled, stable surface — additions/removals here are
    // breaking changes for MCP clients and should be reviewed.
    const handRolled = [
      "docx_apply_command",
      "docx_approve",
      "docx_diff",
      "docx_get_page_text",
      "docx_get_pages",
      "docx_get_text",
      "docx_inspect",
      "docx_list_pending",
      "docx_load",
      "docx_reject",
      "docx_save",
      "docx_search",
      "pdf_add_page_numbers",
      "pdf_add_watermark",
      "pdf_annotations",
      "pdf_delete_pages",
      "pdf_export_markdown",
      "pdf_extract_pages",
      "pdf_fill_form",
      "pdf_flatten_form",
      "pdf_list_form_fields",
      "pdf_load",
      "pdf_merge",
      "pdf_metadata",
      "pdf_outline",
      "pdf_read_page",
      "pdf_reorder_pages",
      "pdf_reset_form",
      "pdf_rotate_pages",
      "pdf_search",
      "pdf_set_metadata",
      "pptx_apply_command",
      "pptx_diff",
      "pptx_get_text",
      "pptx_inspect",
      "pptx_load",
      "pptx_save",
      "pptx_search",
      "xlsx_add_comment",
      "xlsx_add_sheet",
      "xlsx_apply_command",
      "xlsx_approve",
      "xlsx_delete_column",
      "xlsx_delete_row",
      "xlsx_diff",
      "xlsx_get_text",
      "xlsx_insert_column",
      "xlsx_insert_row",
      "xlsx_inspect",
      "xlsx_list_pending",
      "xlsx_list_sheets",
      "xlsx_load",
      "xlsx_merge",
      "xlsx_redo",
      "xlsx_reject",
      "xlsx_rename_sheet",
      "xlsx_save",
      "xlsx_search",
      "xlsx_set_cell",
      "xlsx_set_format",
      "xlsx_set_formula",
      "xlsx_set_range",
      "xlsx_undo",
      "xlsx_unmerge",
    ];
    for (const expected of handRolled) {
      expect(names.has(expected), `missing hand-rolled tool ${expected}`).toBe(true);
    }
  });

  it("registers canonical session/document tools", async () => {
    const client = await makeClient();
    const list = await client.listTools();
    const names = new Set(list.tools.map((t) => t.name));
    for (const expected of [
      "create_session",
      "list_sessions",
      "import_document",
      "create_document",
      "list_documents",
      "get_document",
      "get_document_projection",
      "export_document",
    ]) {
      expect(names.has(expected), `missing canonical tool ${expected}`).toBe(true);
    }
  });

  it("registers canonical command lifecycle tools", async () => {
    const client = await makeClient();
    const list = await client.listTools();
    const names = new Set(list.tools.map((t) => t.name));
    for (const expected of [
      "plan_command",
      "preview_command",
      "apply_command",
      "undo_command",
      "list_pending_changes",
      "approve_change",
      "reject_change",
    ]) {
      expect(names.has(expected), `missing command lifecycle tool ${expected}`).toBe(true);
    }
  });

  it("plans an office-ai/command@1 envelope with stable command-schema fields", async () => {
    const client = await makeClient();
    const created = structured(
      await client.callTool({ name: "create_document", arguments: { format: "xlsx", name: "schema.xlsx" } })
    );
    const document = created.document as { documentId: string; sessionId: string; revision: number };
    const planned = structured(
      await client.callTool({
        name: "plan_command",
        arguments: {
          document_id: document.documentId,
          action_id: "xlsx.set-cell",
          arguments: { sheet: "Sheet1", ref: "A1", value: "schema" },
          target: { anchor: { kind: "range", sheet: "Sheet1", range: "A1:A1" } },
          policy: { mode: "auto_apply" },
          actor_id: "mcp-test",
        },
      })
    );
    expect(planned.ok).toBe(true);
    expect(planned.schema).toBe("office-ai/command-plan@1");
    expect(planned.commandId).toBeTruthy();
    expect(planned.command).toMatchObject({
      schema: "office-ai/command@1",
      format: "xlsx",
      operation: "xlsx:set-cell-value",
      arguments: { sheet: "Sheet1", ref: "A1", value: "schema" },
      target: {
        documentId: document.documentId,
        sessionId: document.sessionId,
        revision: document.revision,
        anchor: { kind: "range", sheet: "Sheet1", range: "A1:A1" },
      },
      source: { surface: "mcp", actorId: "mcp-test" },
      policy: { mode: "pending", requiresReview: true },
    });
    expect((planned.diagnostics as Array<{ level: string; code: string }>).map((d) => d.code)).toEqual(
      expect.arrayContaining(["catalog-review-required", "auto-apply-downgraded-to-pending"])
    );
    expect((planned.diagnostics as Array<{ level: string }>).every((d) => d.level)).toBe(true);
  });

  it("runs DOCX, XLSX, PPTX and PDF commands through plan -> preview -> apply -> export", async () => {
    const client = await makeClient();
    const tmp = mkdtempSync(join(tmpdir(), "officeai-mcp-command-"));
    const cases = [
      {
        format: "docx",
        operation: "docx:insert-text",
        args: { at: { paragraph: 0 }, text: "MCP DOCX " },
        anchor: { kind: "paragraph", index: 0 },
        out: "out.docx",
      },
      {
        format: "xlsx",
        operation: "xlsx:set-cell-value",
        args: { sheet: "Sheet1", ref: "A1", value: "MCP XLSX" },
        anchor: { kind: "range", sheet: "Sheet1", range: "A1:A1" },
        out: "out.xlsx",
      },
      {
        format: "pptx",
        operation: "pptx:add-slide",
        args: {},
        anchor: { kind: "slide_shape", slideIndex: 0 },
        out: "out.pptx",
      },
      {
        format: "pdf",
        operation: "pdf:rotate-pages",
        args: { pages: [1], delta: 90 },
        anchor: { kind: "page_region", page: 1, rect: { x: 0, y: 0, width: 10, height: 10 } },
        out: "out.pdf",
      },
    ] as const;

    for (const entry of cases) {
      const created = structured(
        await client.callTool({
          name: "create_document",
          arguments: { format: entry.format, name: `command.${entry.format}` },
        })
      );
      const document = created.document as { documentId: string };
      const planned = structured(
        await client.callTool({
          name: "plan_command",
          arguments: {
            document_id: document.documentId,
            ...(entry.actionId ? { action_id: entry.actionId } : { operation: entry.operation }),
            arguments: entry.args,
            target: { anchor: entry.anchor },
            policy: { mode: "auto_apply" },
          },
        })
      );
      expect(planned.ok, `${entry.format} plan failed`).toBe(true);
      expect((planned.command as { operation: string }).operation).toBe(entry.operation);

      const preview = structured(
        await client.callTool({ name: "preview_command", arguments: { command_id: planned.commandId } })
      );
      expect(preview.ok, `${entry.format} preview failed`).toBe(true);
      expect(((preview.diff as { changes?: unknown[] }).changes ?? []).length).toBeGreaterThan(0);
      expect(preview.semanticDiff).toMatchObject({
        schema: "office-ai/semantic-diff@1",
        format: entry.format,
        summary: { changeCount: expect.any(Number), text: expect.any(String) },
      });
      expect((preview.diagnostics as unknown[]).length).toBeGreaterThan(0);

      const applied = structured(
        await client.callTool({ name: "apply_command", arguments: { command_id: planned.commandId } })
      );
      expect(applied.ok, `${entry.format} apply failed`).toBe(true);
      expect(applied.stage).toBe("applied");
      expect((applied.mutation as { status: string }).status).toBe("approved");
      expect(((applied.diff as { changes?: unknown[] }).changes ?? []).length).toBeGreaterThan(0);
      expect(applied.semanticDiff).toMatchObject({
        schema: "office-ai/semantic-diff@1",
        format: entry.format,
      });

      const exported = structured(
        await client.callTool({
          name: "export_document",
          arguments: { document_id: document.documentId, out_path: join(tmp, entry.out) },
        })
      );
      expect((exported.exported as { bytes?: number }).bytes ?? 0).toBeGreaterThan(0);
      expect(
        (exported.diagnostics as Array<{ code: string }>).map((diagnostic) => diagnostic.code)
      ).toContain("export-command-basis");

      const activity = structured(
        await client.callTool({
          name: "list_activity",
          arguments: { document_id: document.documentId },
        })
      );
      expect(activity.schema).toBe("office-ai/activity-list@1");
      expect(
        (activity.activity as Array<{ stage: string; exportRef?: { commandIds: string[] } }>).some(
          (item) => item.stage === "exported" && (item.exportRef?.commandIds.length ?? 0) > 0
        )
      ).toBe(true);
    }
  });

  it("supports generic pending review with list, approve and reject", async () => {
    const client = await makeClient();
    const created = structured(
      await client.callTool({ name: "create_document", arguments: { format: "docx", name: "review.docx" } })
    );
    const document = created.document as { documentId: string };
    const first = structured(
      await client.callTool({
        name: "apply_command",
        arguments: {
          document_id: document.documentId,
          operation: "docx:insert-text",
          arguments: { at: { paragraph: 0 }, text: "approved " },
          target: { anchor: { kind: "paragraph", index: 0 } },
          policy: { mode: "pending" },
        },
      })
    );
    expect(first.stage).toBe("queued");
    const firstId = (first.mutation as { id: string }).id;

    const listed = structured(
      await client.callTool({ name: "list_pending_changes", arguments: { document_id: document.documentId } })
    );
    expect((listed.pending as Array<{ mutation: { id: string } }>).map((p) => p.mutation.id)).toContain(
      firstId
    );

    const approved = structured(
      await client.callTool({
        name: "approve_change",
        arguments: { document_id: document.documentId, mutation_id: firstId },
      })
    );
    expect(approved.approved).toBe(firstId);

    const second = structured(
      await client.callTool({
        name: "apply_command",
        arguments: {
          document_id: document.documentId,
          operation: "docx:insert-text",
          arguments: { at: { paragraph: 0 }, text: "rejected " },
          target: { anchor: { kind: "paragraph", index: 0 } },
          policy: { mode: "pending" },
        },
      })
    );
    const secondId = (second.mutation as { id: string }).id;
    const rejected = structured(
      await client.callTool({
        name: "reject_change",
        arguments: { document_id: document.documentId, mutation_id: secondId, reason: "test rejection" },
      })
    );
    expect(rejected.rejected).toBe(secondId);
    expect(rejected.reason).toBe("test rejection");

    const pendingAfter = structured(
      await client.callTool({ name: "list_pending_changes", arguments: { document_id: document.documentId } })
    );
    expect(pendingAfter.pending).toEqual([]);
  });

  it("downgrades review-required actions to pending and warns when exporting unreviewed changes", async () => {
    const client = await makeClient();
    const tmp = mkdtempSync(join(tmpdir(), "officeai-mcp-policy-"));
    const created = structured(
      await client.callTool({ name: "create_document", arguments: { format: "pdf", name: "review.pdf" } })
    );
    const document = created.document as { documentId: string };
    const planned = structured(
      await client.callTool({
        name: "plan_command",
        arguments: {
          document_id: document.documentId,
          action_id: "pdf.rotate-pages",
          arguments: { pages: [1], delta: 90 },
          target: { anchor: { kind: "page_region", page: 1, rect: { x: 0, y: 0, width: 10, height: 10 } } },
          policy: { mode: "auto_apply", requires_review: false },
        },
      })
    );

    expect(planned.ok).toBe(true);
    expect(planned.command).toMatchObject({
      operation: "pdf:rotate-pages",
      policy: { mode: "pending", requiresReview: true },
    });
    expect((planned.diagnostics as Array<{ code: string }>).map((d) => d.code)).toEqual(
      expect.arrayContaining([
        "catalog-review-required",
        "review-opt-out-ignored",
        "auto-apply-downgraded-to-pending",
      ])
    );

    const applied = structured(
      await client.callTool({ name: "apply_command", arguments: { command_id: planned.commandId } })
    );
    expect(applied.stage).toBe("queued");
    expect((applied.mutation as { status: string }).status).toBe("pending");

    const exported = structured(
      await client.callTool({
        name: "export_document",
        arguments: { document_id: document.documentId, out_path: join(tmp, "review.pdf") },
      })
    );
    expect((exported.exported as { bytes?: number }).bytes ?? 0).toBeGreaterThan(0);
    expect((exported.diagnostics as Array<{ code: string }>).map((d) => d.code)).toContain(
      "unreviewed-pending-export"
    );

    const pendingAfterExport = structured(
      await client.callTool({ name: "list_pending_changes", arguments: { document_id: document.documentId } })
    );
    expect(pendingAfterExport.pending as unknown[]).toHaveLength(1);
  });

  it("does not silently auto-apply destructive direct operations", async () => {
    const client = await makeClient();
    const imported = structured(
      await client.callTool({
        name: "import_document",
        arguments: {
          path: fixturePath(requiredMatrixFixture("pdf", { id: "pdf.synthetic.simple-text-3page" })),
        },
      })
    );
    const document = imported.document as { documentId: string };

    const applied = structured(
      await client.callTool({
        name: "apply_command",
        arguments: {
          document_id: document.documentId,
          operation: "pdf:delete-pages",
          arguments: { pages: [1] },
          target: { anchor: { kind: "page_region", page: 1, rect: { x: 0, y: 0, width: 10, height: 10 } } },
          policy: { mode: "auto_apply", requires_review: false },
        },
      })
    );

    expect(applied.ok).toBe(true);
    expect(applied.stage).toBe("queued");
    expect((applied.mutation as { status: string }).status).toBe("pending");
    expect((applied.command as { policy: { mode: string; requiresReview: boolean } }).policy).toEqual({
      mode: "pending",
      requiresReview: true,
    });
    expect((applied.diagnostics as Array<{ code: string }>).map((d) => d.code)).toEqual(
      expect.arrayContaining([
        "destructive-command-review-required",
        "review-opt-out-ignored",
        "auto-apply-downgraded-to-pending",
      ])
    );
  });

  it("supports canonical undo for approved commands", async () => {
    const client = await makeClient();
    const created = structured(
      await client.callTool({ name: "create_document", arguments: { format: "xlsx", name: "undo.xlsx" } })
    );
    const document = created.document as { documentId: string };
    const applied = structured(
      await client.callTool({
        name: "apply_command",
        arguments: {
          document_id: document.documentId,
          operation: "xlsx:set-cell-value",
          arguments: { sheet: "Sheet1", ref: "A1", value: "undo me" },
          target: { anchor: { kind: "range", sheet: "Sheet1", range: "A1:A1" } },
          policy: { mode: "auto_apply", requires_review: false },
        },
      })
    );
    expect(applied.stage).toBe("applied");

    const undone = structured(
      await client.callTool({ name: "undo_command", arguments: { document_id: document.documentId } })
    );
    expect(undone.didUndo).toBe(true);
    expect((undone.diagnostics as Array<{ code: string }>).map((d) => d.code)).toContain("undo-applied");

    const empty = structured(
      await client.callTool({ name: "undo_command", arguments: { document_id: document.documentId } })
    );
    expect(empty.didUndo).toBe(false);
    expect((empty.diagnostics as Array<{ code: string }>).map((d) => d.code)).toContain("undo-empty");
  });

  it("reloads persisted canonical documents after MCP server restart", async () => {
    const firstClient = await makeClient();
    const created = structured(
      await firstClient.callTool({
        name: "create_document",
        arguments: { format: "xlsx", name: "restart.xlsx" },
      })
    );
    const document = created.document as { documentId: string };
    const applied = structured(
      await firstClient.callTool({
        name: "apply_command",
        arguments: {
          document_id: document.documentId,
          operation: "xlsx:set-cell-value",
          arguments: { sheet: "Sheet1", ref: "A1", value: "persisted after restart" },
          target: { anchor: { kind: "range", sheet: "Sheet1", range: "A1:A1" } },
          policy: { mode: "auto_apply" },
        },
      })
    );
    expect(applied.stage).toBe("applied");

    __resetMcpSessionsForTests();

    const restartedClient = await makeClient();
    const loaded = structured(
      await restartedClient.callTool({
        name: "get_document",
        arguments: { document_id: document.documentId },
      })
    );
    expect((loaded.document as { documentId: string }).documentId).toBe(document.documentId);
    const projected = structured(
      await restartedClient.callTool({
        name: "get_document_projection",
        arguments: {
          document_id: document.documentId,
          projection: "json",
          sheet: "Sheet1",
          range: "A1:A1",
        },
      })
    );
    expect(JSON.stringify(projected)).toContain("persisted after restart");
    const listed = structured(await restartedClient.callTool({ name: "list_documents", arguments: {} }));
    expect((listed.documents as Array<{ documentId: string }>).map((doc) => doc.documentId)).toContain(
      document.documentId
    );
  });

  it("fails loud for unsupported operations, invalid anchors and stale revisions", async () => {
    const client = await makeClient();
    const created = structured(
      await client.callTool({ name: "create_document", arguments: { format: "docx", name: "negative.docx" } })
    );
    const document = created.document as { documentId: string; revision: number };

    const unsupported = structured(
      await client.callTool({
        name: "plan_command",
        arguments: {
          document_id: document.documentId,
          operation: "docx:nope",
          arguments: {},
        },
      })
    );
    expect(unsupported.ok).toBe(false);
    expect((unsupported.diagnostics as Array<{ code: string }>).map((d) => d.code)).toContain(
      "unsupported-operation"
    );

    const invalidAnchor = structured(
      await client.callTool({
        name: "plan_command",
        arguments: {
          document_id: document.documentId,
          operation: "docx:insert-text",
          arguments: { at: { paragraph: 0 }, text: "x" },
          target: { anchor: { kind: "range", sheet: "Sheet1", range: "A1" } },
        },
      })
    );
    expect(invalidAnchor.ok).toBe(false);
    expect((invalidAnchor.diagnostics as Array<{ code: string }>).map((d) => d.code)).toContain(
      "invalid-anchor-kind"
    );

    const stale = structured(
      await client.callTool({
        name: "apply_command",
        arguments: {
          document_id: document.documentId,
          operation: "docx:insert-text",
          arguments: { at: { paragraph: 0 }, text: "x" },
          target: { revision: document.revision + 99, anchor: { kind: "paragraph", index: 0 } },
          policy: { mode: "auto_apply" },
        },
      })
    );
    expect(stale.ok).toBe(false);
    expect(stale.stage).toBe("failed");
    expect((stale.diagnostics as Array<{ code: string }>).map((d) => d.code)).toContain("stale-revision");
    const after = structured(
      await client.callTool({ name: "get_document", arguments: { document_id: document.documentId } })
    );
    expect((after.document as { revision: number }).revision).toBe(document.revision);
  });

  it("canonical session flow imports, projects and exports every core format", async () => {
    const client = await makeClient();
    const session = structured(
      await client.callTool({ name: "create_session", arguments: { title: "matrix smoke" } })
    );
    const sessionId = session.sessionId as string;
    const tmp = mkdtempSync(join(tmpdir(), "officeai-mcp-docs-"));
    const extByFormat: Record<FixtureFormat, string> = {
      docx: "docx",
      xlsx: "xlsx",
      pptx: "pptx",
      pdf: "pdf",
    };

    for (const format of ["docx", "xlsx", "pptx", "pdf"] as const) {
      const imported = structured(
        await client.callTool({
          name: "import_document",
          arguments: { session_id: sessionId, path: matrixPath(format) },
        })
      );
      const document = imported.document as { documentId: string; sessionId: string; format: string };
      expect(document.sessionId).toBe(sessionId);
      expect(document.format).toBe(format);

      const summary = structured(
        await client.callTool({
          name: "get_document_projection",
          arguments: { document_id: document.documentId, projection: "summary" },
        })
      );
      expect(summary.documentId).toBe(document.documentId);
      expect(summary.schema).toBe("office-ai/document-projection@1");
      expect((summary.document as { documentId: string }).documentId).toBe(document.documentId);
      expect(summary.summary).toBeTruthy();

      const exported = structured(
        await client.callTool({
          name: "export_document",
          arguments: { document_id: document.documentId, out_path: join(tmp, `out.${extByFormat[format]}`) },
        })
      );
      expect((exported.exported as { bytes?: number }).bytes ?? 0).toBeGreaterThan(0);

      const legacyInspectTool =
        format === "docx"
          ? "docx_inspect"
          : format === "xlsx"
            ? "xlsx_inspect"
            : format === "pptx"
              ? "pptx_inspect"
              : "pdf_metadata";
      const legacy = structured(
        await client.callTool({ name: legacyInspectTool, arguments: { handle: document.documentId } })
      );
      expect(legacy).toBeTruthy();
    }

    const listed = structured(
      await client.callTool({ name: "list_documents", arguments: { session_id: sessionId } })
    );
    expect((listed.documents as unknown[]).length).toBe(4);
  });

  it("CLI session import and MCP import_document agree on projections for the same file", async () => {
    if (!activeOfficeAiDataDir) throw new Error("expected OFFICEAI_DATA_DIR for test");
    const path = await makeFixture();
    const client = await makeClient();

    const mcpImported = structured(await client.callTool({ name: "import_document", arguments: { path } }));
    expect(mcpImported.schema).toBe("office-ai/import-document@1");
    const mcpDocument = mcpImported.document as { documentId: string; format: string };
    expect(mcpDocument.format).toBe("docx");
    const mcpProjection = structured(
      await client.callTool({
        name: "get_document_projection",
        arguments: { document_id: mcpDocument.documentId, projection: "markdown" },
      })
    );

    const importIO = makeIO();
    const cliImportCode = await runCli(
      ["sessions", "import", "--json", "--data-dir", activeOfficeAiDataDir, "--file", path],
      importIO.io
    );
    expect(cliImportCode, importIO.stderr.text()).toBe(0);
    const cliImported = JSON.parse(importIO.stdout.text()) as {
      schema: string;
      document: { documentId: string; format: string };
    };
    expect(cliImported.schema).toBe("office-ai/import-document@1");
    expect(cliImported.document.format).toBe("docx");

    const projectionIO = makeIO();
    const cliProjectionCode = await runCli(
      [
        "sessions",
        "projection",
        "--json",
        "--data-dir",
        activeOfficeAiDataDir,
        "--document-id",
        cliImported.document.documentId,
        "--projection",
        "markdown",
      ],
      projectionIO.io
    );
    expect(cliProjectionCode, projectionIO.stderr.text()).toBe(0);
    const cliProjection = JSON.parse(projectionIO.stdout.text()) as { schema: string; content: string };
    expect(cliProjection.schema).toBe("office-ai/document-projection@1");
    expect(cliProjection.content).toBe(mcpProjection.content);
  });

  it("canonical create_document exports blank documents for every core format", async () => {
    const client = await makeClient();
    const tmp = mkdtempSync(join(tmpdir(), "officeai-mcp-create-"));
    const extByFormat: Record<FixtureFormat, string> = {
      docx: "docx",
      xlsx: "xlsx",
      pptx: "pptx",
      pdf: "pdf",
    };

    for (const format of ["docx", "xlsx", "pptx", "pdf"] as const) {
      const created = structured(
        await client.callTool({ name: "create_document", arguments: { format, name: `blank.${format}` } })
      );
      const document = created.document as { documentId: string; format: string };
      expect(document.format).toBe(format);

      const projected = structured(
        await client.callTool({
          name: "get_document_projection",
          arguments: { document_id: document.documentId, projection: "summary" },
        })
      );
      expect(projected.summary).toBeTruthy();

      const exported = structured(
        await client.callTool({
          name: "export_document",
          arguments: {
            document_id: document.documentId,
            out_path: join(tmp, `blank.${extByFormat[format]}`),
          },
        })
      );
      expect((exported.exported as { bytes?: number }).bytes ?? 0).toBeGreaterThan(0);
    }
  });

  it("canonical import_document reports invalid extensions as tool errors", async () => {
    const client = await makeClient();
    const dir = mkdtempSync(join(tmpdir(), "officeai-mcp-invalid-"));
    const path = join(dir, "not-a-document.txt");
    writeFileSync(path, "hello");
    const result = (await client.callTool({ name: "import_document", arguments: { path } })) as {
      isError?: boolean;
      content?: Array<{ text: string }>;
    };
    expect(result.isError).toBe(true);
    expect(result.content?.[0]?.text ?? "").toContain("Cannot infer document format");
  });

  it("loads matrix-selected fixtures for every core format", async () => {
    const client = await makeClient();

    const docx = structured(
      await client.callTool({ name: "docx_load", arguments: { path: matrixPath("docx") } })
    );
    expect(typeof docx.handle).toBe("string");
    expect((docx.summary as { paragraphs?: number }).paragraphs ?? 0).toBeGreaterThan(0);

    const xlsx = structured(
      await client.callTool({ name: "xlsx_load", arguments: { path: matrixPath("xlsx") } })
    );
    expect(typeof xlsx.handle).toBe("string");
    expect(((xlsx.summary as { sheets?: unknown[] }).sheets ?? []).length).toBeGreaterThan(0);

    const pptx = structured(
      await client.callTool({ name: "pptx_load", arguments: { path: matrixPath("pptx") } })
    );
    expect(typeof pptx.handle).toBe("string");
    expect((pptx.summary as { slides?: number }).slides ?? 0).toBeGreaterThan(0);

    const pdf = structured(
      await client.callTool({ name: "pdf_load", arguments: { path: matrixPath("pdf") } })
    );
    expect(typeof pdf.handle).toBe("string");
    expect((pdf.summary as { pageCount?: number }).pageCount ?? 0).toBeGreaterThan(0);
  });

  it("auto-binds catalogue actions that declare args + buildPayload as MCP tools", async () => {
    const client = await makeClient();
    const list = await client.listTools();
    const names = new Set(list.tools.map((t) => t.name));
    // Sanity-check a few catalogue-driven tool names that the
    // actions-to-mcp adapter is expected to register. New catalogue
    // entries with `buildPayload` will appear here automatically; the
    // assertion intentionally stops at "must contain" rather than
    // pinning the exact shape so the catalogue can grow.
    expect(names.has("docx_delete_row")).toBe(true);
    expect(names.has("docx_delete_column")).toBe(true);
    expect(names.has("docx_delete_table")).toBe(true);
    expect(names.has("docx_accept_all_changes")).toBe(true);
    expect(names.has("docx_reject_all_changes")).toBe(true);
    expect(names.has("docx_insert_page_break")).toBe(true);
  });

  it("auto-bound docx_insert_page_break dispatches the bus command and bumps the revision", async () => {
    const client = await makeClient();
    const path = await makeFixture();
    const loaded = structured(await client.callTool({ name: "docx_load", arguments: { path } }));
    const handle = loaded.handle as string;
    const inspectResult = structured(await client.callTool({ name: "docx_inspect", arguments: { handle } }));
    const beforeRev = (inspectResult.revision as number | undefined) ?? 0;
    const text = structured(
      await client.callTool({ name: "docx_get_text", arguments: { handle, format: "json" } })
    );
    const paragraphs = text.paragraphs as Array<{ id: string }>;
    const targetId = paragraphs[1].id;
    const result = structured(
      await client.callTool({
        name: "docx_insert_page_break",
        arguments: { handle, paragraphId: targetId, offset: 0 },
      })
    );
    expect(result.commandType).toBe("docx:insert-page-break");
    expect(result.status).toBe("approved");
    const after = structured(await client.callTool({ name: "docx_inspect", arguments: { handle } }));
    const afterRev = (after.revision as number | undefined) ?? 0;
    expect(afterRev).toBeGreaterThan(beforeRev);
  });

  it("docx_load returns a handle and an inspection summary", async () => {
    const client = await makeClient();
    const path = await makeFixture();
    const out = structured(await client.callTool({ name: "docx_load", arguments: { path } }));
    expect(typeof out.handle).toBe("string");
    const summary = out.summary as { paragraphs: number; tables: number; comments: number };
    expect(summary.paragraphs).toBe(3);
    expect(summary.comments).toBe(0);
    expect(summary.tables).toBe(0);
  });

  it("docx_inspect returns counts and parts list", async () => {
    const client = await makeClient();
    const path = await makeFixture();
    const loaded = structured(await client.callTool({ name: "docx_load", arguments: { path } }));
    const handle = loaded.handle as string;
    const insp = structured(await client.callTool({ name: "docx_inspect", arguments: { handle } }));
    expect(insp.format).toBe("docx");
    expect(insp.paragraphs).toBe(3);
    expect(Array.isArray(insp.parts)).toBe(true);
    expect((insp.parts as string[]).some((p) => p.includes("document.xml"))).toBe(true);
  });

  it("docx_get_text supports markdown, json, and text formats", async () => {
    const client = await makeClient();
    const path = await makeFixture();
    const loaded = structured(await client.callTool({ name: "docx_load", arguments: { path } }));
    const handle = loaded.handle as string;

    const md = structured(
      await client.callTool({ name: "docx_get_text", arguments: { handle, format: "markdown" } })
    );
    expect(typeof md.content).toBe("string");
    expect(md.content as string).toContain("# Hello");

    const json = structured(
      await client.callTool({ name: "docx_get_text", arguments: { handle, format: "json" } })
    );
    expect(json.format).toBe("docx");
    expect(Array.isArray(json.paragraphs)).toBe(true);
    expect((json.paragraphs as Array<{ text: string }>).length).toBe(3);

    const text = structured(
      await client.callTool({ name: "docx_get_text", arguments: { handle, format: "text" } })
    );
    expect(text.content as string).toContain("first paragraph body");
    expect(text.content as string).not.toContain("#");
  });

  it("docx_get_pages returns at least one doc-start page for a single-page doc", async () => {
    const client = await makeClient();
    const path = await makeFixture();
    const loaded = structured(await client.callTool({ name: "docx_load", arguments: { path } }));
    const handle = loaded.handle as string;
    const out = structured(await client.callTool({ name: "docx_get_pages", arguments: { handle } }));
    const pages = out.pages as Array<{
      pageNumber: number;
      trigger: string;
      preview: string;
    }>;
    expect(pages.length).toBeGreaterThanOrEqual(1);
    expect(pages[0].pageNumber).toBe(1);
    expect(pages[0].trigger).toBe("doc-start");
    expect(out.total).toBe(pages.length);
  });

  it("docx_get_page_text returns markdown for an in-range page and errors for out-of-range", async () => {
    const client = await makeClient();
    const path = await makeFixture();
    const loaded = structured(await client.callTool({ name: "docx_load", arguments: { path } }));
    const handle = loaded.handle as string;
    const md = structured(
      await client.callTool({ name: "docx_get_page_text", arguments: { handle, page: 1 } })
    );
    expect(md.format).toBe("markdown");
    expect(typeof md.content).toBe("string");
    expect(md.content as string).toContain("# Hello");

    const errResult = (await client.callTool({
      name: "docx_get_page_text",
      arguments: { handle, page: 999 },
    })) as { isError?: boolean; content?: Array<{ text: string }> };
    expect(errResult.isError).toBe(true);
    const text = errResult.content?.[0]?.text ?? "";
    expect(text).toContain("out-of-range");
  });

  it("docx_get_text with with_page_sections injects page anchors", async () => {
    const client = await makeClient();
    const path = await makeFixture();
    const loaded = structured(await client.callTool({ name: "docx_load", arguments: { path } }));
    const handle = loaded.handle as string;
    const md = structured(
      await client.callTool({
        name: "docx_get_text",
        arguments: { handle, format: "markdown", with_page_sections: true },
      })
    );
    expect(md.content as string).toContain("<!-- page 1 -->");
    expect(md.content as string).toContain("## Page 1");
  });

  it("docx_search returns matches", async () => {
    const client = await makeClient();
    const path = await makeFixture();
    const loaded = structured(await client.callTool({ name: "docx_load", arguments: { path } }));
    const handle = loaded.handle as string;
    const r = structured(
      await client.callTool({
        name: "docx_search",
        arguments: { handle, query: "paragraph" },
      })
    );
    const matches = r.matches as Array<{ paragraphIndex: number; match: string }>;
    expect(matches.length).toBeGreaterThanOrEqual(2);
    expect(matches[0].match.toLowerCase()).toBe("paragraph");
  });

  it("docx_apply_command can insert text and docx_save persists it", async () => {
    const client = await makeClient();
    const path = await makeFixture();
    const loaded = structured(await client.callTool({ name: "docx_load", arguments: { path } }));
    const handle = loaded.handle as string;

    const apply = structured(
      await client.callTool({
        name: "docx_apply_command",
        arguments: {
          handle,
          type: "docx:insert-text",
          payload: { at: { paragraph: 1, run: 0, offset: 0 }, text: "MCP " },
        },
      })
    );
    expect((apply.mutation as { status: string }).status).toBe("approved");

    const dir = mkdtempSync(join(tmpdir(), "mcp-out-"));
    const out = join(dir, "out.docx");
    const saved = structured(
      await client.callTool({ name: "docx_save", arguments: { handle, out_path: out } })
    );
    expect(saved.wrote).toBe(out);
    const reloaded = readFileSync(out);
    expect(reloaded.byteLength).toBeGreaterThan(0);
  });

  it("end-to-end comment lifecycle via docx_apply_command", async () => {
    const client = await makeClient();
    const path = await makeFixture();
    const loaded = structured(await client.callTool({ name: "docx_load", arguments: { path } }));
    const handle = loaded.handle as string;

    await client.callTool({
      name: "docx_apply_command",
      arguments: {
        handle,
        type: "docx:add-comment",
        payload: {
          range: {
            start: { paragraph: 1, run: 0, offset: 0 },
            end: { paragraph: 1, run: 0, offset: 5 },
          },
          text: "rephrase?",
          author: "AI",
          initials: "AI",
        },
      },
    });

    const inspectAfterAdd = structured(
      await client.callTool({ name: "docx_inspect", arguments: { handle } })
    );
    expect(inspectAfterAdd.comments).toBe(1);

    const projection = structured(
      await client.callTool({ name: "docx_get_text", arguments: { handle, format: "json" } })
    );
    const comments = projection.comments as Array<{ id: string; resolved?: boolean; parentId?: string }>;
    expect(comments).toHaveLength(1);
    const commentId = comments[0].id;

    await client.callTool({
      name: "docx_apply_command",
      arguments: {
        handle,
        type: "docx:resolve-comment",
        payload: { commentId },
      },
    });
    const afterResolve = structured(
      await client.callTool({ name: "docx_get_text", arguments: { handle, format: "json" } })
    );
    expect((afterResolve.comments as Array<{ resolved?: boolean }>)[0].resolved).toBe(true);

    await client.callTool({
      name: "docx_apply_command",
      arguments: {
        handle,
        type: "docx:reply-comment",
        payload: { parentId: commentId, text: "ack", author: "Bob" },
      },
    });
    const afterReply = structured(
      await client.callTool({ name: "docx_get_text", arguments: { handle, format: "json" } })
    );
    const replyComments = afterReply.comments as Array<{ id: string; parentId?: string }>;
    expect(replyComments).toHaveLength(2);
    expect(replyComments.find((c) => c.id !== commentId)?.parentId).toBe(commentId);

    await client.callTool({
      name: "docx_apply_command",
      arguments: {
        handle,
        type: "docx:delete-comment",
        payload: { commentId },
      },
    });
    const afterDelete = structured(await client.callTool({ name: "docx_inspect", arguments: { handle } }));
    expect(afterDelete.comments).toBe(0);
  });

  it("docx_diff reports changes between two handles", async () => {
    const client = await makeClient();
    const path = await makeFixture();
    const a = structured(await client.callTool({ name: "docx_load", arguments: { path } }));
    const b = structured(await client.callTool({ name: "docx_load", arguments: { path } }));

    await client.callTool({
      name: "docx_apply_command",
      arguments: {
        handle: b.handle,
        type: "docx:insert-text",
        payload: { at: { paragraph: 1, run: 0, offset: 0 }, text: "DIFF " },
      },
    });

    const diff = structured(
      await client.callTool({
        name: "docx_diff",
        arguments: { before: a.handle, after: b.handle },
      })
    );
    const paragraphs = diff.paragraphs as { added: number; removed: number; modified: number };
    expect(paragraphs.modified).toBe(1);
  });

  it("pending review flow: apply auto_approve=false, list, then approve & reject", async () => {
    const client = await makeClient();
    const path = await makeFixture();
    const loaded = structured(await client.callTool({ name: "docx_load", arguments: { path } }));
    const handle = loaded.handle as string;

    const apply = structured(
      await client.callTool({
        name: "docx_apply_command",
        arguments: {
          handle,
          type: "docx:insert-text",
          payload: { at: { paragraph: 1, run: 0, offset: 0 }, text: "PENDING " },
          auto_approve: false,
        },
      })
    );
    expect((apply.mutation as { status: string }).status).toBe("pending");
    const mutationId = (apply.mutation as { id: string }).id;

    const list1 = structured(await client.callTool({ name: "docx_list_pending", arguments: { handle } }));
    expect(list1.pending as Array<{ id: string }>).toHaveLength(1);
    expect((list1.pending as Array<{ id: string }>)[0].id).toBe(mutationId);

    const approve = structured(
      await client.callTool({
        name: "docx_approve",
        arguments: { handle, mutation_id: mutationId },
      })
    );
    expect(approve.approved).toBe(mutationId);

    const list2 = structured(await client.callTool({ name: "docx_list_pending", arguments: { handle } }));
    expect(list2.pending as unknown[]).toHaveLength(0);

    const apply2 = structured(
      await client.callTool({
        name: "docx_apply_command",
        arguments: {
          handle,
          type: "docx:insert-text",
          payload: { at: { paragraph: 1, run: 0, offset: 0 }, text: "REJECTME " },
          auto_approve: false,
        },
      })
    );
    const id2 = (apply2.mutation as { id: string }).id;
    const reject = structured(
      await client.callTool({
        name: "docx_reject",
        arguments: { handle, mutation_id: id2, reason: "no thanks" },
      })
    );
    expect(reject.rejected).toBe(id2);
    expect(reject.reason).toBe("no thanks");

    const list3 = structured(await client.callTool({ name: "docx_list_pending", arguments: { handle } }));
    expect(list3.pending as unknown[]).toHaveLength(0);
  });

  it("returns an error for unknown handles", async () => {
    const client = await makeClient();
    const r = await client.callTool({
      name: "docx_inspect",
      arguments: { handle: "nope" },
    });
    expect(r.isError).toBe(true);
  });
});

describe("OfficeAI MCP server — PPTX tools", () => {
  beforeEach(() => __resetMcpSessionsForTests());
  afterEach(() => __resetMcpSessionsForTests());

  const PPTX_FIXTURE = resolve(
    fileURLToPath(new URL("../../../fixtures/pptx/synthetic/03-title-and-content.pptx", import.meta.url))
  );
  const PPTX_TABLE_FIXTURE = resolve(
    fileURLToPath(new URL("../../../fixtures/pptx/synthetic/06-with-table.pptx", import.meta.url))
  );
  const PPTX_CHART_FIXTURE = resolve(
    fileURLToPath(new URL("../../../fixtures/pptx/synthetic/09-with-chart.pptx", import.meta.url))
  );
  const PPTX_ANIM_FIXTURE = resolve(
    fileURLToPath(new URL("../../../fixtures/pptx/synthetic/10-with-anim.pptx", import.meta.url))
  );

  it("pptx_load returns a handle and an inspection summary", async () => {
    const client = await makeClient();
    const out = structured(await client.callTool({ name: "pptx_load", arguments: { path: PPTX_FIXTURE } }));
    expect(typeof out.handle).toBe("string");
    const summary = out.summary as {
      slides: number;
      shapeCounts: { text: number };
      slideSize: { cxEmu: number; cyEmu: number };
    };
    expect(summary.slides).toBe(1);
    expect(summary.shapeCounts.text).toBeGreaterThanOrEqual(1);
    expect(summary.slideSize.cxEmu).toBeGreaterThan(0);
  });

  it("pptx_get_text supports markdown / json / text", async () => {
    const client = await makeClient();
    const loaded = structured(
      await client.callTool({ name: "pptx_load", arguments: { path: PPTX_FIXTURE } })
    );
    const handle = loaded.handle as string;

    const md = structured(
      await client.callTool({ name: "pptx_get_text", arguments: { handle, format: "markdown" } })
    );
    expect((md.content as string).startsWith("# Presentation")).toBe(true);

    const json = structured(
      await client.callTool({ name: "pptx_get_text", arguments: { handle, format: "json" } })
    );
    expect(json.format).toBe("pptx");
    expect(Array.isArray(json.slides)).toBe(true);

    const text = structured(
      await client.callTool({ name: "pptx_get_text", arguments: { handle, format: "text" } })
    );
    expect(typeof text.content).toBe("string");
  });

  it("pptx_apply_command + pptx_save round-trips an edit through disk", async () => {
    const client = await makeClient();
    const loaded = structured(
      await client.callTool({ name: "pptx_load", arguments: { path: PPTX_FIXTURE } })
    );
    const handle = loaded.handle as string;

    const apply = structured(
      await client.callTool({
        name: "pptx_apply_command",
        arguments: { handle, type: "pptx:add-slide", payload: {} },
      })
    );
    expect((apply.mutation as { status: string }).status).toBe("approved");

    const dir = mkdtempSync(join(tmpdir(), "mcp-pptx-out-"));
    const out = join(dir, "out.pptx");
    const saved = structured(
      await client.callTool({ name: "pptx_save", arguments: { handle, out_path: out } })
    );
    expect(saved.wrote).toBe(out);
    expect(readFileSync(out).byteLength).toBeGreaterThan(0);
  });

  it("pptx_diff reports slide-level changes between two handles", async () => {
    const client = await makeClient();
    const a = structured(await client.callTool({ name: "pptx_load", arguments: { path: PPTX_FIXTURE } }));
    const b = structured(await client.callTool({ name: "pptx_load", arguments: { path: PPTX_FIXTURE } }));

    await client.callTool({
      name: "pptx_apply_command",
      arguments: { handle: b.handle, type: "pptx:add-slide", payload: {} },
    });

    const diff = structured(
      await client.callTool({ name: "pptx_diff", arguments: { before: a.handle, after: b.handle } })
    );
    expect((diff.slides as { added: number }).added).toBe(1);
  });

  it("pptx_search returns text matches", async () => {
    const client = await makeClient();
    const loaded = structured(
      await client.callTool({ name: "pptx_load", arguments: { path: PPTX_FIXTURE } })
    );
    const r = structured(
      await client.callTool({
        name: "pptx_search",
        arguments: { handle: loaded.handle, query: "Title" },
      })
    );
    expect(Array.isArray(r.matches)).toBe(true);
  });

  it("pptx_apply_command dispatches typed table commands", async () => {
    const client = await makeClient();
    const loaded = structured(
      await client.callTool({ name: "pptx_load", arguments: { path: PPTX_TABLE_FIXTURE } })
    );
    const handle = loaded.handle as string;

    const json = structured(
      await client.callTool({ name: "pptx_get_text", arguments: { handle, format: "json" } })
    ) as unknown as {
      slides: Array<{ index: number; shapes: Array<{ id: string; kind: string }> }>;
    };
    let slideIndex = -1;
    let tableId = "";
    for (const slide of json.slides) {
      const t = slide.shapes.find((s) => s.kind === "table");
      if (t) {
        slideIndex = slide.index;
        tableId = t.id;
        break;
      }
    }
    expect(tableId).not.toBe("");

    const apply = structured(
      await client.callTool({
        name: "pptx_apply_command",
        arguments: {
          handle,
          type: "pptx:table-set-cell-text",
          payload: { slideIndex, shapeId: tableId, row: 0, column: 0, text: "MCP Cell" },
        },
      })
    );
    expect((apply.mutation as { status: string }).status).toBe("approved");

    const after = structured(
      await client.callTool({ name: "pptx_get_text", arguments: { handle, format: "text" } })
    );
    expect((after.content as string).includes("MCP Cell")).toBe(true);
  });

  it("pptx_apply_command dispatches typed chart commands and json projection surfaces them", async () => {
    const client = await makeClient();
    const loaded = structured(
      await client.callTool({ name: "pptx_load", arguments: { path: PPTX_CHART_FIXTURE } })
    );
    const handle = loaded.handle as string;

    const summary = loaded.summary as { shapeCounts: { chart: number } };
    expect(summary.shapeCounts.chart).toBeGreaterThanOrEqual(1);

    const json = structured(
      await client.callTool({ name: "pptx_get_text", arguments: { handle, format: "json" } })
    ) as unknown as {
      slides: Array<{
        index: number;
        shapes: Array<{ id: string; kind: string; chart?: { chartType: string } }>;
      }>;
    };
    let slideIndex = -1;
    let chartId = "";
    for (const slide of json.slides) {
      const c = slide.shapes.find((s) => s.kind === "chart");
      if (c) {
        slideIndex = slide.index;
        chartId = c.id;
        break;
      }
    }
    expect(chartId).not.toBe("");

    const apply = structured(
      await client.callTool({
        name: "pptx_apply_command",
        arguments: {
          handle,
          type: "pptx:set-chart-title",
          payload: { slideIndex, shapeId: chartId, title: "MCP Chart Title" },
        },
      })
    );
    expect((apply.mutation as { status: string }).status).toBe("approved");

    const apply2 = structured(
      await client.callTool({
        name: "pptx_apply_command",
        arguments: {
          handle,
          type: "pptx:set-chart-type",
          payload: { slideIndex, shapeId: chartId, chartType: "pie" },
        },
      })
    );
    expect((apply2.mutation as { status: string }).status).toBe("approved");

    const md = structured(
      await client.callTool({
        name: "pptx_get_text",
        arguments: { handle, format: "markdown", slide: slideIndex },
      })
    );
    expect((md.content as string).includes("MCP Chart Title")).toBe(true);

    const after = structured(
      await client.callTool({ name: "pptx_get_text", arguments: { handle, format: "json" } })
    ) as unknown as {
      slides: Array<{ shapes: Array<{ kind: string; chart?: { chartType: string; title?: string } }> }>;
    };
    const ch = after.slides.flatMap((s) => s.shapes).find((sh) => sh.kind === "chart");
    expect(ch?.chart?.chartType).toBe("pie");
    expect(ch?.chart?.title).toBe("MCP Chart Title");
  });

  it("pptx_apply_command dispatches typed animation commands and json projection surfaces them", async () => {
    const client = await makeClient();
    const loaded = structured(
      await client.callTool({ name: "pptx_load", arguments: { path: PPTX_ANIM_FIXTURE } })
    );
    const handle = loaded.handle as string;

    const summary = loaded.summary as { animations: number; transitions: number };
    expect(summary.animations).toBeGreaterThan(0);
    expect(summary.transitions).toBeGreaterThan(0);

    const json = structured(
      await client.callTool({ name: "pptx_get_text", arguments: { handle, format: "json" } })
    ) as unknown as {
      slides: Array<{
        index: number;
        transition?: { kind: string };
        animations?: Array<{ id: string; effect: string; targetCNvPrId: number }>;
      }>;
    };
    const slide0 = json.slides[0]!;
    expect(slide0.transition?.kind).toBe("fade");
    expect(slide0.animations?.length).toBe(2);

    const apply = structured(
      await client.callTool({
        name: "pptx_apply_command",
        arguments: {
          handle,
          type: "pptx:set-slide-transition",
          payload: { slideIndex: 0, kind: "push", speed: "fast" },
        },
      })
    );
    expect((apply.mutation as { status: string }).status).toBe("approved");

    const dropId = slide0.animations![0]!.id;
    const apply2 = structured(
      await client.callTool({
        name: "pptx_apply_command",
        arguments: {
          handle,
          type: "pptx:remove-shape-animation",
          payload: { slideIndex: 0, animationId: dropId },
        },
      })
    );
    expect((apply2.mutation as { status: string }).status).toBe("approved");

    const after = structured(
      await client.callTool({ name: "pptx_get_text", arguments: { handle, format: "json" } })
    ) as unknown as {
      slides: Array<{
        transition?: { kind: string; speed?: string };
        animations?: Array<{ effect: string }>;
      }>;
    };
    expect(after.slides[0]!.transition?.kind).toBe("push");
    expect(after.slides[0]!.transition?.speed).toBe("fast");
    expect(after.slides[0]!.animations?.length).toBe(1);

    const md = structured(
      await client.callTool({
        name: "pptx_get_text",
        arguments: { handle, format: "markdown", slide: 0 },
      })
    );
    expect((md.content as string).includes("transition")).toBe(true);
    expect((md.content as string).includes("push")).toBe(true);
  });

  it("returns an error for unknown pptx handles", async () => {
    const client = await makeClient();
    const r = await client.callTool({ name: "pptx_inspect", arguments: { handle: "nope" } });
    expect(r.isError).toBe(true);
  });
});

describe("OfficeAI MCP server — xlsx tools", () => {
  beforeEach(() => __resetMcpSessionsForTests());
  afterEach(() => __resetMcpSessionsForTests());

  async function loadFixture(client: Client, name: string): Promise<{ handle: string; sheet: string }> {
    const dir = mkdtempSync(join(tmpdir(), "mcp-xlsx-"));
    const path = copyXlsxFixture(name, dir);
    const out = structured(await client.callTool({ name: "xlsx_load", arguments: { path } }));
    const summary = out.summary as { sheets: Array<{ name: string; kind: string }> };
    const sheet = summary.sheets.find((s) => s.kind === "worksheet")!.name;
    return { handle: out.handle as string, sheet };
  }

  it("xlsx_load returns a handle and an inspection summary", async () => {
    const client = await makeClient();
    const dir = mkdtempSync(join(tmpdir(), "mcp-xlsx-load-"));
    const path = copyXlsxFixture("01-single-sheet-numbers.xlsx", dir);
    const out = structured(await client.callTool({ name: "xlsx_load", arguments: { path } }));
    expect(typeof out.handle).toBe("string");
    expect(out.path).toBe(path);
    const summary = out.summary as { format: string; sheets: unknown[] };
    expect(summary.format).toBe("xlsx");
    expect(Array.isArray(summary.sheets)).toBe(true);
  });

  it("xlsx_apply_command with xlsx:set-cell-value updates the snapshot", async () => {
    const client = await makeClient();
    const { handle, sheet } = await loadFixture(client, "01-single-sheet-numbers.xlsx");
    const apply = structured(
      await client.callTool({
        name: "xlsx_apply_command",
        arguments: {
          handle,
          type: "xlsx:set-cell-value",
          payload: { sheet, ref: "AA50", value: "via-apply-command" },
        },
      })
    );
    expect((apply.mutation as { status: string }).status).toBe("approved");

    const projection = structured(
      await client.callTool({
        name: "xlsx_get_text",
        arguments: { handle, format: "json", sheet, range: "AA50:AA50" },
      })
    );
    expect((projection.cells as Array<{ value: unknown }>)[0].value).toBe("via-apply-command");
  });

  it("xlsx_set_cell convenience tool produces an equivalent mutation", async () => {
    const client = await makeClient();
    const { handle, sheet } = await loadFixture(client, "01-single-sheet-numbers.xlsx");
    const apply = structured(
      await client.callTool({
        name: "xlsx_set_cell",
        arguments: { handle, sheet, ref: "AB1", value: 99 },
      })
    );
    expect((apply.mutation as { status: string }).status).toBe("approved");

    const projection = structured(
      await client.callTool({
        name: "xlsx_get_text",
        arguments: { handle, format: "json", sheet, range: "AB1:AB1" },
      })
    );
    expect((projection.cells as Array<{ value: unknown }>)[0].value).toBe(99);
  });

  it("xlsx_undo reverts the last mutation; xlsx_redo reapplies it", async () => {
    const client = await makeClient();
    const { handle, sheet } = await loadFixture(client, "01-single-sheet-numbers.xlsx");

    // Establish a known "after" value so the round-trip is observable.
    await client.callTool({
      name: "xlsx_set_cell",
      arguments: { handle, sheet, ref: "AD1", value: 42 },
    });
    const before = structured(
      await client.callTool({
        name: "xlsx_get_text",
        arguments: { handle, format: "json", sheet, range: "AD1:AD1" },
      })
    );
    expect((before.cells as Array<{ value: unknown }>)[0].value).toBe(42);

    const undone = structured(await client.callTool({ name: "xlsx_undo", arguments: { handle } }));
    expect(undone.did_undo).toBe(true);
    expect(undone.can_redo).toBe(true);
    const after = structured(
      await client.callTool({
        name: "xlsx_get_text",
        arguments: { handle, format: "json", sheet, range: "AD1:AD1" },
      })
    );
    // Cell is now empty; the projection drops the cell entirely.
    expect((after.cells as Array<unknown>).length).toBe(0);

    const redone = structured(await client.callTool({ name: "xlsx_redo", arguments: { handle } }));
    expect(redone.did_redo).toBe(true);
    const restored = structured(
      await client.callTool({
        name: "xlsx_get_text",
        arguments: { handle, format: "json", sheet, range: "AD1:AD1" },
      })
    );
    expect((restored.cells as Array<{ value: unknown }>)[0].value).toBe(42);

    // Empty undo / redo are no-ops, not errors.
    await client.callTool({ name: "xlsx_redo", arguments: { handle } });
    const noopRedo = structured(await client.callTool({ name: "xlsx_redo", arguments: { handle } }));
    expect(noopRedo.did_redo).toBe(false);
  });

  it("xlsx_save writes the modified workbook to disk", async () => {
    const client = await makeClient();
    const { handle, sheet } = await loadFixture(client, "01-single-sheet-numbers.xlsx");
    await client.callTool({
      name: "xlsx_set_cell",
      arguments: { handle, sheet, ref: "AC1", value: "saved!" },
    });
    const dir = mkdtempSync(join(tmpdir(), "mcp-xlsx-save-"));
    const out = join(dir, "out.xlsx");
    const saved = structured(
      await client.callTool({ name: "xlsx_save", arguments: { handle, out_path: out } })
    );
    expect(saved.wrote).toBe(out);
    const reloaded = await XlsxAgent.fromBuffer(readFileSync(out));
    const reloadedSheet = reloaded.getSnapshot().root.sheets.find((s) => s.name === sheet)!;
    const matched = [...reloadedSheet.cells.values()].find((c) => c.value === "saved!");
    expect(matched).toBeDefined();
  });

  it("xlsx_diff works for both handle-pair and disk modes", async () => {
    const client = await makeClient();
    const dir = mkdtempSync(join(tmpdir(), "mcp-xlsx-diff-"));
    const path = copyXlsxFixture("01-single-sheet-numbers.xlsx", dir);
    const a = structured(await client.callTool({ name: "xlsx_load", arguments: { path } }));
    const b = structured(await client.callTool({ name: "xlsx_load", arguments: { path } }));
    const sheet = (a.summary as { sheets: Array<{ name: string; kind: string }> }).sheets.find(
      (s) => s.kind === "worksheet"
    )!.name;

    await client.callTool({
      name: "xlsx_set_cell",
      arguments: { handle: b.handle, sheet, ref: "AD1", value: "diff" },
    });

    const handlePairDiff = structured(
      await client.callTool({
        name: "xlsx_diff",
        arguments: { before: a.handle, after: b.handle },
      })
    );
    expect(handlePairDiff.format).toBe("xlsx");
    expect((handlePairDiff.changes as unknown[]).length).toBeGreaterThanOrEqual(1);

    const diskDiff = structured(
      await client.callTool({
        name: "xlsx_diff",
        arguments: { handle: b.handle, against: "disk" },
      })
    );
    expect(diskDiff.format).toBe("xlsx");
    expect((diskDiff.changes as unknown[]).length).toBeGreaterThanOrEqual(1);
  });

  it("xlsx pending review flow: list, approve, reject", async () => {
    const client = await makeClient();
    const { handle, sheet } = await loadFixture(client, "01-single-sheet-numbers.xlsx");

    const apply = structured(
      await client.callTool({
        name: "xlsx_apply_command",
        arguments: {
          handle,
          type: "xlsx:set-cell-value",
          payload: { sheet, ref: "AE1", value: "pending" },
          auto_approve: false,
        },
      })
    );
    expect((apply.mutation as { status: string }).status).toBe("pending");
    const id = (apply.mutation as { id: string }).id;

    const list1 = structured(await client.callTool({ name: "xlsx_list_pending", arguments: { handle } }));
    expect((list1.pending as Array<{ id: string }>).map((p) => p.id)).toContain(id);

    const approve = structured(
      await client.callTool({ name: "xlsx_approve", arguments: { handle, mutation_id: id } })
    );
    expect(approve.approved).toBe(id);

    const list2 = structured(await client.callTool({ name: "xlsx_list_pending", arguments: { handle } }));
    expect(list2.pending as unknown[]).toHaveLength(0);

    const apply2 = structured(
      await client.callTool({
        name: "xlsx_apply_command",
        arguments: {
          handle,
          type: "xlsx:set-cell-value",
          payload: { sheet, ref: "AF1", value: "REJECTME" },
          auto_approve: false,
        },
      })
    );
    const id2 = (apply2.mutation as { id: string }).id;
    const rej = structured(
      await client.callTool({
        name: "xlsx_reject",
        arguments: { handle, mutation_id: id2, reason: "nope" },
      })
    );
    expect(rej.rejected).toBe(id2);
    expect(rej.reason).toBe("nope");
  });

  it("xlsx_list_sheets returns sheets in tab order", async () => {
    const client = await makeClient();
    const dir = mkdtempSync(join(tmpdir(), "mcp-xlsx-list-"));
    const path = copyXlsxFixture("02-multi-sheet.xlsx", dir);
    const loaded = structured(await client.callTool({ name: "xlsx_load", arguments: { path } }));
    const list = structured(
      await client.callTool({ name: "xlsx_list_sheets", arguments: { handle: loaded.handle } })
    );
    const sheets = list.sheets as Array<{ name: string; index: number }>;
    expect(sheets.length).toBeGreaterThan(1);
    const indexes = sheets.map((s) => s.index);
    expect(indexes).toEqual([...indexes].sort((a, b) => a - b));
  });

  it("returns an error for unknown xlsx handles", async () => {
    const client = await makeClient();
    const r = await client.callTool({
      name: "xlsx_inspect",
      arguments: { handle: "nope" },
    });
    expect(r.isError).toBe(true);
  });
});
