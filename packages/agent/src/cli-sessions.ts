import { randomUUID } from "node:crypto";
import { readFile, writeFile } from "node:fs/promises";
import { basename, extname, resolve } from "node:path";
import { Command, Option } from "commander";
import { DocxAgent } from "@officeai/docx";
import { PdfAgent } from "@officeai/pdf";
import { PptxAgent } from "@officeai/pptx";
import { XlsxAgent } from "@officeai/xlsx";
import { CliError, parseIntOpt, stringifyJson, type IO } from "./cli-shared.js";
import {
  createLocalSessionStore,
  type StoredCommandLogEntry,
  type StoredDiagnostic,
  type StoredDocumentRecord,
  type StoredOfficeFormat,
  type StoredSessionRecord,
} from "./session-store.js";
import {
  projectOfficeDocument,
  projectionDocumentEnvelope,
  type ProjectionKind,
  type ProjectionSource,
} from "./projections.js";

type EditableStoredOfficeFormat = Exclude<StoredOfficeFormat, "image">;

interface SessionCliOutputOptions {
  readonly json?: boolean;
  readonly pretty?: boolean;
  readonly quiet?: boolean;
}

interface SessionCliStoreOptions extends SessionCliOutputOptions {
  readonly dataDir?: string;
}

export function registerSessionSubcommands(sessions: Command, io: IO): void {
  sessions
    .command("create")
    .description("Create a path-free local OfficeAI session and print its sessionId.")
    .option("--title <title>", "Human-readable session title.", "CLI session")
    .option("--json", "Emit machine-readable JSON.", false)
    .option("--pretty", "Pretty-print JSON output.", false)
    .option("--quiet", "Suppress stdout; useful when only the exit code matters.", false)
    .option("--data-dir <path>", "Override OFFICEAI_DATA_DIR.")
    .action(async (opts: SessionCliStoreOptions & { title: string }) => {
      const store = createLocalSessionStore({ dataDir: opts.dataDir });
      const now = new Date().toISOString();
      const session = {
        id: `session_${randomUUID()}`,
        title: opts.title,
        createdAt: now,
        updatedAt: now,
        documentIds: [],
      };
      await store.putSession(session);
      writeSessionCliOutput(
        io,
        opts,
        {
          schema: "office-ai/session@1",
          surface: "cli",
          dataDir: store.dataDir,
          session: sessionPayload(session),
          nextActions: ["office-agent sessions import", "office-agent sessions documents"],
        },
        `Created ${session.id} (${session.title})`
      );
    });

  sessions
    .command("list")
    .description("List local OfficeAI sessions from the canonical data-dir store.")
    .option("--json", "Emit machine-readable JSON.", false)
    .option("--pretty", "Pretty-print JSON output.", false)
    .option("--quiet", "Suppress stdout; useful when only the exit code matters.", false)
    .option("--data-dir <path>", "Override OFFICEAI_DATA_DIR.")
    .action(async (opts: SessionCliStoreOptions) => {
      const store = createLocalSessionStore({ dataDir: opts.dataDir });
      const records = await store.listSessions();
      writeSessionCliOutput(
        io,
        opts,
        {
          schema: "office-ai/session-list@1",
          surface: "cli",
          dataDir: store.dataDir,
          sessions: records.map(sessionPayload),
        },
        records.length === 0
          ? `No OfficeAI sessions in ${store.dataDir}`
          : records
              .map((session) => `${session.id}\t${session.documentIds.length}\t${session.title}`)
              .join("\n")
      );
    });

  sessions
    .command("documents")
    .description("List canonical OfficeAI documents, optionally restricted to one session.")
    .option("--session-id <sessionId>", "Restrict to one sessionId.")
    .option("--json", "Emit machine-readable JSON.", false)
    .option("--pretty", "Pretty-print JSON output.", false)
    .option("--quiet", "Suppress stdout; useful when only the exit code matters.", false)
    .option("--data-dir <path>", "Override OFFICEAI_DATA_DIR.")
    .action(async (opts: SessionCliStoreOptions & { sessionId?: string }) => {
      const store = createLocalSessionStore({ dataDir: opts.dataDir });
      const documents = await store.listDocuments(opts.sessionId);
      writeSessionCliOutput(
        io,
        opts,
        {
          schema: "office-ai/document-list@1",
          surface: "cli",
          dataDir: store.dataDir,
          documents: documents.map(documentPayload),
        },
        documents.length === 0
          ? `No OfficeAI documents in ${opts.sessionId ?? store.dataDir}`
          : documents.map((doc) => `${doc.id}\t${doc.sessionId}\t${doc.format}\t${doc.name}`).join("\n")
      );
    });

  sessions
    .command("import")
    .description(
      "Import a DOCX/XLSX/PPTX/PDF file as a canonical session document. The path is an edge input; the output is sessionId/documentId."
    )
    .requiredOption("--file <path>", "Path to a .docx, .xlsx, .pptx or .pdf file.")
    .option("--session-id <sessionId>", "Append to an existing sessionId.")
    .option("--title <title>", "Title for a newly created session.", "CLI imports")
    .option("--name <name>", "Display name for the imported document.")
    .addOption(
      new Option("--format <fmt>", "Override extension-based format detection.").choices([
        "docx",
        "xlsx",
        "pptx",
        "pdf",
      ])
    )
    .option("--json", "Emit machine-readable JSON.", false)
    .option("--pretty", "Pretty-print JSON output.", false)
    .option("--quiet", "Suppress stdout; useful when only the exit code matters.", false)
    .option("--data-dir <path>", "Override OFFICEAI_DATA_DIR.")
    .action(
      async (
        opts: SessionCliStoreOptions & {
          file: string;
          sessionId?: string;
          title: string;
          name?: string;
          format?: EditableStoredOfficeFormat;
        }
      ) => {
        const store = createLocalSessionStore({ dataDir: opts.dataDir });
        const abs = resolve(opts.file);
        const format = inferSessionFormatFromPath(abs, opts.format);
        const bytes = new Uint8Array(await readFile(abs));
        const prepared = await prepareSessionBytes(format, bytes);
        const now = new Date().toISOString();
        const documentId = `doc_${randomUUID()}`;
        const session = await sessionForCliDocument({
          store,
          sessionId: opts.sessionId,
          documentId,
          title: opts.title,
          now,
        });
        const displayName = ensureSessionExtension(opts.name ?? basename(abs), format);
        const diagnostics: StoredDiagnostic[] = [
          { level: "info", code: "imported", message: `Imported ${displayName} as ${format}.` },
        ];
        const commandLog = [
          cliCommandLogEntry({
            operation: "import_document",
            status: "applied",
            stage: "imported",
            now,
            diagnostics,
            sessionId: session.id,
            documentId,
            revision: prepared.revision,
            argumentsSummary: `--file ${basename(abs)} --format ${format}`,
          }),
        ];

        await store.putSession(session);
        const document = await store.putDocument(
          {
            id: documentId,
            sessionId: session.id,
            format,
            name: displayName,
            sourcePath: abs,
            status: "ready",
            createdAt: now,
            updatedAt: now,
            revision: prepared.revision,
            diagnostics,
            exportHistory: [],
            pendingChanges: [],
            commandLog,
          },
          { originalBytes: bytes, workingBytes: bytes }
        );
        const persistedSession = await store.getSession(session.id);
        writeSessionCliOutput(
          io,
          opts,
          {
            schema: "office-ai/import-document@1",
            surface: "cli",
            dataDir: store.dataDir,
            session: sessionPayload(persistedSession),
            document: documentPayload(document),
            summary: projectOfficeDocument(documentProjectionMeta(document), prepared.source, {
              projection: "summary",
            }),
            nextActions: ["office-agent sessions projection", "office-agent sessions export"],
          },
          `Imported ${displayName} as ${document.id} in ${session.id}`
        );
      }
    );

  sessions
    .command("projection")
    .description(
      "Read a canonical document projection from the local session store without passing file paths."
    )
    .requiredOption(
      "--document-id <documentId>",
      "DocumentId returned by sessions import or MCP import_document."
    )
    .addOption(
      new Option("--projection <kind>", "Projection kind.")
        .choices(["summary", "markdown", "json", "text", "page"])
        .default("markdown")
    )
    .option("--page <n>", "1-based DOCX/PDF page for projection=page.", parseIntOpt)
    .option("--sheet <name>", "XLSX sheet name for markdown/json projections.")
    .option("--range <a1>", "XLSX A1 range for json projections; requires --sheet.")
    .option("--slide <n>", "0-based PPTX slide for json/page projections.", parseIntOpt)
    .option("--max-rows <n>", "XLSX markdown row limit.", parseIntOpt)
    .option("--max-cols <n>", "XLSX markdown column limit.", parseIntOpt)
    .option("--json", "Emit the full machine-readable projection envelope.", false)
    .option("--pretty", "Pretty-print JSON output.", false)
    .option("--quiet", "Suppress stdout; useful when only the exit code matters.", false)
    .option("--data-dir <path>", "Override OFFICEAI_DATA_DIR.")
    .action(
      async (
        opts: SessionCliStoreOptions & {
          documentId: string;
          projection: ProjectionKind;
          page?: number;
          sheet?: string;
          range?: string;
          slide?: number;
          maxRows?: number;
          maxCols?: number;
        }
      ) => {
        const store = createLocalSessionStore({ dataDir: opts.dataDir });
        const document = await requireStoredDocument(store, opts.documentId);
        if (!isEditableStoredOfficeFormat(document.format)) {
          throw new CliError(
            64,
            `sessions projection: ${document.format} documents use viewer routes and do not expose CLI projections.`
          );
        }
        const prepared = await prepareSessionBytes(document.format, await store.readWorkingBytes(document));
        const payload = projectOfficeDocument(documentProjectionMeta(document), prepared.source, {
          projection: opts.projection,
          ...(opts.page !== undefined ? { page: opts.page } : {}),
          ...(opts.sheet !== undefined ? { sheet: opts.sheet } : {}),
          ...(opts.range !== undefined ? { range: opts.range } : {}),
          ...(opts.slide !== undefined ? { slide: opts.slide } : {}),
          ...(opts.maxRows !== undefined ? { maxRows: opts.maxRows } : {}),
          ...(opts.maxCols !== undefined ? { maxCols: opts.maxCols } : {}),
        });
        const content = (payload as { content?: unknown }).content;
        writeSessionCliOutput(
          io,
          opts,
          payload,
          typeof content === "string" && opts.projection !== "json"
            ? content
            : stringifyJson(payload, opts.pretty === true)
        );
      }
    );

  sessions
    .command("export")
    .description(
      "Export a canonical session document to a local file. The path is an explicit edge output and is recorded in export history."
    )
    .requiredOption(
      "--document-id <documentId>",
      "DocumentId returned by sessions import or MCP import_document."
    )
    .requiredOption("--out <path>", "Path to write the exported DOCX/XLSX/PPTX/PDF file.")
    .option("--json", "Emit machine-readable JSON.", false)
    .option("--pretty", "Pretty-print JSON output.", false)
    .option("--quiet", "Suppress stdout; useful when only the exit code matters.", false)
    .option("--data-dir <path>", "Override OFFICEAI_DATA_DIR.")
    .action(async (opts: SessionCliStoreOptions & { documentId: string; out: string }) => {
      const store = createLocalSessionStore({ dataDir: opts.dataDir });
      const document = await requireStoredDocument(store, opts.documentId);
      const bytes = await store.readWorkingBytes(document);
      const target = resolve(opts.out);
      await writeFile(target, Buffer.from(bytes));
      const now = new Date().toISOString();
      const exported = {
        path: target,
        bytes: bytes.byteLength,
        exportedAt: now,
      };
      const basis = document.commandLog
        .filter((entry) => entry.stage !== "exported")
        .map((entry) => entry.commandId ?? entry.id)
        .slice(-20);
      const diagnostics: StoredDiagnostic[] = [
        { level: "info", code: "exported", message: `Exported ${document.name} as ${document.format}.` },
        {
          level: "info",
          code: "export-command-basis",
          message:
            basis.length > 0
              ? `Export command basis: ${basis.join(", ")}.`
              : "Export command basis is empty.",
        },
      ];
      const updated = await store.putDocument(
        storedDocumentInput(document, {
          updatedAt: now,
          diagnostics,
          exportHistory: [...document.exportHistory, exported],
          commandLog: [
            ...document.commandLog,
            cliCommandLogEntry({
              operation: "export_document",
              status: "exported",
              stage: "exported",
              now,
              diagnostics,
              sessionId: document.sessionId,
              documentId: document.id,
              revision: document.revision,
              exportRef: {
                exportedAt: exported.exportedAt,
                bytes: exported.bytes,
                commandIds: basis,
              },
            }),
          ],
        })
      );
      const session = await store.getSession(document.sessionId);
      await store.putSession({
        id: session.id,
        title: session.title,
        createdAt: session.createdAt,
        updatedAt: now,
        documentIds: session.documentIds,
      });
      writeSessionCliOutput(
        io,
        opts,
        {
          schema: "office-ai/export-document@1",
          surface: "cli",
          dataDir: store.dataDir,
          document: documentPayload(updated),
          exported,
          diagnostics,
        },
        `Exported ${document.id} to ${target}`
      );
    });

  sessions
    .command("inspect")
    .description("Inspect session-store schema state without mutating documents.")
    .option("--json", "Emit machine-readable JSON.", false)
    .option("--pretty", "Pretty-print JSON output.", false)
    .option("--data-dir <path>", "Override OFFICEAI_DATA_DIR.")
    .action(async (opts: Record<string, unknown>) => {
      const store = createLocalSessionStore({
        dataDir: typeof opts.dataDir === "string" ? opts.dataDir : undefined,
      });
      const result = await store.inspectDataDir();
      if (opts.json === true) {
        io.stdout.write(stringifyJson(result, opts.pretty === true) + "\n");
      } else {
        io.stdout.write(renderSessionMaintenance("office-ai sessions inspect", result.diagnostics) + "\n");
      }
    });
  sessions
    .command("migrate")
    .description("Migrate old local session-store metadata after creating backups.")
    .option("--json", "Emit machine-readable JSON.", false)
    .option("--pretty", "Pretty-print JSON output.", false)
    .option("--data-dir <path>", "Override OFFICEAI_DATA_DIR.")
    .action(async (opts: Record<string, unknown>) => {
      const store = createLocalSessionStore({
        dataDir: typeof opts.dataDir === "string" ? opts.dataDir : undefined,
      });
      const result = await store.migrateDataDir();
      if (opts.json === true) {
        io.stdout.write(stringifyJson(result, opts.pretty === true) + "\n");
      } else {
        io.stdout.write(renderSessionMaintenance("office-ai sessions migrate", result.diagnostics) + "\n");
      }
    });
  sessions
    .command("cleanup")
    .description("Remove temporary session-store files without deleting original or working artifacts.")
    .option("--json", "Emit machine-readable JSON.", false)
    .option("--pretty", "Pretty-print JSON output.", false)
    .option("--data-dir <path>", "Override OFFICEAI_DATA_DIR.")
    .action(async (opts: Record<string, unknown>) => {
      const store = createLocalSessionStore({
        dataDir: typeof opts.dataDir === "string" ? opts.dataDir : undefined,
      });
      const result = await store.cleanupTemporaryArtifacts();
      if (opts.json === true) {
        io.stdout.write(stringifyJson(result, opts.pretty === true) + "\n");
      } else {
        io.stdout.write(renderSessionMaintenance("office-ai sessions cleanup", result.diagnostics) + "\n");
      }
    });
}

function renderSessionMaintenance(
  title: string,
  diagnostics: ReadonlyArray<{ readonly level: string; readonly code: string; readonly message: string }>
): string {
  const lines = [title];
  for (const diagnostic of diagnostics) {
    lines.push(`${diagnostic.level.toUpperCase().padEnd(8)} ${diagnostic.code} - ${diagnostic.message}`);
  }
  return lines.join("\n");
}

function writeSessionCliOutput(io: IO, opts: SessionCliOutputOptions, payload: unknown, human: string): void {
  if (opts.quiet === true) return;
  io.stdout.write((opts.json === true ? stringifyJson(payload, opts.pretty === true) : human) + "\n");
}

function sessionPayload(
  session: StoredSessionRecord | Omit<StoredSessionRecord, "schema" | "schemaVersion" | "version" | "lease">
): Record<string, unknown> {
  return {
    schema: "office-ai/session@1",
    sessionId: session.id,
    title: session.title,
    createdAt: session.createdAt,
    updatedAt: session.updatedAt,
    documentCount: session.documentIds.length,
    documentIds: session.documentIds,
  };
}

function documentPayload(document: StoredDocumentRecord): Record<string, unknown> {
  return projectionDocumentEnvelope(documentProjectionMeta(document));
}

function documentProjectionMeta(document: StoredDocumentRecord): {
  readonly documentId: string;
  readonly sessionId: string;
  readonly format: StoredOfficeFormat;
  readonly name: string;
  readonly status: "ready" | "error";
  readonly revision: number;
  readonly sourcePath?: string;
  readonly diagnostics: ReadonlyArray<unknown>;
  readonly exportHistory: ReadonlyArray<unknown>;
  readonly createdAt: string;
  readonly updatedAt: string;
} {
  return {
    documentId: document.id,
    sessionId: document.sessionId,
    format: document.format,
    name: document.name,
    status: document.status,
    revision: document.revision,
    ...(document.sourcePath ? { sourcePath: document.sourcePath } : {}),
    diagnostics: document.diagnostics,
    exportHistory: document.exportHistory,
    createdAt: document.createdAt,
    updatedAt: document.updatedAt,
  };
}

async function sessionForCliDocument(opts: {
  readonly store: ReturnType<typeof createLocalSessionStore>;
  readonly sessionId?: string;
  readonly documentId: string;
  readonly title: string;
  readonly now: string;
}): Promise<Omit<StoredSessionRecord, "schema" | "schemaVersion" | "version" | "lease">> {
  if (opts.sessionId) {
    const existing = await opts.store.getSession(opts.sessionId);
    return {
      id: existing.id,
      title: existing.title,
      createdAt: existing.createdAt,
      updatedAt: opts.now,
      documentIds: [...new Set([...existing.documentIds, opts.documentId])],
    };
  }
  return {
    id: `session_${randomUUID()}`,
    title: opts.title,
    createdAt: opts.now,
    updatedAt: opts.now,
    documentIds: [opts.documentId],
  };
}

function inferSessionFormatFromPath(
  path: string,
  explicit?: EditableStoredOfficeFormat
): EditableStoredOfficeFormat {
  if (explicit) return explicit;
  const ext = extname(path).toLowerCase();
  switch (ext) {
    case ".docx":
      return "docx";
    case ".xlsx":
      return "xlsx";
    case ".pptx":
      return "pptx";
    case ".pdf":
      return "pdf";
    default:
      throw new CliError(
        64,
        `sessions import: cannot infer document format from extension "${ext || "<none>"}"; pass --format`
      );
  }
}

function ensureSessionExtension(name: string, format: EditableStoredOfficeFormat): string {
  return name.toLowerCase().endsWith(`.${format}`) ? name : `${name}.${format}`;
}

function isEditableStoredOfficeFormat(format: StoredOfficeFormat): format is EditableStoredOfficeFormat {
  return format === "docx" || format === "xlsx" || format === "pptx" || format === "pdf";
}

async function prepareSessionBytes(
  format: EditableStoredOfficeFormat,
  bytes: Uint8Array
): Promise<{ readonly source: ProjectionSource; readonly revision: number }> {
  switch (format) {
    case "docx": {
      const agent = await DocxAgent.fromBuffer(bytes);
      return { source: { format, agent }, revision: agent.getSnapshot().revision };
    }
    case "xlsx": {
      const agent = await XlsxAgent.fromBuffer(bytes);
      return { source: { format, agent }, revision: agent.getSnapshot().revision };
    }
    case "pptx": {
      const agent = await PptxAgent.fromBuffer(bytes);
      return { source: { format, agent }, revision: agent.getSnapshot().revision };
    }
    case "pdf": {
      const agent = await PdfAgent.fromBuffer(bytes);
      return { source: { format, agent }, revision: agent.getSnapshot().revision };
    }
  }
}

async function requireStoredDocument(
  store: ReturnType<typeof createLocalSessionStore>,
  documentId: string
): Promise<StoredDocumentRecord> {
  const document = await store.getDocument(documentId);
  if (!document) {
    throw new CliError(2, `Unknown document-id "${documentId}". Run office-agent sessions documents first.`);
  }
  return document;
}

function cliCommandLogEntry(args: {
  readonly operation: string;
  readonly status: string;
  readonly stage: string;
  readonly now: string;
  readonly diagnostics?: ReadonlyArray<StoredDiagnostic>;
  readonly sessionId: string;
  readonly documentId: string;
  readonly revision: number;
  readonly argumentsSummary?: string;
  readonly exportRef?: StoredCommandLogEntry["exportRef"];
}): StoredCommandLogEntry {
  const id = `log_${randomUUID()}`;
  return {
    schema: "office-ai/audit-log-entry@1",
    schemaVersion: 1,
    id,
    commandId: id,
    operation: args.operation,
    status: args.status,
    stage: args.stage,
    source: "cli",
    recordedAt: args.now,
    ...(args.diagnostics ? { diagnostics: args.diagnostics } : {}),
    provenance: {
      surface: "cli",
      sessionId: args.sessionId,
      documentId: args.documentId,
      targetRevision: args.revision,
      ...(args.argumentsSummary ? { argumentsSummary: args.argumentsSummary } : {}),
    },
    ...(args.exportRef ? { exportRef: args.exportRef } : {}),
  };
}

function storedDocumentInput(
  document: StoredDocumentRecord,
  overrides: Partial<Omit<StoredDocumentRecord, "schema" | "schemaVersion" | "version" | "artifacts">>
): Omit<StoredDocumentRecord, "schema" | "schemaVersion" | "version" | "artifacts"> {
  return {
    id: overrides.id ?? document.id,
    sessionId: overrides.sessionId ?? document.sessionId,
    format: overrides.format ?? document.format,
    name: overrides.name ?? document.name,
    status: overrides.status ?? document.status,
    ...((overrides.sourcePath ?? document.sourcePath)
      ? { sourcePath: overrides.sourcePath ?? document.sourcePath }
      : {}),
    createdAt: overrides.createdAt ?? document.createdAt,
    updatedAt: overrides.updatedAt ?? document.updatedAt,
    revision: overrides.revision ?? document.revision,
    diagnostics: overrides.diagnostics ?? document.diagnostics,
    exportHistory: overrides.exportHistory ?? document.exportHistory,
    pendingChanges: overrides.pendingChanges ?? document.pendingChanges,
    commandLog: overrides.commandLog ?? document.commandLog,
  };
}
