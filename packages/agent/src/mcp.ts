/**
 * OfficeAI Model Context Protocol server.
 *
 * Exposes OfficeAI document sessions as generic MCP tools, plus the legacy
 * format-specific tools. Sessions live in-process today; canonical
 * document IDs are stable across tool calls in one server lifetime and also
 * work as handles for the matching docx_* / xlsx_* / pptx_* / pdf_* tools.
 *
 * Transport-agnostic: `runMcpStdioServer()` wires up `StdioServerTransport`
 * for the published binary, but tests use `InMemoryTransport` directly via
 * the exported `createMcpServer()` factory.
 */

import { readFile, writeFile } from "node:fs/promises";
import { basename, extname, resolve } from "node:path";
import { randomUUID } from "node:crypto";
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { z } from "zod";
import {
  createCommandEnvelope,
  previewCommandEnvelope,
  validateCommandEnvelope,
  type ActionDescriptor,
  type CommandDiagnostic,
  type CommandEnvelope,
  type CommandHandler,
  type CommandPolicyMode,
  type CommandSource,
  type DocumentSnapshot,
  type Mutation,
} from "@officeai/core";
import { DocxAgent, allDocxHandlers, docxActions } from "@officeai/docx";
import { XlsxAgent, allXlsxHandlers, diffXlsxSnapshots, xlsxActions } from "@officeai/xlsx";
import { PptxAgent, allPptxHandlers, pptxActions } from "@officeai/pptx";
import { PdfAgent, allPdfHandlers, pdfActions } from "@officeai/pdf";
import { registerActionsAsMcpTools, type McpDispatchContext } from "./actions-to-mcp.js";
import {
  addPageNumbers,
  addWatermark,
  deletePages,
  extractPages,
  mergePdfs,
  reorderPages,
  rotatePages,
  setMetadata,
} from "@officeai/pdf-edit";
import { fillForm, flattenForm, resetForm } from "@officeai/pdf-forms";
import {
  projectAnnotations,
  projectFormFields,
  projectMetadata,
  projectOutline,
  projectPage,
  projectSearch,
} from "./pdf-cli.js";
import { diffSnapshots, inspectSnapshot, snapshotToJsonProjection } from "./cli.js";
import { inspectXlsxSnapshot, xlsxRangeToJson } from "./cli-xlsx.js";
import {
  diffSnapshots as pptxDiffSnapshots,
  inspectSnapshot as pptxInspectSnapshot,
  snapshotToJsonProjection as pptxSnapshotToJsonProjection,
} from "./pptx-cli.js";

const sessions = new Map<string, DocxAgent>();
const sessionPaths = new Map<string, string>();
const xlsxSessions = new Map<string, XlsxAgent>();
const xlsxSessionPaths = new Map<string, string>();
const pptxSessions = new Map<string, PptxAgent>();
const pptxSessionPaths = new Map<string, string>();
const pdfSessions = new Map<string, { agent: PdfAgent; bytes: Uint8Array }>();
const pdfSessionPaths = new Map<string, string>();

type OfficeFormat = "docx" | "xlsx" | "pptx" | "pdf";
type DocumentStatus = "ready" | "error";
type OfficeAgent = DocxAgent | XlsxAgent | PptxAgent | PdfAgent;
type AnyMutation = Mutation<DocumentSnapshot>;

interface McpDiagnostic {
  readonly level: "info" | "warning" | "error" | "destructive";
  readonly code: string;
  readonly message: string;
}

interface ExportRecord {
  readonly path: string;
  readonly bytes: number;
  readonly exportedAt: string;
}

interface SessionRecord {
  readonly id: string;
  title: string;
  readonly createdAt: string;
  updatedAt: string;
  readonly documentIds: Set<string>;
}

interface DocumentRecord {
  readonly id: string;
  readonly sessionId: string;
  readonly format: OfficeFormat;
  name: string;
  sourcePath?: string;
  readonly createdAt: string;
  updatedAt: string;
  status: DocumentStatus;
  diagnostics: McpDiagnostic[];
  readonly exportHistory: ExportRecord[];
}

const mcpSessions = new Map<string, SessionRecord>();
const mcpDocuments = new Map<string, DocumentRecord>();
const plannedCommands = new Map<string, CommandEnvelope>();

/** Test hook: drop in-memory state between test cases. */
export function __resetMcpSessionsForTests(): void {
  sessions.clear();
  sessionPaths.clear();
  xlsxSessions.clear();
  xlsxSessionPaths.clear();
  pptxSessions.clear();
  pptxSessionPaths.clear();
  pdfSessions.clear();
  pdfSessionPaths.clear();
  mcpSessions.clear();
  mcpDocuments.clear();
  plannedCommands.clear();
}

function lookupAgent(handle: string): DocxAgent {
  const agent = sessions.get(handle);
  if (!agent) {
    throw new Error(`Unknown DOCX handle: "${handle}". Call docx_load first.`);
  }
  return agent;
}

function lookupXlsxAgent(handle: string): XlsxAgent {
  const agent = xlsxSessions.get(handle);
  if (!agent) {
    throw new Error(`Unknown XLSX handle: "${handle}". Call xlsx_load first.`);
  }
  return agent;
}

function lookupPptxAgent(handle: string): PptxAgent {
  const agent = pptxSessions.get(handle);
  if (!agent) {
    throw new Error(`Unknown PPTX handle: "${handle}". Call pptx_load first.`);
  }
  return agent;
}

function lookupPdfSession(handle: string): { agent: PdfAgent; bytes: Uint8Array } {
  const session = pdfSessions.get(handle);
  if (!session) {
    throw new Error(`Unknown PDF handle: "${handle}". Call pdf_load first.`);
  }
  return session;
}

function ok(payload: unknown): {
  content: Array<{ type: "text"; text: string }>;
  structuredContent: Record<string, unknown>;
} {
  const text = JSON.stringify(payload, null, 2);
  return {
    content: [{ type: "text", text }],
    structuredContent: payload as Record<string, unknown>,
  };
}

function fail(message: string): {
  isError: true;
  content: Array<{ type: "text"; text: string }>;
} {
  return {
    isError: true,
    content: [{ type: "text", text: message }],
  };
}

function nowIso(): string {
  return new Date().toISOString();
}

function createSessionRecord(title = "OfficeAI session"): SessionRecord {
  const ts = nowIso();
  const record: SessionRecord = {
    id: `session_${randomUUID()}`,
    title,
    createdAt: ts,
    updatedAt: ts,
    documentIds: new Set(),
  };
  mcpSessions.set(record.id, record);
  return record;
}

function getOrCreateSession(sessionId?: string, title?: string): SessionRecord {
  if (!sessionId) return createSessionRecord(title);
  const existing = mcpSessions.get(sessionId);
  if (existing) return existing;
  throw new Error(`Unknown session_id "${sessionId}". Call create_session first or omit session_id.`);
}

function ensureDefaultSession(): SessionRecord {
  const first = mcpSessions.values().next();
  if (!first.done) return first.value;
  return createSessionRecord();
}

function touchSession(session: SessionRecord): void {
  session.updatedAt = nowIso();
}

function inferFormatFromPath(path: string, requested?: OfficeFormat): OfficeFormat {
  if (requested) return requested;
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
      throw new Error(`Cannot infer document format from extension "${ext || "<none>"}". Pass format.`);
  }
}

function registerDocumentRecord(opts: {
  readonly id: string;
  readonly sessionId?: string;
  readonly format: OfficeFormat;
  readonly name: string;
  readonly sourcePath?: string;
  readonly diagnostics?: McpDiagnostic[];
}): DocumentRecord {
  const session = opts.sessionId ? getOrCreateSession(opts.sessionId) : ensureDefaultSession();
  const existing = mcpDocuments.get(opts.id);
  const ts = nowIso();
  const record: DocumentRecord =
    existing ??
    ({
      id: opts.id,
      sessionId: session.id,
      format: opts.format,
      name: opts.name,
      sourcePath: opts.sourcePath,
      createdAt: ts,
      updatedAt: ts,
      status: "ready",
      diagnostics: [],
      exportHistory: [],
    } satisfies DocumentRecord);

  record.name = opts.name;
  record.sourcePath = opts.sourcePath;
  record.status = "ready";
  record.updatedAt = ts;
  record.diagnostics = opts.diagnostics ?? [];
  mcpDocuments.set(record.id, record);
  session.documentIds.add(record.id);
  touchSession(session);
  return record;
}

function lookupDocument(documentId: string): DocumentRecord {
  const record = mcpDocuments.get(documentId);
  if (!record) {
    throw new Error(`Unknown document_id "${documentId}". Call import_document or create_document first.`);
  }
  return record;
}

function revisionFor(record: DocumentRecord): number {
  switch (record.format) {
    case "docx":
      return lookupAgent(record.id).getSnapshot().revision;
    case "xlsx":
      return lookupXlsxAgent(record.id).getSnapshot().revision;
    case "pptx":
      return lookupPptxAgent(record.id).getSnapshot().revision;
    case "pdf":
      return lookupPdfSession(record.id).agent.getSnapshot().revision;
  }
}

function snapshotFor(record: DocumentRecord): DocumentSnapshot {
  switch (record.format) {
    case "docx":
      return lookupAgent(record.id).getSnapshot();
    case "xlsx":
      return lookupXlsxAgent(record.id).getSnapshot();
    case "pptx":
      return lookupPptxAgent(record.id).getSnapshot();
    case "pdf":
      return lookupPdfSession(record.id).agent.getSnapshot();
  }
}

function agentFor(record: DocumentRecord): OfficeAgent {
  switch (record.format) {
    case "docx":
      return lookupAgent(record.id);
    case "xlsx":
      return lookupXlsxAgent(record.id);
    case "pptx":
      return lookupPptxAgent(record.id);
    case "pdf":
      return lookupPdfSession(record.id).agent;
  }
}

function handlersFor(format: OfficeFormat): ReadonlyArray<CommandHandler<unknown, DocumentSnapshot>> {
  switch (format) {
    case "docx":
      return allDocxHandlers as ReadonlyArray<CommandHandler<unknown, DocumentSnapshot>>;
    case "xlsx":
      return allXlsxHandlers as ReadonlyArray<CommandHandler<unknown, DocumentSnapshot>>;
    case "pptx":
      return allPptxHandlers as ReadonlyArray<CommandHandler<unknown, DocumentSnapshot>>;
    case "pdf":
      return allPdfHandlers as ReadonlyArray<CommandHandler<unknown, DocumentSnapshot>>;
  }
}

function handlerFor(
  format: OfficeFormat,
  operation: string
): CommandHandler<unknown, DocumentSnapshot> | null {
  return handlersFor(format).find((handler) => handler.type === operation) ?? null;
}

function allActionDescriptors(): ReadonlyArray<ActionDescriptor> {
  return [...docxActions, ...xlsxActions, ...pptxActions, ...pdfActions];
}

function actionById(actionId: string): ActionDescriptor | undefined {
  return allActionDescriptors().find((action) => action.id === actionId);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function stringField(input: Record<string, unknown>, key: string): string | undefined {
  const value = input[key];
  return typeof value === "string" && value.length > 0 ? value : undefined;
}

function commandPolicyFrom(input: Record<string, unknown>): Partial<CommandEnvelope["policy"]> {
  const raw = input.policy;
  if (!isRecord(raw)) return {};
  const mode = raw.mode;
  const requiresReview = raw.requires_review ?? raw.requiresReview;
  return {
    ...(mode === "dry_run" || mode === "auto_apply" || mode === "pending"
      ? { mode: mode satisfies CommandPolicyMode }
      : {}),
    ...(typeof requiresReview === "boolean" ? { requiresReview } : {}),
  };
}

function commandSourceFrom(input: Record<string, unknown>): CommandEnvelope["source"] {
  const actorId = stringField(input, "actor_id") ?? stringField(input, "actorId");
  return {
    surface: "mcp",
    ...(actorId ? { actorId } : {}),
  };
}

function resolveCommandOperation(
  input: Record<string, unknown>,
  record: DocumentRecord
): {
  readonly operation: string;
  readonly args: unknown;
  readonly action?: Pick<
    ActionDescriptor,
    "id" | "label" | "requiresReview" | "supportsDiff" | "supportsDryRun"
  >;
} {
  const actionId = stringField(input, "action_id") ?? stringField(input, "actionId");
  const directOperation = stringField(input, "operation");
  const rawArgs = input.arguments ?? {};

  if (!actionId) {
    if (!directOperation) throw new Error("operation is required when action_id is not supplied.");
    return { operation: directOperation, args: rawArgs };
  }

  const action = actionById(actionId);
  if (!action) throw new Error(`Unknown action_id "${actionId}".`);
  if (action.format !== record.format) {
    throw new Error(
      `Action ${actionId} targets ${action.format}, but document ${record.id} is ${record.format}.`
    );
  }
  if (!action.commandType) {
    throw new Error(`Action ${actionId} does not map to a document command.`);
  }
  if (directOperation && directOperation !== action.commandType) {
    throw new Error(`Action ${actionId} maps to ${action.commandType}, not ${directOperation}.`);
  }

  const parsedArgs = isRecord(rawArgs) ? rawArgs : {};
  return {
    operation: action.commandType,
    args: action.buildPayload ? action.buildPayload(parsedArgs) : rawArgs,
    action: {
      id: action.id,
      label: action.label,
      requiresReview: action.requiresReview,
      supportsDiff: action.supportsDiff,
      supportsDryRun: action.supportsDryRun,
    },
  };
}

function commandEnvelopeFromInput(input: Record<string, unknown>): {
  readonly envelope: CommandEnvelope;
  readonly record: DocumentRecord;
  readonly action?: Pick<
    ActionDescriptor,
    "id" | "label" | "requiresReview" | "supportsDiff" | "supportsDryRun"
  >;
} {
  const documentId = stringField(input, "document_id") ?? stringField(input, "documentId");
  if (!documentId) throw new Error("document_id is required.");
  const record = lookupDocument(documentId);
  const resolved = resolveCommandOperation(input, record);
  const target = isRecord(input.target) ? input.target : {};
  const revision =
    typeof target.revision === "number" && Number.isInteger(target.revision)
      ? target.revision
      : revisionFor(record);
  const requestedFormat = stringField(input, "format") as OfficeFormat | undefined;
  const mode = commandPolicyFrom(input).mode;

  const envelope = createCommandEnvelope({
    id: stringField(input, "command_id") ?? stringField(input, "commandId"),
    format: requestedFormat ?? record.format,
    operation: resolved.operation,
    arguments: resolved.args,
    target: {
      sessionId: record.sessionId,
      documentId: record.id,
      revision,
      ...(target.anchor !== undefined ? { anchor: target.anchor } : {}),
    },
    source: commandSourceFrom(input),
    policy: {
      ...commandPolicyFrom(input),
      ...(resolved.action?.requiresReview !== undefined
        ? { requiresReview: resolved.action.requiresReview }
        : {}),
      ...(mode ? { mode } : {}),
    },
  });

  return { envelope, record, ...(resolved.action ? { action: resolved.action } : {}) };
}

function commandEnvelopeFromRaw(raw: Record<string, unknown>): CommandEnvelope {
  const targetRaw = isRecord(raw.target) ? raw.target : {};
  const sourceRaw = isRecord(raw.source) ? raw.source : {};
  const policyRaw = isRecord(raw.policy) ? raw.policy : {};
  const documentId = stringField(targetRaw, "documentId") ?? stringField(targetRaw, "document_id");
  const sessionId = stringField(targetRaw, "sessionId") ?? stringField(targetRaw, "session_id");
  const revision = targetRaw.revision;
  const operation = stringField(raw, "operation");
  const format = stringField(raw, "format") as OfficeFormat | undefined;
  const id = stringField(raw, "id");
  if (!id) throw new Error("command.id is required.");
  if (!format) throw new Error("command.format is required.");
  if (!operation) throw new Error("command.operation is required.");
  if (!documentId) throw new Error("command.target.documentId is required.");
  if (!sessionId) throw new Error("command.target.sessionId is required.");
  if (typeof revision !== "number" || !Number.isInteger(revision)) {
    throw new Error("command.target.revision must be an integer.");
  }
  const mode = policyRaw.mode;
  const surface = sourceRaw.surface;
  return {
    id,
    format,
    operation,
    arguments: raw.arguments ?? {},
    target: {
      sessionId,
      documentId,
      revision,
      ...(targetRaw.anchor !== undefined ? { anchor: targetRaw.anchor } : {}),
    },
    source: {
      surface:
        surface === "web" || surface === "cli" || surface === "internal" || surface === "mcp"
          ? surface
          : "mcp",
      ...(typeof sourceRaw.actorId === "string" ? { actorId: sourceRaw.actorId } : {}),
    },
    policy: {
      mode: mode === "dry_run" || mode === "auto_apply" || mode === "pending" ? mode : "pending",
      requiresReview:
        typeof policyRaw.requiresReview === "boolean"
          ? policyRaw.requiresReview
          : typeof policyRaw.requires_review === "boolean"
            ? policyRaw.requires_review
            : true,
    },
    createdAt: typeof raw.createdAt === "number" ? raw.createdAt : Date.now(),
  };
}

function resolveCommandReference(input: Record<string, unknown>): {
  readonly envelope: CommandEnvelope;
  readonly record: DocumentRecord;
} {
  const commandId = stringField(input, "command_id") ?? stringField(input, "commandId");
  if (commandId) {
    const envelope = plannedCommands.get(commandId);
    if (!envelope) throw new Error(`Unknown command_id "${commandId}". Call plan_command first.`);
    return { envelope, record: lookupDocument(envelope.target.documentId) };
  }

  if (isRecord(input.command)) {
    const envelope = commandEnvelopeFromRaw(input.command);
    return { envelope, record: lookupDocument(envelope.target.documentId) };
  }

  return commandEnvelopeFromInput(input);
}

function commandDiagnostics(record: DocumentRecord, envelope: CommandEnvelope): CommandDiagnostic[] {
  const diagnostics: CommandDiagnostic[] = [
    ...validateCommandEnvelope(envelope, snapshotFor(record)).diagnostics,
    ...anchorDiagnostics(record, envelope.target.anchor),
  ];
  if (!handlerFor(envelope.format as OfficeFormat, envelope.operation)) {
    diagnostics.push({
      level: "error",
      code: "unsupported-operation",
      message: `${envelope.operation} is not registered for ${envelope.format}.`,
    });
  }
  if (diagnostics.length === 0) {
    diagnostics.push({
      level: "info",
      code: "command-valid",
      message: `${envelope.operation} is valid for ${record.format} revision ${revisionFor(record)}.`,
    });
  }
  return diagnostics;
}

function anchorDiagnostics(record: DocumentRecord, anchor: unknown): CommandDiagnostic[] {
  if (anchor === undefined) {
    return [
      {
        level: "info",
        code: "anchor-default",
        message: "No anchor supplied; command payload owns targeting.",
      },
    ];
  }
  if (!isRecord(anchor)) {
    return [{ level: "error", code: "invalid-anchor", message: "target.anchor must be an object." }];
  }
  const kind = stringField(anchor, "kind");
  switch (record.format) {
    case "docx":
      return validateDocxAnchor(record, kind, anchor);
    case "xlsx":
      return validateXlsxAnchor(record, kind, anchor);
    case "pptx":
      return validatePptxAnchor(record, kind, anchor);
    case "pdf":
      return validatePdfAnchor(record, kind, anchor);
  }
}

function validateDocxAnchor(
  record: DocumentRecord,
  kind: string | undefined,
  anchor: Record<string, unknown>
): CommandDiagnostic[] {
  if (kind !== "paragraph") {
    return [
      { level: "error", code: "invalid-anchor-kind", message: "DOCX anchors must use kind='paragraph'." },
    ];
  }
  const index = anchor.index;
  if (typeof index !== "number" || !Number.isInteger(index) || index < 0) {
    return [
      { level: "error", code: "invalid-anchor-index", message: "DOCX paragraph anchor requires index >= 0." },
    ];
  }
  const body = (snapshotFor(record).root as { body?: unknown }).body;
  if (!Array.isArray(body) || !body[index] || (body[index] as { kind?: unknown }).kind !== "paragraph") {
    return [{ level: "error", code: "anchor-not-found", message: `DOCX paragraph ${index} was not found.` }];
  }
  return [{ level: "info", code: "anchor-resolved", message: `Resolved DOCX paragraph ${index}.` }];
}

function validateXlsxAnchor(
  record: DocumentRecord,
  kind: string | undefined,
  anchor: Record<string, unknown>
): CommandDiagnostic[] {
  if (kind !== "range") {
    return [{ level: "error", code: "invalid-anchor-kind", message: "XLSX anchors must use kind='range'." }];
  }
  const sheet = stringField(anchor, "sheet");
  const range = stringField(anchor, "range");
  if (!sheet || !range) {
    return [
      {
        level: "error",
        code: "invalid-anchor-range",
        message: "XLSX range anchor requires sheet and range.",
      },
    ];
  }
  const sheets = (snapshotFor(record).root as { sheets?: ReadonlyArray<{ name?: string }> }).sheets ?? [];
  if (!sheets.some((s) => s.name === sheet)) {
    return [{ level: "error", code: "anchor-not-found", message: `XLSX sheet "${sheet}" was not found.` }];
  }
  return [{ level: "info", code: "anchor-resolved", message: `Resolved XLSX range ${sheet}!${range}.` }];
}

function validatePptxAnchor(
  record: DocumentRecord,
  kind: string | undefined,
  anchor: Record<string, unknown>
): CommandDiagnostic[] {
  if (kind !== "slide_shape" && kind !== "slide-shape") {
    return [
      { level: "error", code: "invalid-anchor-kind", message: "PPTX anchors must use kind='slide_shape'." },
    ];
  }
  const slideIndex = anchor.slideIndex ?? anchor.slide_index;
  if (typeof slideIndex !== "number" || !Number.isInteger(slideIndex) || slideIndex < 0) {
    return [
      { level: "error", code: "invalid-anchor-slide", message: "PPTX anchor requires slideIndex >= 0." },
    ];
  }
  const slides = (snapshotFor(record).root as { slides?: ReadonlyArray<{ shapes?: unknown }> }).slides ?? [];
  const slide = slides[slideIndex];
  if (!slide) {
    return [{ level: "error", code: "anchor-not-found", message: `PPTX slide ${slideIndex} was not found.` }];
  }
  const shapeId = stringField(anchor, "shapeId") ?? stringField(anchor, "shape_id");
  if (shapeId && !shapeExists((slide as { shapes?: unknown }).shapes, shapeId)) {
    return [
      {
        level: "error",
        code: "anchor-not-found",
        message: `PPTX shape "${shapeId}" was not found on slide ${slideIndex}.`,
      },
    ];
  }
  return [
    {
      level: "info",
      code: "anchor-resolved",
      message: shapeId
        ? `Resolved PPTX slide ${slideIndex} shape ${shapeId}.`
        : `Resolved PPTX slide ${slideIndex}.`,
    },
  ];
}

function validatePdfAnchor(
  record: DocumentRecord,
  kind: string | undefined,
  anchor: Record<string, unknown>
): CommandDiagnostic[] {
  if (kind !== "page_region" && kind !== "page-region") {
    return [
      { level: "error", code: "invalid-anchor-kind", message: "PDF anchors must use kind='page_region'." },
    ];
  }
  const page = anchor.page;
  const rect = anchor.rect;
  if (typeof page !== "number" || !Number.isInteger(page) || page < 1) {
    return [
      { level: "error", code: "invalid-anchor-page", message: "PDF page-region anchor requires page >= 1." },
    ];
  }
  if (!isRecord(rect) || !["x", "y", "width", "height"].every((key) => typeof rect[key] === "number")) {
    return [
      {
        level: "error",
        code: "invalid-anchor-rect",
        message: "PDF page-region anchor requires rect { x, y, width, height }.",
      },
    ];
  }
  const pages = (snapshotFor(record).root as { pages?: ReadonlyArray<unknown> }).pages ?? [];
  if (page > pages.length) {
    return [{ level: "error", code: "anchor-not-found", message: `PDF page ${page} was not found.` }];
  }
  return [{ level: "info", code: "anchor-resolved", message: `Resolved PDF page ${page} region.` }];
}

function shapeExists(shapes: unknown, shapeId: string): boolean {
  if (!Array.isArray(shapes)) return false;
  for (const shape of shapes) {
    if (!isRecord(shape)) continue;
    if (shape.id === shapeId) return true;
    if (shapeExists(shape.children, shapeId)) return true;
  }
  return false;
}

function hasBlockingCommandDiagnostics(diagnostics: ReadonlyArray<CommandDiagnostic>): boolean {
  return diagnostics.some((d) => d.level === "error" || d.level === "destructive");
}

function commandSourceForEnvelope(envelope: CommandEnvelope): CommandSource {
  if (envelope.policy.mode === "pending") return "agent";
  if (envelope.source.surface === "internal") return "system";
  if (envelope.source.surface === "mcp") return "agent";
  return "human";
}

async function dispatchCommand(record: DocumentRecord, envelope: CommandEnvelope): Promise<AnyMutation> {
  const command = {
    type: envelope.operation,
    payload: envelope.arguments,
    source: commandSourceForEnvelope(envelope),
    ...(envelope.source.actorId ? { agentId: envelope.source.actorId } : {}),
  };
  switch (record.format) {
    case "docx":
      return (await lookupAgent(record.id).applyCommand(command)) as unknown as AnyMutation;
    case "xlsx":
      return (await lookupXlsxAgent(record.id).applyCommand(command)) as unknown as AnyMutation;
    case "pptx":
      return (await lookupPptxAgent(record.id).applyCommand(command)) as unknown as AnyMutation;
    case "pdf":
      return (await lookupPdfSession(record.id).agent.applyCommand(command)) as unknown as AnyMutation;
  }
}

function pendingMutationsFor(record: DocumentRecord): ReadonlyArray<AnyMutation> {
  return agentFor(record).getPendingMutations() as ReadonlyArray<AnyMutation>;
}

function approveMutationFor(record: DocumentRecord, mutationId: string): void {
  agentFor(record).approveMutation(mutationId);
  touchDocument(record);
}

function rejectMutationFor(record: DocumentRecord, mutationId: string): void {
  agentFor(record).rejectMutation(mutationId);
  touchDocument(record);
}

function undoFor(record: DocumentRecord): AnyMutation | null {
  const mutation = agentFor(record).undo() as AnyMutation | null;
  if (mutation) touchDocument(record);
  return mutation;
}

function touchDocument(record: DocumentRecord, diagnostics?: ReadonlyArray<McpDiagnostic>): void {
  record.updatedAt = nowIso();
  if (diagnostics) record.diagnostics = [...diagnostics];
  touchSession(getOrCreateSession(record.sessionId));
}

function mutationSummary(mutation: AnyMutation): Record<string, unknown> {
  return {
    id: mutation.id,
    operation: mutation.command.type,
    source: mutation.command.source,
    ...(mutation.command.agentId ? { actorId: mutation.command.agentId } : {}),
    status: mutation.status,
    timestamp: mutation.command.timestamp,
    ...(mutation.rejection ? { rejection: mutation.rejection } : {}),
  };
}

function commandEnvelopePayload(envelope: CommandEnvelope): Record<string, unknown> {
  return {
    schema: "office-ai/command@1",
    ...envelope,
  };
}

function summaryFor(record: DocumentRecord): unknown {
  switch (record.format) {
    case "docx":
      return inspectSnapshot(lookupAgent(record.id).getSnapshot());
    case "xlsx":
      return inspectXlsxSnapshot(lookupXlsxAgent(record.id).getSnapshot());
    case "pptx":
      return pptxInspectSnapshot(lookupPptxAgent(record.id).getSnapshot());
    case "pdf":
      return projectMetadata(lookupPdfSession(record.id).agent.getSnapshot());
  }
}

function documentEnvelope(record: DocumentRecord): Record<string, unknown> {
  return {
    schema: "office-ai/document@1",
    documentId: record.id,
    sessionId: record.sessionId,
    format: record.format,
    name: record.name,
    status: record.status,
    revision: revisionFor(record),
    ...(record.sourcePath ? { sourcePath: record.sourcePath } : {}),
    diagnostics: record.diagnostics,
    exportHistory: record.exportHistory,
    createdAt: record.createdAt,
    updatedAt: record.updatedAt,
  };
}

function docxPlainText(agent: DocxAgent): string {
  const lines: string[] = [];
  for (const b of agent.getSnapshot().root.body) {
    if (b.kind === "paragraph") {
      lines.push(
        b.children
          .map((c) =>
            c.kind === "run" ? c.children.map((g) => (g.kind === "text" ? g.text : "")).join("") : ""
          )
          .join("")
      );
    }
  }
  return lines.join("\n");
}

function projectionFor(
  record: DocumentRecord,
  opts: {
    readonly projection: "summary" | "markdown" | "json" | "text" | "page";
    readonly page?: number;
    readonly sheet?: string;
    readonly range?: string;
    readonly slide?: number;
    readonly maxRows?: number;
    readonly maxCols?: number;
  }
): Record<string, unknown> {
  const base = documentEnvelope(record);
  if (opts.projection === "summary") {
    return { ...base, projection: "summary", summary: summaryFor(record) };
  }

  switch (record.format) {
    case "docx": {
      const agent = lookupAgent(record.id);
      if (opts.projection === "json") {
        return { ...base, projection: "json", content: snapshotToJsonProjection(agent.getSnapshot()) };
      }
      if (opts.projection === "text") {
        return { ...base, projection: "text", content: docxPlainText(agent) };
      }
      if (opts.projection === "page") {
        const page = opts.page ?? 1;
        return { ...base, projection: "page", page, content: agent.getPageMarkdown(page) };
      }
      return { ...base, projection: "markdown", content: agent.toMarkdown() };
    }
    case "xlsx": {
      const agent = lookupXlsxAgent(record.id);
      if (opts.projection === "json") {
        return { ...base, projection: "json", content: xlsxRangeToJson(agent, opts.sheet, opts.range) };
      }
      return {
        ...base,
        projection: "markdown",
        content: agent.toMarkdown({
          ...(opts.sheet ? { sheet: opts.sheet } : {}),
          ...(opts.maxRows !== undefined ? { maxRows: opts.maxRows } : {}),
          ...(opts.maxCols !== undefined ? { maxCols: opts.maxCols } : {}),
        }),
      };
    }
    case "pptx": {
      const agent = lookupPptxAgent(record.id);
      const snap = agent.getSnapshot();
      const range =
        opts.slide !== undefined ? { startSlide: opts.slide, endSlide: opts.slide + 1 } : undefined;
      if (opts.projection === "json" || opts.projection === "page") {
        return { ...base, projection: opts.projection, content: pptxSnapshotToJsonProjection(snap, range) };
      }
      return {
        ...base,
        projection: "markdown",
        content: range ? pptxSnapshotToJsonProjection(snap, range) : agent.toMarkdown(),
      };
    }
    case "pdf": {
      const session = lookupPdfSession(record.id);
      if (opts.projection === "json") {
        return { ...base, projection: "json", content: projectMetadata(session.agent.getSnapshot()) };
      }
      if (opts.projection === "page") {
        const page = opts.page ?? 1;
        return { ...base, projection: "page", page, content: projectPage(session.agent.getSnapshot(), page) };
      }
      return { ...base, projection: "markdown", content: session.agent.toMarkdown() };
    }
  }
}

async function exportDocument(record: DocumentRecord, outPath?: string): Promise<ExportRecord> {
  const target = outPath ? resolve(outPath) : record.sourcePath;
  if (!target) {
    throw new Error("export_document requires out_path for documents created without a source path.");
  }

  let bytes: Uint8Array | Buffer;
  switch (record.format) {
    case "docx": {
      const agent = lookupAgent(record.id);
      agent.getPendingMutations().forEach((m) => agent.approveMutation(m.id));
      bytes = Buffer.from(await agent.exportFile());
      break;
    }
    case "xlsx": {
      const agent = lookupXlsxAgent(record.id);
      agent.getPendingMutations().forEach((m) => agent.approveMutation(m.id));
      bytes = Buffer.from(await agent.exportFile());
      break;
    }
    case "pptx": {
      const agent = lookupPptxAgent(record.id);
      agent.getPendingMutations().forEach((m) => agent.approveMutation(m.id));
      bytes = Buffer.from(await agent.exportFile());
      break;
    }
    case "pdf":
      bytes = await lookupPdfSession(record.id).agent.exportFile();
      break;
  }

  await writeFile(target, bytes);
  const exported: ExportRecord = { path: target, bytes: bytes.byteLength, exportedAt: nowIso() };
  record.exportHistory.push(exported);
  record.updatedAt = exported.exportedAt;
  touchSession(getOrCreateSession(record.sessionId));
  return exported;
}

function registerCommandLifecycleTools(server: McpServer): void {
  const commandInputSchema = {
    document_id: z.string().describe("documentId returned by import_document or create_document."),
    format: z.enum(["docx", "xlsx", "pptx", "pdf"]).optional().describe("Optional format guard."),
    action_id: z
      .string()
      .optional()
      .describe("Optional action catalogue id such as xlsx.set-cell; maps to operation/payload."),
    operation: z.string().optional().describe("Command operation such as docx:insert-text."),
    arguments: z.record(z.string(), z.unknown()).optional().describe("Command payload."),
    target: z
      .object({
        revision: z.number().int().nonnegative().optional(),
        anchor: z.unknown().optional(),
      })
      .optional()
      .describe("Optional revision and format-specific anchor."),
    policy: z
      .object({
        mode: z.enum(["dry_run", "auto_apply", "pending"]).optional(),
        requires_review: z.boolean().optional(),
      })
      .optional(),
    actor_id: z.string().optional().describe("Optional actor id recorded on the command source."),
  };

  const commandReferenceSchema = {
    command_id: z.string().optional().describe("Command id returned by plan_command."),
    command: z.record(z.string(), z.unknown()).optional().describe("Full office-ai/command@1 envelope."),
    ...Object.fromEntries(
      Object.entries(commandInputSchema).map(([key, value]) => [
        key,
        key === "document_id" ? value.optional() : value,
      ])
    ),
  };

  server.registerTool(
    "plan_command",
    {
      description:
        "Create and validate a structured office-ai/command@1 envelope for a DOCX/XLSX/PPTX/PDF document. No mutation is applied.",
      inputSchema: commandInputSchema,
    },
    async (input) => {
      try {
        const { envelope, record, action } = commandEnvelopeFromInput(input);
        const diagnostics = commandDiagnostics(record, envelope);
        plannedCommands.set(envelope.id, envelope);
        touchDocument(record, diagnostics);
        return ok({
          schema: "office-ai/command-plan@1",
          ok: !hasBlockingCommandDiagnostics(diagnostics),
          commandId: envelope.id,
          command: commandEnvelopePayload(envelope),
          document: documentEnvelope(record),
          ...(action ? { action } : {}),
          diagnostics,
          nextActions: ["preview_command", "apply_command"],
        });
      } catch (err) {
        return fail(`plan_command failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  server.registerTool(
    "preview_command",
    {
      description:
        "Preview a planned or inline office-ai/command@1 envelope. Returns diagnostics and a structured diff without applying the mutation.",
      inputSchema: commandReferenceSchema,
    },
    async (input) => {
      try {
        const { envelope, record } = resolveCommandReference(input);
        const diagnostics = commandDiagnostics(record, envelope);
        if (hasBlockingCommandDiagnostics(diagnostics)) {
          touchDocument(record, diagnostics);
          return ok({
            schema: "office-ai/command-preview@1",
            ok: false,
            stage: "failed",
            commandId: envelope.id,
            command: commandEnvelopePayload(envelope),
            document: documentEnvelope(record),
            diagnostics,
          });
        }
        const handler = handlerFor(record.format, envelope.operation);
        if (!handler) throw new Error(`${envelope.operation} is not registered for ${record.format}.`);
        const preview = previewCommandEnvelope(envelope, snapshotFor(record), handler);
        const mergedDiagnostics =
          preview.diagnostics.length > 0
            ? preview.diagnostics
            : [
                {
                  level: "info" as const,
                  code: "preview-ready",
                  message: `${envelope.operation} preview produced ${preview.diff?.changes.length ?? 0} change(s).`,
                },
              ];
        touchDocument(record, mergedDiagnostics);
        return ok({
          schema: "office-ai/command-preview@1",
          ok: preview.ok,
          stage: preview.stage,
          commandId: envelope.id,
          command: commandEnvelopePayload(envelope),
          document: documentEnvelope(record),
          diagnostics: mergedDiagnostics,
          ...(preview.diff ? { diff: preview.diff } : {}),
        });
      } catch (err) {
        return fail(`preview_command failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  server.registerTool(
    "apply_command",
    {
      description:
        "Apply a planned or inline office-ai/command@1 envelope. policy.mode=auto_apply approves immediately; pending leaves the mutation for review; dry_run is rejected.",
      inputSchema: commandReferenceSchema,
    },
    async (input) => {
      try {
        const { envelope, record } = resolveCommandReference(input);
        const diagnostics = commandDiagnostics(record, envelope);
        if (envelope.policy.mode === "dry_run") {
          diagnostics.push({
            level: "error",
            code: "dry-run-apply",
            message: "Dry-run commands must be previewed; apply requires policy.mode auto_apply or pending.",
          });
        }
        if (hasBlockingCommandDiagnostics(diagnostics)) {
          touchDocument(record, diagnostics);
          return ok({
            schema: "office-ai/command-apply@1",
            ok: false,
            stage: "failed",
            commandId: envelope.id,
            command: commandEnvelopePayload(envelope),
            document: documentEnvelope(record),
            diagnostics,
          });
        }

        const mutation = await dispatchCommand(record, envelope);
        if (envelope.policy.mode === "auto_apply" && mutation.status === "pending") {
          approveMutationFor(record, mutation.id);
        }
        const finalDiagnostics: CommandDiagnostic[] =
          mutation.status === "rejected"
            ? [
                {
                  level: "error",
                  code: mutation.rejection?.code ?? "command-rejected",
                  message: mutation.rejection?.message ?? `${envelope.operation} was rejected.`,
                },
              ]
            : [
                {
                  level: "info",
                  code: mutation.status === "pending" ? "command-pending" : "command-applied",
                  message:
                    mutation.status === "pending"
                      ? `${envelope.operation} is pending review.`
                      : `${envelope.operation} was applied.`,
                },
              ];
        touchDocument(record, finalDiagnostics);
        return ok({
          schema: "office-ai/command-apply@1",
          ok: mutation.status !== "rejected",
          stage:
            mutation.status === "pending" ? "queued" : mutation.status === "rejected" ? "failed" : "applied",
          commandId: envelope.id,
          command: commandEnvelopePayload(envelope),
          document: documentEnvelope(record),
          diagnostics: finalDiagnostics,
          mutation: mutationSummary(mutation),
          diff: mutation.diff,
          nextActions:
            mutation.status === "pending" ? ["list_pending_changes", "approve_change"] : ["export_document"],
        });
      } catch (err) {
        return fail(`apply_command failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  server.registerTool(
    "undo_command",
    {
      description: "Undo the most recent approved mutation for a canonical OfficeAI document.",
      inputSchema: {
        document_id: z.string().describe("documentId returned by import_document or create_document."),
      },
    },
    async ({ document_id }) => {
      try {
        const record = lookupDocument(document_id);
        const mutation = undoFor(record);
        const diagnostics: CommandDiagnostic[] = mutation
          ? [{ level: "info", code: "undo-applied", message: `Undid mutation ${mutation.id}.` }]
          : [{ level: "info", code: "undo-empty", message: "No approved mutation is available to undo." }];
        touchDocument(record, diagnostics);
        return ok({
          schema: "office-ai/command-undo@1",
          ok: true,
          didUndo: Boolean(mutation),
          document: documentEnvelope(record),
          diagnostics,
          ...(mutation ? { mutation: mutationSummary(mutation), diff: mutation.diff } : {}),
        });
      } catch (err) {
        return fail(`undo_command failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  server.registerTool(
    "list_pending_changes",
    {
      description: "List pending review mutations for one canonical document or for all canonical documents.",
      inputSchema: {
        document_id: z.string().optional().describe("Optional documentId; omit to list all pending changes."),
      },
    },
    async ({ document_id }) => {
      try {
        const records = document_id ? [lookupDocument(document_id)] : [...mcpDocuments.values()];
        return ok({
          schema: "office-ai/pending-changes@1",
          pending: records.flatMap((record) =>
            pendingMutationsFor(record).map((mutation) => ({
              documentId: record.id,
              sessionId: record.sessionId,
              format: record.format,
              documentName: record.name,
              mutation: mutationSummary(mutation),
              diff: mutation.diff,
            }))
          ),
        });
      } catch (err) {
        return fail(`list_pending_changes failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  server.registerTool(
    "approve_change",
    {
      description: "Approve a pending mutation by id for a canonical OfficeAI document.",
      inputSchema: {
        document_id: z.string().describe("documentId returned by import_document or create_document."),
        mutation_id: z.string().describe("Mutation id returned by apply_command or list_pending_changes."),
      },
    },
    async ({ document_id, mutation_id }) => {
      try {
        const record = lookupDocument(document_id);
        approveMutationFor(record, mutation_id);
        const diagnostics: CommandDiagnostic[] = [
          { level: "info", code: "change-approved", message: `Approved mutation ${mutation_id}.` },
        ];
        touchDocument(record, diagnostics);
        return ok({
          schema: "office-ai/change-review@1",
          ok: true,
          approved: mutation_id,
          document: documentEnvelope(record),
          diagnostics,
        });
      } catch (err) {
        return fail(`approve_change failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  server.registerTool(
    "reject_change",
    {
      description: "Reject a pending mutation by id for a canonical OfficeAI document.",
      inputSchema: {
        document_id: z.string().describe("documentId returned by import_document or create_document."),
        mutation_id: z.string().describe("Mutation id returned by apply_command or list_pending_changes."),
        reason: z.string().optional().describe("Optional human-readable rejection reason."),
      },
    },
    async ({ document_id, mutation_id, reason }) => {
      try {
        const record = lookupDocument(document_id);
        rejectMutationFor(record, mutation_id);
        const diagnostics: CommandDiagnostic[] = [
          {
            level: "info",
            code: "change-rejected",
            message: reason
              ? `Rejected mutation ${mutation_id}: ${reason}`
              : `Rejected mutation ${mutation_id}.`,
          },
        ];
        touchDocument(record, diagnostics);
        return ok({
          schema: "office-ai/change-review@1",
          ok: true,
          rejected: mutation_id,
          ...(reason ? { reason } : {}),
          document: documentEnvelope(record),
          diagnostics,
        });
      } catch (err) {
        return fail(`reject_change failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );
}

function registerSessionDocumentTools(server: McpServer): void {
  server.registerTool(
    "create_session",
    {
      description:
        "Create an in-process OfficeAI session. Use the returned sessionId with import_document/create_document/list_documents.",
      inputSchema: {
        title: z.string().optional().describe("Optional human-readable session title."),
      },
    },
    async ({ title }) => {
      const session = createSessionRecord(title ?? "OfficeAI session");
      return ok({
        schema: "office-ai/session@1",
        sessionId: session.id,
        title: session.title,
        createdAt: session.createdAt,
        updatedAt: session.updatedAt,
        documentCount: session.documentIds.size,
        nextActions: ["import_document", "create_document", "list_documents"],
      });
    }
  );

  server.registerTool(
    "list_sessions",
    {
      description: "List in-process OfficeAI sessions and their document counts.",
      inputSchema: {},
    },
    async () =>
      ok({
        schema: "office-ai/session-list@1",
        sessions: [...mcpSessions.values()].map((session) => ({
          sessionId: session.id,
          title: session.title,
          createdAt: session.createdAt,
          updatedAt: session.updatedAt,
          documentCount: session.documentIds.size,
        })),
      })
  );

  server.registerTool(
    "import_document",
    {
      description:
        "Import a DOCX/XLSX/PPTX/PDF from disk into a canonical OfficeAI document session. The returned documentId is also a compatible handle for the matching legacy format tools.",
      inputSchema: {
        path: z
          .string()
          .describe("Absolute or workspace-relative path to a .docx, .xlsx, .pptx or .pdf file."),
        session_id: z.string().optional().describe("Optional sessionId from create_session."),
        format: z
          .enum(["docx", "xlsx", "pptx", "pdf"])
          .optional()
          .describe("Override extension-based format detection."),
        name: z.string().optional().describe("Optional display name. Defaults to the file basename."),
      },
    },
    async ({ path, session_id, format, name }) => {
      try {
        const abs = resolve(path);
        const detected = inferFormatFromPath(abs, format as OfficeFormat | undefined);
        const buf = await readFile(abs);
        const documentId = `doc_${randomUUID()}`;
        const displayName = name ?? basename(abs);

        switch (detected) {
          case "docx": {
            const agent = await DocxAgent.fromBuffer(buf);
            sessions.set(documentId, agent);
            sessionPaths.set(documentId, abs);
            break;
          }
          case "xlsx": {
            const agent = await XlsxAgent.fromBuffer(buf);
            xlsxSessions.set(documentId, agent);
            xlsxSessionPaths.set(documentId, abs);
            break;
          }
          case "pptx": {
            const agent = await PptxAgent.fromBuffer(buf);
            pptxSessions.set(documentId, agent);
            pptxSessionPaths.set(documentId, abs);
            break;
          }
          case "pdf": {
            const bytes = new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
            const agent = await PdfAgent.fromBuffer(bytes);
            pdfSessions.set(documentId, { agent, bytes });
            pdfSessionPaths.set(documentId, abs);
            break;
          }
        }

        const record = registerDocumentRecord({
          id: documentId,
          sessionId: session_id,
          format: detected,
          name: displayName,
          sourcePath: abs,
          diagnostics: [
            { level: "info", code: "imported", message: `Imported ${displayName} as ${detected}.` },
          ],
        });
        return ok({
          schema: "office-ai/import-document@1",
          document: documentEnvelope(record),
          summary: summaryFor(record),
          nextActions: ["get_document_projection", "export_document"],
        });
      } catch (err) {
        return fail(`import_document failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  server.registerTool(
    "create_document",
    {
      description:
        "Create a blank DOCX/XLSX/PPTX/PDF document in a canonical OfficeAI session. The returned documentId is also a compatible handle for the matching legacy format tools.",
      inputSchema: {
        format: z.enum(["docx", "xlsx", "pptx", "pdf"]),
        session_id: z.string().optional().describe("Optional sessionId from create_session."),
        name: z.string().optional().describe("Optional display name. Defaults to untitled.<format>."),
      },
    },
    async ({ format, session_id, name }) => {
      try {
        const documentId = `doc_${randomUUID()}`;
        const detected = format as OfficeFormat;
        const displayName = name ?? `untitled.${detected}`;

        switch (detected) {
          case "docx":
            sessions.set(documentId, await DocxAgent.empty());
            break;
          case "xlsx":
            xlsxSessions.set(documentId, await XlsxAgent.empty());
            break;
          case "pptx":
            pptxSessions.set(documentId, await PptxAgent.empty());
            break;
          case "pdf": {
            const agent = await PdfAgent.empty();
            const bytes = await agent.exportFile();
            pdfSessions.set(documentId, { agent, bytes });
            break;
          }
        }

        const record = registerDocumentRecord({
          id: documentId,
          sessionId: session_id,
          format: detected,
          name: displayName,
          diagnostics: [{ level: "info", code: "created", message: `Created blank ${detected} document.` }],
        });
        return ok({
          schema: "office-ai/create-document@1",
          document: documentEnvelope(record),
          summary: summaryFor(record),
          nextActions: ["get_document_projection", "export_document"],
        });
      } catch (err) {
        return fail(`create_document failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  server.registerTool(
    "list_documents",
    {
      description: "List canonical OfficeAI documents, optionally restricted to one session.",
      inputSchema: {
        session_id: z.string().optional().describe("Optional sessionId from create_session."),
      },
    },
    async ({ session_id }) => {
      try {
        const records = session_id
          ? (() => {
              const session = mcpSessions.get(session_id);
              if (!session) throw new Error(`Unknown session_id "${session_id}".`);
              return [...session.documentIds].map((id) => lookupDocument(id));
            })()
          : [...mcpDocuments.values()];
        return ok({
          schema: "office-ai/document-list@1",
          documents: records.map((record) => documentEnvelope(record)),
        });
      } catch (err) {
        return fail(`list_documents failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  server.registerTool(
    "get_document",
    {
      description: "Return canonical document metadata, diagnostics, export history and a format summary.",
      inputSchema: {
        document_id: z.string().describe("documentId returned by import_document or create_document."),
      },
    },
    async ({ document_id }) => {
      try {
        const record = lookupDocument(document_id);
        return ok({
          schema: "office-ai/get-document@1",
          document: documentEnvelope(record),
          summary: summaryFor(record),
        });
      } catch (err) {
        return fail(`get_document failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  server.registerTool(
    "get_document_projection",
    {
      description:
        "Read a canonical document projection without moving binary bytes. Supports summary/markdown/json/text/page, with format-specific windowing fields.",
      inputSchema: {
        document_id: z.string().describe("documentId returned by import_document or create_document."),
        projection: z.enum(["summary", "markdown", "json", "text", "page"]).optional().default("markdown"),
        page: z.number().int().positive().optional().describe("1-based DOCX/PDF page for projection='page'."),
        sheet: z.string().optional().describe("XLSX sheet name for markdown/json projections."),
        range: z.string().optional().describe("XLSX A1 range for json projections; requires sheet."),
        slide: z.number().int().min(0).optional().describe("0-based PPTX slide for json/page projections."),
        max_rows: z.number().int().positive().optional().describe("XLSX markdown row limit."),
        max_cols: z.number().int().positive().optional().describe("XLSX markdown column limit."),
      },
    },
    async ({ document_id, projection, page, sheet, range, slide, max_rows, max_cols }) => {
      try {
        return ok(
          projectionFor(lookupDocument(document_id), {
            projection: projection ?? "markdown",
            ...(page !== undefined ? { page } : {}),
            ...(sheet !== undefined ? { sheet } : {}),
            ...(range !== undefined ? { range } : {}),
            ...(slide !== undefined ? { slide } : {}),
            ...(max_rows !== undefined ? { maxRows: max_rows } : {}),
            ...(max_cols !== undefined ? { maxCols: max_cols } : {}),
          })
        );
      } catch (err) {
        return fail(`get_document_projection failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  server.registerTool(
    "export_document",
    {
      description:
        "Export a canonical OfficeAI document as a real DOCX/XLSX/PPTX/PDF file. Defaults to the original import path when available; pass out_path for created documents or copy-style exports.",
      inputSchema: {
        document_id: z.string().describe("documentId returned by import_document or create_document."),
        out_path: z
          .string()
          .optional()
          .describe("Optional output path. Required for documents created without a source path."),
      },
    },
    async ({ document_id, out_path }) => {
      try {
        const record = lookupDocument(document_id);
        const exported = await exportDocument(record, out_path);
        return ok({
          schema: "office-ai/export-document@1",
          document: documentEnvelope(record),
          exported,
          nextActions: ["get_document", "get_document_projection"],
        });
      } catch (err) {
        return fail(`export_document failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );
}

/**
 * Build a fresh MCP server with all OfficeAI tools registered. Exposed for
 * tests so they can wire it to an in-memory transport pair without touching
 * stdio.
 */
export function createMcpServer(): McpServer {
  const server = new McpServer(
    { name: "officeai", version: "0.1.0" },
    {
      capabilities: { tools: {} },
      instructions:
        "MCP-first document tools for DOCX, XLSX, PPTX and PDF. Prefer create_session/import_document/create_document/get_document_projection plus plan_command/preview_command/apply_command/export_document for canonical cross-format flows. Legacy docx_*/xlsx_*/pptx_*/pdf_* tools remain available; canonical documentId values also work as the matching legacy handles.",
    }
  );

  registerSessionDocumentTools(server);
  registerCommandLifecycleTools(server);

  // ── docx_load ─────────────────────────────────────────────────────────
  server.registerTool(
    "docx_load",
    {
      description: "Load a .docx file from disk. Returns an opaque `handle` to use with subsequent tools.",
      inputSchema: {
        path: z.string().describe("Absolute or workspace-relative path to a .docx file."),
      },
    },
    async ({ path }) => {
      try {
        const buf = await readFile(resolve(path));
        const agent = await DocxAgent.fromBuffer(buf);
        const handle = randomUUID();
        sessions.set(handle, agent);
        sessionPaths.set(handle, resolve(path));
        registerDocumentRecord({
          id: handle,
          format: "docx",
          name: basename(resolve(path)),
          sourcePath: resolve(path),
        });
        return ok({ handle, path: resolve(path), summary: inspectSnapshot(agent.getSnapshot()) });
      } catch (err) {
        return fail(`docx_load failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── docx_save ─────────────────────────────────────────────────────────
  server.registerTool(
    "docx_save",
    {
      description:
        "Serialize the current snapshot back to disk. Defaults to the path passed to docx_load; pass `out_path` to write elsewhere.",
      inputSchema: {
        handle: z.string().describe("Handle returned by docx_load."),
        out_path: z
          .string()
          .optional()
          .describe("Optional output path. Defaults to the original path passed to docx_load."),
      },
    },
    async ({ handle, out_path }) => {
      try {
        const agent = lookupAgent(handle);
        agent.getPendingMutations().forEach((m) => agent.approveMutation(m.id));
        const target = out_path ? resolve(out_path) : sessionPaths.get(handle);
        if (!target) {
          return fail(`docx_save: no path known for handle "${handle}". Pass out_path explicitly.`);
        }
        const buf = Buffer.from(await agent.exportFile());
        await writeFile(target, buf);
        return ok({ wrote: target, bytes: buf.byteLength, revision: agent.getSnapshot().revision });
      } catch (err) {
        return fail(`docx_save failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── docx_inspect ──────────────────────────────────────────────────────
  server.registerTool(
    "docx_inspect",
    {
      description: "Return a structural summary (paragraphs, tables, comments, parts).",
      inputSchema: { handle: z.string() },
    },
    async ({ handle }) => {
      try {
        const agent = lookupAgent(handle);
        return ok(inspectSnapshot(agent.getSnapshot()));
      } catch (err) {
        return fail(`docx_inspect failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── docx_get_text ─────────────────────────────────────────────────────
  server.registerTool(
    "docx_get_text",
    {
      description:
        "Return the document content as Markdown (default), structured JSON, or plain text. Pass `with_page_sections: true` to interleave `<!-- page N -->` anchors so the LLM can cite pages.",
      inputSchema: {
        handle: z.string(),
        format: z.enum(["markdown", "json", "text"]).optional().default("markdown"),
        with_page_sections: z
          .boolean()
          .optional()
          .default(false)
          .describe(
            "When true and format is markdown, prepend each page with a <!-- page N --> anchor + ## Page N heading."
          ),
      },
    },
    async ({ handle, format, with_page_sections }) => {
      try {
        const agent = lookupAgent(handle);
        const fmt = format ?? "markdown";
        switch (fmt) {
          case "markdown":
            return ok({
              format: fmt,
              content: agent.toMarkdown(with_page_sections ? { withPageSections: true } : undefined),
            });
          case "json":
            return ok(snapshotToJsonProjection(agent.getSnapshot()));
          case "text": {
            const lines: string[] = [];
            for (const b of agent.getSnapshot().root.body) {
              if (b.kind === "paragraph") {
                lines.push(
                  b.children
                    .map((c) =>
                      c.kind === "run"
                        ? c.children.map((g) => (g.kind === "text" ? g.text : "")).join("")
                        : ""
                    )
                    .join("")
                );
              }
            }
            return ok({ format: fmt, content: lines.join("\n") });
          }
          default: {
            const _exhaustive: never = fmt;
            void _exhaustive;
            return fail(`unknown format: ${String(format)}`);
          }
        }
      } catch (err) {
        return fail(`docx_get_text failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── docx_get_pages ────────────────────────────────────────────────────
  server.registerTool(
    "docx_get_pages",
    {
      description:
        "List the document's logical pages with their body-block range, the trigger that started each page (doc-start, page-break, last-rendered, section-break), and a short text preview. Page numbers are 1-based and global across the document.",
      inputSchema: { handle: z.string() },
    },
    async ({ handle }) => {
      try {
        const agent = lookupAgent(handle);
        const pages = agent.getPages();
        return ok({ pages, total: pages.length });
      } catch (err) {
        return fail(`docx_get_pages failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── docx_get_page_text ────────────────────────────────────────────────
  server.registerTool(
    "docx_get_page_text",
    {
      description:
        "Return the markdown (default) or plain-text projection of a single page. Pass `page` (1-based) and optionally `format`. Use docx_get_pages first to discover the available page numbers.",
      inputSchema: {
        handle: z.string(),
        page: z.number().int().positive(),
        format: z.enum(["markdown", "text"]).optional().default("markdown"),
      },
    },
    async ({ handle, page, format }) => {
      try {
        const agent = lookupAgent(handle);
        const fmt = format ?? "markdown";
        const pages = agent.getPages();
        const info = pages.find((p) => p.pageNumber === page);
        if (!info) {
          return fail(
            `docx_get_page_text: page ${page} out-of-range (document has ${pages.length} page${pages.length === 1 ? "" : "s"})`
          );
        }
        const content = fmt === "markdown" ? agent.getPageMarkdown(page) : agent.getPageText(page);
        if (content === null) {
          return fail(`docx_get_page_text: page ${page} out-of-range`);
        }
        return ok({
          pageNumber: page,
          startBlockIndex: info.startBlockIndex,
          endBlockIndex: info.endBlockIndex,
          format: fmt,
          content,
        });
      } catch (err) {
        return fail(`docx_get_page_text failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── docx_search ───────────────────────────────────────────────────────
  server.registerTool(
    "docx_search",
    {
      description: "Search the document body for text. Optional case-sensitive and regex flags.",
      inputSchema: {
        handle: z.string(),
        query: z.string().min(1),
        case_sensitive: z.boolean().optional().default(false),
        regex: z.boolean().optional().default(false),
      },
    },
    async ({ handle, query, case_sensitive, regex }) => {
      try {
        const agent = lookupAgent(handle);
        const results = agent.search({
          query,
          caseSensitive: case_sensitive ?? false,
          regex: regex ?? false,
        });
        return ok({ matches: results });
      } catch (err) {
        return fail(`docx_search failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── docx_apply_command ────────────────────────────────────────────────
  server.registerTool(
    "docx_apply_command",
    {
      description:
        "Apply a docx command (e.g. `docx:insert-text`, `docx:resolve-comment`). Pass an arbitrary payload object — schemas live in `@officeai/docx/commands/payloads`.",
      inputSchema: {
        handle: z.string(),
        type: z.string().describe('Command type, e.g. "docx:insert-text"'),
        payload: z.record(z.string(), z.unknown()).describe("Command payload."),
        source: z.enum(["agent", "human", "system"]).optional().default("agent"),
        agent_id: z.string().optional().default("officeai-mcp"),
        auto_approve: z
          .boolean()
          .optional()
          .default(true)
          .describe(
            "When true (default) any pending mutation produced by an `agent` source is immediately approved. Set false to leave it pending in the bus for downstream review."
          ),
      },
    },
    async ({ handle, type, payload, source, agent_id, auto_approve }) => {
      try {
        const agent = lookupAgent(handle);
        const mutation = await agent.applyCommand({
          type,
          payload,
          source: source ?? "agent",
          ...((source ?? "agent") === "agent" ? { agentId: agent_id ?? "officeai-mcp" } : {}),
        });
        if ((auto_approve ?? true) && mutation.status === "pending") {
          agent.approveMutation(mutation.id);
        }
        return ok({
          mutation: {
            id: mutation.id,
            status: agent.getPendingMutations().some((m) => m.id === mutation.id)
              ? "pending"
              : mutation.status,
            ...(mutation.rejection ? { rejection: mutation.rejection } : {}),
          },
          revision: agent.getSnapshot().revision,
        });
      } catch (err) {
        return fail(`docx_apply_command failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── docx_list_pending ─────────────────────────────────────────────────
  server.registerTool(
    "docx_list_pending",
    {
      description:
        "List mutations that are still in `pending` state on the bus (typically agent-authored writes invoked with auto_approve=false). Each entry includes the mutation id, command type, source, and the agent id that submitted it.",
      inputSchema: { handle: z.string() },
    },
    async ({ handle }) => {
      try {
        const agent = lookupAgent(handle);
        const pending = agent.getPendingMutations().map((m) => ({
          id: m.id,
          command: { type: m.command.type, source: m.command.source },
          ...(m.command.source === "agent" && "agentId" in m.command ? { agentId: m.command.agentId } : {}),
          revision: m.after.revision,
          status: m.status,
        }));
        return ok({ pending });
      } catch (err) {
        return fail(`docx_list_pending failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── docx_approve ──────────────────────────────────────────────────────
  server.registerTool(
    "docx_approve",
    {
      description:
        "Approve a pending mutation. After approval the mutation is committed and shows up in the snapshot's history.",
      inputSchema: {
        handle: z.string(),
        mutation_id: z.string().describe("Mutation id from docx_list_pending."),
      },
    },
    async ({ handle, mutation_id }) => {
      try {
        const agent = lookupAgent(handle);
        agent.approveMutation(mutation_id);
        return ok({ approved: mutation_id, revision: agent.getSnapshot().revision });
      } catch (err) {
        return fail(`docx_approve failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── docx_reject ───────────────────────────────────────────────────────
  server.registerTool(
    "docx_reject",
    {
      description:
        "Reject a pending mutation with an optional human-readable reason. The snapshot is unaffected; the mutation is dropped from the pending queue.",
      inputSchema: {
        handle: z.string(),
        mutation_id: z.string().describe("Mutation id from docx_list_pending."),
        reason: z.string().optional(),
      },
    },
    async ({ handle, mutation_id, reason }) => {
      try {
        const agent = lookupAgent(handle);
        agent.rejectMutation(mutation_id);
        return ok({ rejected: mutation_id, ...(reason ? { reason } : {}) });
      } catch (err) {
        return fail(`docx_reject failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── docx_diff ─────────────────────────────────────────────────────────
  server.registerTool(
    "docx_diff",
    {
      description:
        "Diff two loaded handles, OR diff a handle against the file on disk it was loaded from. Pass either {before, after} (two handles) or {handle, against: 'disk'}.",
      inputSchema: {
        before: z.string().optional(),
        after: z.string().optional(),
        handle: z.string().optional(),
        against: z.enum(["disk"]).optional(),
      },
    },
    async ({ before, after, handle, against }) => {
      try {
        if (before && after) {
          const a = lookupAgent(before);
          const b = lookupAgent(after);
          return ok(diffSnapshots(a.getSnapshot(), b.getSnapshot()));
        }
        if (handle && against === "disk") {
          const agent = lookupAgent(handle);
          const path = sessionPaths.get(handle);
          if (!path) return fail(`docx_diff: no on-disk path for handle "${handle}".`);
          const buf = await readFile(path);
          const baseline = await DocxAgent.fromBuffer(buf);
          return ok(diffSnapshots(baseline.getSnapshot(), agent.getSnapshot()));
        }
        return fail("docx_diff: pass either {before, after} (two handles) or {handle, against: 'disk'}.");
      } catch (err) {
        return fail(`docx_diff failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  registerXlsxTools(server);
  registerPdfTools(server);
  registerCatalogueActionTools(server);

  // ── pptx_load ─────────────────────────────────────────────────────────
  server.registerTool(
    "pptx_load",
    {
      description:
        "Load a .pptx file from disk. Returns an opaque `handle` to use with subsequent pptx_* tools.",
      inputSchema: {
        path: z.string().describe("Absolute or workspace-relative path to a .pptx file."),
      },
    },
    async ({ path }) => {
      try {
        const buf = await readFile(resolve(path));
        const agent = await PptxAgent.fromBuffer(buf);
        const handle = randomUUID();
        pptxSessions.set(handle, agent);
        pptxSessionPaths.set(handle, resolve(path));
        registerDocumentRecord({
          id: handle,
          format: "pptx",
          name: basename(resolve(path)),
          sourcePath: resolve(path),
        });
        return ok({ handle, path: resolve(path), summary: pptxInspectSnapshot(agent.getSnapshot()) });
      } catch (err) {
        return fail(`pptx_load failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── pptx_save ─────────────────────────────────────────────────────────
  server.registerTool(
    "pptx_save",
    {
      description:
        "Serialize the current PPTX snapshot back to disk. Defaults to the path passed to pptx_load; pass `out_path` to write elsewhere.",
      inputSchema: {
        handle: z.string().describe("Handle returned by pptx_load."),
        out_path: z
          .string()
          .optional()
          .describe("Optional output path. Defaults to the original path passed to pptx_load."),
      },
    },
    async ({ handle, out_path }) => {
      try {
        const agent = lookupPptxAgent(handle);
        agent.getPendingMutations().forEach((m) => agent.approveMutation(m.id));
        const target = out_path ? resolve(out_path) : pptxSessionPaths.get(handle);
        if (!target) {
          return fail(`pptx_save: no path known for handle "${handle}". Pass out_path explicitly.`);
        }
        const buf = Buffer.from(await agent.exportFile());
        await writeFile(target, buf);
        return ok({ wrote: target, bytes: buf.byteLength, revision: agent.getSnapshot().revision });
      } catch (err) {
        return fail(`pptx_save failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── pptx_inspect ──────────────────────────────────────────────────────
  server.registerTool(
    "pptx_inspect",
    {
      description: "Return a structural summary (slide count, shape kinds, masters/layouts, parts).",
      inputSchema: { handle: z.string() },
    },
    async ({ handle }) => {
      try {
        const agent = lookupPptxAgent(handle);
        return ok(pptxInspectSnapshot(agent.getSnapshot()));
      } catch (err) {
        return fail(`pptx_inspect failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── pptx_get_text ─────────────────────────────────────────────────────
  server.registerTool(
    "pptx_get_text",
    {
      description:
        "Return the presentation as Markdown (default), structured JSON, or plain text. Optionally restrict to a single slide.",
      inputSchema: {
        handle: z.string(),
        format: z.enum(["markdown", "json", "text"]).optional().default("markdown"),
        slide: z.number().int().min(0).optional().describe("Restrict to a single 0-based slide index."),
      },
    },
    async ({ handle, format, slide }) => {
      try {
        const agent = lookupPptxAgent(handle);
        const snap = agent.getSnapshot();
        const range = slide !== undefined ? { startSlide: slide, endSlide: slide + 1 } : undefined;
        const fmt = format ?? "markdown";
        switch (fmt) {
          case "markdown": {
            // Markdown projection lives in @officeai/pptx; for slice, fall back to JSON.
            if (!range) return ok({ format: fmt, content: agent.toMarkdown() });
            const proj = pptxSnapshotToJsonProjection(snap, range);
            const lines = ["# Presentation"];
            for (const s of proj.slides) {
              lines.push(`## Slide ${s.index + 1} — \`${s.partPath}\` (slideId=${s.slideId})`);
              if (s.transition) {
                const speed = s.transition.speed ? ` (${s.transition.speed})` : "";
                lines.push(`- _transition_: **${s.transition.kind}**${speed}`);
              }
              if (s.animations && s.animations.length > 0) {
                lines.push(`- _animations_:`);
                for (const a of s.animations) {
                  const dur = a.durationMs !== undefined ? ` ${a.durationMs}ms` : "";
                  lines.push(
                    `  - \`${a.id}\` ${a.order + 1}. **${a.category}/${a.preset}**${dur} → cNvPr=${a.targetCNvPrId}`
                  );
                }
              }
              for (const sh of s.shapes) {
                if (sh.kind === "text" && sh.text) lines.push(`> ${sh.text.replaceAll("\n", " · ")}`);
                if (sh.kind === "table" && sh.table) {
                  for (const row of sh.table.cells) {
                    lines.push(
                      `| ${row.map((c) => (c.length > 0 ? c.replaceAll("\n", " · ") : "(empty)")).join(" | ")} |`
                    );
                  }
                }
                if (sh.kind === "chart" && sh.chart) {
                  lines.push(
                    `> chart (${sh.chart.chartType})${sh.chart.title ? ` — ${sh.chart.title}` : ""}`
                  );
                  if (sh.chart.categories.length > 0) {
                    lines.push(`> categories: ${sh.chart.categories.join(", ")}`);
                  }
                  for (const ser of sh.chart.series) {
                    lines.push(`> ${ser.name ? `${ser.name}: ` : ""}[${ser.values.join(", ")}]`);
                  }
                }
              }
            }
            return ok({ format: fmt, content: lines.join("\n") });
          }
          case "json":
            return ok(pptxSnapshotToJsonProjection(snap, range));
          case "text": {
            const proj = pptxSnapshotToJsonProjection(snap, range);
            const lines: string[] = [];
            for (const s of proj.slides) {
              for (const sh of s.shapes) {
                if (sh.kind === "text" && sh.text) lines.push(sh.text);
                if (sh.kind === "table" && sh.table) {
                  for (const row of sh.table.cells) {
                    for (const cell of row) {
                      if (cell.length > 0) lines.push(cell);
                    }
                  }
                }
                if (sh.kind === "chart" && sh.chart && sh.chart.title) {
                  lines.push(sh.chart.title);
                }
              }
            }
            return ok({ format: fmt, content: lines.join("\n") });
          }
          default: {
            const _exhaustive: never = fmt;
            void _exhaustive;
            return fail(`unknown format: ${String(format)}`);
          }
        }
      } catch (err) {
        return fail(`pptx_get_text failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── pptx_search ───────────────────────────────────────────────────────
  server.registerTool(
    "pptx_search",
    {
      description: "Search every slide's text content for a query. Optional case-sensitive and regex flags.",
      inputSchema: {
        handle: z.string(),
        query: z.string().min(1),
        case_sensitive: z.boolean().optional().default(false),
        regex: z.boolean().optional().default(false),
      },
    },
    async ({ handle, query, case_sensitive, regex }) => {
      try {
        const agent = lookupPptxAgent(handle);
        const matches = agent.search({
          query,
          caseSensitive: case_sensitive ?? false,
          regex: regex ?? false,
        });
        return ok({ matches });
      } catch (err) {
        return fail(`pptx_search failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── pptx_apply_command ────────────────────────────────────────────────
  server.registerTool(
    "pptx_apply_command",
    {
      description:
        "Apply a typed pptx command (e.g. `pptx:set-text`, `pptx:add-slide`, `pptx:insert-image`, `pptx:set-chart-title`, `pptx:set-slide-transition`, `pptx:add-shape-animation`). Pass an arbitrary payload object — schemas live in `@officeai/pptx/commands/payloads`.",
      inputSchema: {
        handle: z.string(),
        type: z.string().describe('Command type, e.g. "pptx:set-text"'),
        payload: z.record(z.string(), z.unknown()).describe("Command payload."),
        source: z.enum(["agent", "human", "system"]).optional().default("agent"),
        agent_id: z.string().optional().default("officeai-mcp"),
        auto_approve: z
          .boolean()
          .optional()
          .default(true)
          .describe(
            "When true (default) any pending mutation produced by an `agent` source is immediately approved. Set false to leave it pending in the bus for downstream review."
          ),
      },
    },
    async ({ handle, type, payload, source, agent_id, auto_approve }) => {
      try {
        const agent = lookupPptxAgent(handle);
        const mutation = await agent.applyCommand({
          type,
          payload,
          source: source ?? "agent",
          ...((source ?? "agent") === "agent" ? { agentId: agent_id ?? "officeai-mcp" } : {}),
        });
        if ((auto_approve ?? true) && mutation.status === "pending") {
          agent.approveMutation(mutation.id);
        }
        return ok({
          mutation: {
            id: mutation.id,
            status: agent.getPendingMutations().some((m) => m.id === mutation.id)
              ? "pending"
              : mutation.status,
            ...(mutation.rejection ? { rejection: mutation.rejection } : {}),
          },
          revision: agent.getSnapshot().revision,
        });
      } catch (err) {
        return fail(`pptx_apply_command failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── pptx_diff ─────────────────────────────────────────────────────────
  server.registerTool(
    "pptx_diff",
    {
      description:
        "Diff two loaded PPTX handles, OR diff a handle against the file on disk it was loaded from. Pass either {before, after} (two handles) or {handle, against: 'disk'}.",
      inputSchema: {
        before: z.string().optional(),
        after: z.string().optional(),
        handle: z.string().optional(),
        against: z.enum(["disk"]).optional(),
      },
    },
    async ({ before, after, handle, against }) => {
      try {
        if (before && after) {
          const a = lookupPptxAgent(before);
          const b = lookupPptxAgent(after);
          return ok(pptxDiffSnapshots(a.getSnapshot(), b.getSnapshot()));
        }
        if (handle && against === "disk") {
          const agent = lookupPptxAgent(handle);
          const path = pptxSessionPaths.get(handle);
          if (!path) return fail(`pptx_diff: no on-disk path for handle "${handle}".`);
          const buf = await readFile(path);
          const baseline = await PptxAgent.fromBuffer(buf);
          return ok(pptxDiffSnapshots(baseline.getSnapshot(), agent.getSnapshot()));
        }
        return fail("pptx_diff: pass either {before, after} (two handles) or {handle, against: 'disk'}.");
      } catch (err) {
        return fail(`pptx_diff failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  return server;
}

// ──────────────────────────────────────────────────────────────────────────
// xlsx_* tools
// ──────────────────────────────────────────────────────────────────────────

function registerXlsxTools(server: McpServer): void {
  // ── xlsx_load ─────────────────────────────────────────────────────────
  server.registerTool(
    "xlsx_load",
    {
      description: "Load a .xlsx file from disk. Returns an opaque `handle` to use with subsequent tools.",
      inputSchema: {
        path: z.string().describe("Absolute or workspace-relative path to a .xlsx file."),
      },
    },
    async ({ path }) => {
      try {
        const buf = await readFile(resolve(path));
        const agent = await XlsxAgent.fromBuffer(buf);
        const handle = randomUUID();
        xlsxSessions.set(handle, agent);
        xlsxSessionPaths.set(handle, resolve(path));
        registerDocumentRecord({
          id: handle,
          format: "xlsx",
          name: basename(resolve(path)),
          sourcePath: resolve(path),
        });
        return ok({ handle, path: resolve(path), summary: inspectXlsxSnapshot(agent.getSnapshot()) });
      } catch (err) {
        return fail(`xlsx_load failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── xlsx_save ─────────────────────────────────────────────────────────
  server.registerTool(
    "xlsx_save",
    {
      description:
        "Serialize the current xlsx snapshot back to disk. Defaults to the path passed to xlsx_load; pass `out_path` to write elsewhere.",
      inputSchema: {
        handle: z.string().describe("Handle returned by xlsx_load."),
        out_path: z
          .string()
          .optional()
          .describe("Optional output path. Defaults to the original path passed to xlsx_load."),
      },
    },
    async ({ handle, out_path }) => {
      try {
        const agent = lookupXlsxAgent(handle);
        agent.getPendingMutations().forEach((m) => agent.approveMutation(m.id));
        const target = out_path ? resolve(out_path) : xlsxSessionPaths.get(handle);
        if (!target) {
          return fail(`xlsx_save: no path known for handle "${handle}". Pass out_path explicitly.`);
        }
        const buf = Buffer.from(await agent.exportFile());
        await writeFile(target, buf);
        return ok({ wrote: target, bytes: buf.byteLength, revision: agent.getSnapshot().revision });
      } catch (err) {
        return fail(`xlsx_save failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── xlsx_inspect ──────────────────────────────────────────────────────
  server.registerTool(
    "xlsx_inspect",
    {
      description: "Return a structural summary (sheets, cells, parts, comments, merges).",
      inputSchema: { handle: z.string() },
    },
    async ({ handle }) => {
      try {
        const agent = lookupXlsxAgent(handle);
        return ok(inspectXlsxSnapshot(agent.getSnapshot()));
      } catch (err) {
        return fail(`xlsx_inspect failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── xlsx_list_sheets ──────────────────────────────────────────────────
  server.registerTool(
    "xlsx_list_sheets",
    {
      description: "List sheets in tab order, with name, index, kind, and visibility state.",
      inputSchema: { handle: z.string() },
    },
    async ({ handle }) => {
      try {
        const agent = lookupXlsxAgent(handle);
        return ok({ sheets: agent.listSheets() });
      } catch (err) {
        return fail(`xlsx_list_sheets failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── xlsx_get_text ─────────────────────────────────────────────────────
  server.registerTool(
    "xlsx_get_text",
    {
      description: "Return one or more sheets as Markdown (default) or as a sparse JSON cell projection.",
      inputSchema: {
        handle: z.string(),
        format: z.enum(["markdown", "json"]).optional().default("markdown"),
        sheet: z.string().optional(),
        range: z.string().optional().describe("Optional A1 range; when set, also requires `sheet`."),
        max_rows: z.number().int().positive().optional(),
        max_cols: z.number().int().positive().optional(),
      },
    },
    async ({ handle, format, sheet, range, max_rows, max_cols }) => {
      try {
        const agent = lookupXlsxAgent(handle);
        const fmt = format ?? "markdown";
        switch (fmt) {
          case "markdown": {
            const md = agent.toMarkdown({
              ...(sheet ? { sheet } : {}),
              ...(max_rows !== undefined ? { maxRows: max_rows } : {}),
              ...(max_cols !== undefined ? { maxCols: max_cols } : {}),
            });
            return ok({ format: fmt, content: md });
          }
          case "json":
            return ok(xlsxRangeToJson(agent, sheet, range));
          default: {
            const _exhaustive: never = fmt;
            void _exhaustive;
            return fail(`unknown format: ${String(format)}`);
          }
        }
      } catch (err) {
        return fail(`xlsx_get_text failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── xlsx_search ───────────────────────────────────────────────────────
  server.registerTool(
    "xlsx_search",
    {
      description: "Search workbook cells for text. Optional sheet filter, case-sensitive, and regex flags.",
      inputSchema: {
        handle: z.string(),
        query: z.string().min(1),
        sheet: z.string().optional(),
        case_sensitive: z.boolean().optional().default(false),
        regex: z.boolean().optional().default(false),
      },
    },
    async ({ handle, query, sheet, case_sensitive, regex }) => {
      try {
        const agent = lookupXlsxAgent(handle);
        const results = agent.search({
          query,
          caseSensitive: case_sensitive ?? false,
          regex: regex ?? false,
          ...(sheet ? { sheet } : {}),
        });
        return ok({ matches: results });
      } catch (err) {
        return fail(`xlsx_search failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── xlsx_apply_command ────────────────────────────────────────────────
  server.registerTool(
    "xlsx_apply_command",
    {
      description:
        "Apply an xlsx command (e.g. `xlsx:set-cell-value`, `xlsx:add-sheet`). Pass an arbitrary payload object — schemas live in `@officeai/xlsx/commands/payloads`.",
      inputSchema: {
        handle: z.string(),
        type: z.string().describe('Command type, e.g. "xlsx:set-cell-value"'),
        payload: z.record(z.string(), z.unknown()).describe("Command payload."),
        source: z.enum(["agent", "human", "system"]).optional().default("agent"),
        agent_id: z.string().optional().default("officeai-mcp"),
        auto_approve: z
          .boolean()
          .optional()
          .default(true)
          .describe(
            "When true (default) any pending mutation produced by an `agent` source is immediately approved. Set false to leave it pending in the bus for downstream review."
          ),
      },
    },
    async ({ handle, type, payload, source, agent_id, auto_approve }) => {
      try {
        return await applyXlsxCommand(handle, type, payload, source, agent_id, auto_approve);
      } catch (err) {
        return fail(`xlsx_apply_command failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── xlsx_list_pending ─────────────────────────────────────────────────
  server.registerTool(
    "xlsx_list_pending",
    {
      description:
        "List xlsx mutations still in `pending` state on the bus (typically agent-authored writes invoked with auto_approve=false).",
      inputSchema: { handle: z.string() },
    },
    async ({ handle }) => {
      try {
        const agent = lookupXlsxAgent(handle);
        const pending = agent.getPendingMutations().map((m) => ({
          id: m.id,
          command: { type: m.command.type, source: m.command.source },
          ...(m.command.source === "agent" && "agentId" in m.command ? { agentId: m.command.agentId } : {}),
          revision: m.after.revision,
          status: m.status,
        }));
        return ok({ pending });
      } catch (err) {
        return fail(`xlsx_list_pending failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── xlsx_approve ──────────────────────────────────────────────────────
  server.registerTool(
    "xlsx_approve",
    {
      description:
        "Approve a pending xlsx mutation. After approval the mutation is committed and shows up in the snapshot's history.",
      inputSchema: {
        handle: z.string(),
        mutation_id: z.string().describe("Mutation id from xlsx_list_pending."),
      },
    },
    async ({ handle, mutation_id }) => {
      try {
        const agent = lookupXlsxAgent(handle);
        agent.approveMutation(mutation_id);
        return ok({ approved: mutation_id, revision: agent.getSnapshot().revision });
      } catch (err) {
        return fail(`xlsx_approve failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── xlsx_reject ───────────────────────────────────────────────────────
  server.registerTool(
    "xlsx_reject",
    {
      description:
        "Reject a pending xlsx mutation with an optional human-readable reason. The snapshot is unaffected; the mutation is dropped from the pending queue.",
      inputSchema: {
        handle: z.string(),
        mutation_id: z.string().describe("Mutation id from xlsx_list_pending."),
        reason: z.string().optional(),
      },
    },
    async ({ handle, mutation_id, reason }) => {
      try {
        const agent = lookupXlsxAgent(handle);
        agent.rejectMutation(mutation_id);
        return ok({ rejected: mutation_id, ...(reason ? { reason } : {}) });
      } catch (err) {
        return fail(`xlsx_reject failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── xlsx_undo ─────────────────────────────────────────────────────────
  server.registerTool(
    "xlsx_undo",
    {
      description:
        "Undo the most recent approved mutation on this xlsx handle. No-op when the history is empty. Returns the resulting revision and a `did_undo` flag.",
      inputSchema: { handle: z.string() },
    },
    async ({ handle }) => {
      try {
        const agent = lookupXlsxAgent(handle);
        const m = agent.undo();
        return ok({
          did_undo: m !== null,
          revision: agent.getSnapshot().revision,
          can_undo: agent.canUndo(),
          can_redo: agent.canRedo(),
          undone: m ? { id: m.id, type: m.command.type } : null,
        });
      } catch (err) {
        return fail(`xlsx_undo failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── xlsx_redo ─────────────────────────────────────────────────────────
  server.registerTool(
    "xlsx_redo",
    {
      description:
        "Re-apply the most recently undone mutation on this xlsx handle. No-op when the redo stack is empty (e.g. immediately after a fresh authored mutation, which kills the redo trail).",
      inputSchema: { handle: z.string() },
    },
    async ({ handle }) => {
      try {
        const agent = lookupXlsxAgent(handle);
        const m = agent.redo();
        return ok({
          did_redo: m !== null,
          revision: agent.getSnapshot().revision,
          can_undo: agent.canUndo(),
          can_redo: agent.canRedo(),
          redone: m ? { id: m.id, type: m.command.type, status: m.status } : null,
        });
      } catch (err) {
        return fail(`xlsx_redo failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── xlsx_diff ─────────────────────────────────────────────────────────
  server.registerTool(
    "xlsx_diff",
    {
      description:
        "Diff two loaded xlsx handles, OR diff a handle against the file on disk it was loaded from. Pass either {before, after} (two handles) or {handle, against: 'disk'}.",
      inputSchema: {
        before: z.string().optional(),
        after: z.string().optional(),
        handle: z.string().optional(),
        against: z.enum(["disk"]).optional(),
      },
    },
    async ({ before, after, handle, against }) => {
      try {
        if (before && after) {
          const a = lookupXlsxAgent(before);
          const b = lookupXlsxAgent(after);
          return ok(diffXlsxSnapshots(a.getSnapshot(), b.getSnapshot()));
        }
        if (handle && against === "disk") {
          const agent = lookupXlsxAgent(handle);
          const path = xlsxSessionPaths.get(handle);
          if (!path) return fail(`xlsx_diff: no on-disk path for handle "${handle}".`);
          const buf = await readFile(path);
          const baseline = await XlsxAgent.fromBuffer(buf);
          return ok(diffXlsxSnapshots(baseline.getSnapshot(), agent.getSnapshot()));
        }
        return fail("xlsx_diff: pass either {before, after} (two handles) or {handle, against: 'disk'}.");
      } catch (err) {
        return fail(`xlsx_diff failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── Convenience write tools ───────────────────────────────────────────
  // Each one collapses internally to xlsx_apply_command but keeps the
  // wire payload trivial for LLM clients.

  const convenienceCommonSchema = {
    source: z.enum(["agent", "human", "system"]).optional().default("agent"),
    agent_id: z.string().optional().default("officeai-mcp"),
    auto_approve: z.boolean().optional().default(true),
  } as const;

  server.registerTool(
    "xlsx_set_cell",
    {
      description: "Set a single cell's literal value (collapses to xlsx:set-cell-value).",
      inputSchema: {
        handle: z.string(),
        sheet: z.string(),
        ref: z.string().describe("A1 single-cell ref, e.g. 'B2'"),
        value: z.unknown(),
        ...convenienceCommonSchema,
      },
    },
    async ({ handle, sheet, ref, value, source, agent_id, auto_approve }) => {
      try {
        return await applyXlsxCommand(
          handle,
          "xlsx:set-cell-value",
          { sheet, ref, value },
          source,
          agent_id,
          auto_approve
        );
      } catch (err) {
        return fail(`xlsx_set_cell failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  server.registerTool(
    "xlsx_set_formula",
    {
      description: "Set a single cell's formula (collapses to xlsx:set-cell-formula).",
      inputSchema: {
        handle: z.string(),
        sheet: z.string(),
        ref: z.string(),
        formula: z.string(),
        ...convenienceCommonSchema,
      },
    },
    async ({ handle, sheet, ref, formula, source, agent_id, auto_approve }) => {
      try {
        return await applyXlsxCommand(
          handle,
          "xlsx:set-cell-formula",
          { sheet, ref, formula },
          source,
          agent_id,
          auto_approve
        );
      } catch (err) {
        return fail(`xlsx_set_formula failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  server.registerTool(
    "xlsx_set_range",
    {
      description: "Set a 2-D matrix of cell values (collapses to xlsx:set-range-values).",
      inputSchema: {
        handle: z.string(),
        sheet: z.string(),
        range: z.string(),
        values: z.array(z.array(z.unknown())),
        ...convenienceCommonSchema,
      },
    },
    async ({ handle, sheet, range, values, source, agent_id, auto_approve }) => {
      try {
        return await applyXlsxCommand(
          handle,
          "xlsx:set-range-values",
          { sheet, range, values },
          source,
          agent_id,
          auto_approve
        );
      } catch (err) {
        return fail(`xlsx_set_range failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  server.registerTool(
    "xlsx_set_format",
    {
      description: "Apply a CellFormatPatch to a range (collapses to xlsx:set-cell-format).",
      inputSchema: {
        handle: z.string(),
        sheet: z.string(),
        range: z.string(),
        format: z.record(z.string(), z.unknown()),
        ...convenienceCommonSchema,
      },
    },
    async ({ handle, sheet, range, format, source, agent_id, auto_approve }) => {
      try {
        return await applyXlsxCommand(
          handle,
          "xlsx:set-cell-format",
          { sheet, range, format },
          source,
          agent_id,
          auto_approve
        );
      } catch (err) {
        return fail(`xlsx_set_format failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  server.registerTool(
    "xlsx_add_sheet",
    {
      description: "Append (or insert at `at`) a new worksheet (collapses to xlsx:add-sheet).",
      inputSchema: {
        handle: z.string(),
        name: z.string(),
        at: z.number().int().nonnegative().optional(),
        ...convenienceCommonSchema,
      },
    },
    async ({ handle, name, at, source, agent_id, auto_approve }) => {
      try {
        const payload: Record<string, unknown> = { name };
        if (at !== undefined) payload.at = at;
        return await applyXlsxCommand(handle, "xlsx:add-sheet", payload, source, agent_id, auto_approve);
      } catch (err) {
        return fail(`xlsx_add_sheet failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  server.registerTool(
    "xlsx_rename_sheet",
    {
      description: "Rename a worksheet (collapses to xlsx:rename-sheet).",
      inputSchema: {
        handle: z.string(),
        name: z.string(),
        new_name: z.string(),
        ...convenienceCommonSchema,
      },
    },
    async ({ handle, name, new_name, source, agent_id, auto_approve }) => {
      try {
        return await applyXlsxCommand(
          handle,
          "xlsx:rename-sheet",
          { name, newName: new_name },
          source,
          agent_id,
          auto_approve
        );
      } catch (err) {
        return fail(`xlsx_rename_sheet failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  registerStructuralTool(server, "xlsx_insert_row", "xlsx:insert-row", "Insert N blank rows");
  registerStructuralTool(server, "xlsx_insert_column", "xlsx:insert-column", "Insert N blank columns");
  registerStructuralTool(server, "xlsx_delete_row", "xlsx:delete-row", "Delete N rows");
  registerStructuralTool(server, "xlsx_delete_column", "xlsx:delete-column", "Delete N columns");

  server.registerTool(
    "xlsx_merge",
    {
      description: "Merge an A1 range covering ≥2 cells (collapses to xlsx:merge-cells).",
      inputSchema: {
        handle: z.string(),
        sheet: z.string(),
        range: z.string(),
        ...convenienceCommonSchema,
      },
    },
    async ({ handle, sheet, range, source, agent_id, auto_approve }) => {
      try {
        return await applyXlsxCommand(
          handle,
          "xlsx:merge-cells",
          { sheet, range },
          source,
          agent_id,
          auto_approve
        );
      } catch (err) {
        return fail(`xlsx_merge failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  server.registerTool(
    "xlsx_unmerge",
    {
      description: "Unmerge an existing merged range (collapses to xlsx:unmerge-cells).",
      inputSchema: {
        handle: z.string(),
        sheet: z.string(),
        range: z.string(),
        ...convenienceCommonSchema,
      },
    },
    async ({ handle, sheet, range, source, agent_id, auto_approve }) => {
      try {
        return await applyXlsxCommand(
          handle,
          "xlsx:unmerge-cells",
          { sheet, range },
          source,
          agent_id,
          auto_approve
        );
      } catch (err) {
        return fail(`xlsx_unmerge failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  server.registerTool(
    "xlsx_add_comment",
    {
      description: "Attach a classic note to a single cell (collapses to xlsx:add-comment).",
      inputSchema: {
        handle: z.string(),
        sheet: z.string(),
        ref: z.string(),
        text: z.string(),
        author: z.string(),
        ...convenienceCommonSchema,
      },
    },
    async ({ handle, sheet, ref, text, author, source, agent_id, auto_approve }) => {
      try {
        return await applyXlsxCommand(
          handle,
          "xlsx:add-comment",
          { sheet, ref, text, author },
          source,
          agent_id,
          auto_approve
        );
      } catch (err) {
        return fail(`xlsx_add_comment failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );
}

function registerStructuralTool(
  server: McpServer,
  toolName: string,
  commandType: string,
  description: string
): void {
  server.registerTool(
    toolName,
    {
      description: `${description} starting at a 1-based index (collapses to ${commandType}).`,
      inputSchema: {
        handle: z.string(),
        sheet: z.string(),
        at: z.number().int().positive(),
        count: z.number().int().positive(),
        source: z.enum(["agent", "human", "system"]).optional().default("agent"),
        agent_id: z.string().optional().default("officeai-mcp"),
        auto_approve: z.boolean().optional().default(true),
      },
    },
    async ({ handle, sheet, at, count, source, agent_id, auto_approve }) => {
      try {
        return await applyXlsxCommand(
          handle,
          commandType,
          { sheet, at, count },
          source,
          agent_id,
          auto_approve
        );
      } catch (err) {
        return fail(`${toolName} failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );
}

/**
 * Shared apply pipeline used by `xlsx_apply_command` and every
 * convenience tool. Returns an `ok(...)` payload describing the
 * mutation status (with the post-auto-approve resolution) and the new
 * snapshot revision.
 */
async function applyXlsxCommand(
  handle: string,
  type: string,
  payload: unknown,
  source: "agent" | "human" | "system" | undefined,
  agentId: string | undefined,
  autoApprove: boolean | undefined
): Promise<ReturnType<typeof ok>> {
  const agent = lookupXlsxAgent(handle);
  const effectiveSource = source ?? "agent";
  const mutation = await agent.applyCommand({
    type,
    payload,
    source: effectiveSource,
    ...(effectiveSource === "agent" ? { agentId: agentId ?? "officeai-mcp" } : {}),
  });
  if ((autoApprove ?? true) && mutation.status === "pending") {
    agent.approveMutation(mutation.id);
  }
  return ok({
    mutation: {
      id: mutation.id,
      status: agent.getPendingMutations().some((m) => m.id === mutation.id) ? "pending" : mutation.status,
      ...(mutation.rejection ? { rejection: mutation.rejection } : {}),
    },
    revision: agent.getSnapshot().revision,
  });
}

// ──────────────────────────────────────────────────────────────────────────
// pdf_* tools
// ──────────────────────────────────────────────────────────────────────────

/**
 * Register the PDF surface as MCP tools. Mirrors `office-agent pdf …` —
 * the same JSON schemas (office-agent/pdf-<verb>@1) flow through the
 * read tools, and every mutation tool returns the `{ schema, in, out,
 * bytes, summary }` envelope a CLI caller would have seen.
 *
 * Reads are handle-based (cheap re-projection from a single parsed
 * snapshot). Mutations are file-in / file-out and re-load the input
 * each call: the PdfAgent command bus only covers the page-rotation +
 * reorder subset today, so we drive `@officeai/pdf-edit` /
 * `@officeai/pdf-forms` directly here for parity with the CLI.
 */
function registerPdfTools(server: McpServer): void {
  // ── pdf_load ──────────────────────────────────────────────────────────
  server.registerTool(
    "pdf_load",
    {
      description:
        "Load a .pdf file from disk. Returns an opaque `handle` plus a metadata summary; pass the handle to subsequent pdf_* read tools. PDF mutation tools take paths directly and do not need a handle.",
      inputSchema: {
        path: z.string().describe("Absolute or workspace-relative path to a .pdf file."),
      },
    },
    async ({ path }) => {
      try {
        const buf = await readFile(resolve(path));
        const bytes = new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
        const agent = await PdfAgent.fromBuffer(bytes);
        const handle = randomUUID();
        pdfSessions.set(handle, { agent, bytes });
        pdfSessionPaths.set(handle, resolve(path));
        registerDocumentRecord({
          id: handle,
          format: "pdf",
          name: basename(resolve(path)),
          sourcePath: resolve(path),
        });
        return ok({ handle, path: resolve(path), summary: projectMetadata(agent.getSnapshot()) });
      } catch (err) {
        return fail(`pdf_load failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── pdf_metadata ──────────────────────────────────────────────────────
  server.registerTool(
    "pdf_metadata",
    {
      description:
        "Return PDF document metadata, page count, signature count, encryption flags, engine, version, and linearization. Schema: office-agent/pdf-read-metadata@1.",
      inputSchema: { handle: z.string() },
    },
    async ({ handle }) => {
      try {
        return ok(projectMetadata(lookupPdfSession(handle).agent.getSnapshot()));
      } catch (err) {
        return fail(`pdf_metadata failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── pdf_read_page ─────────────────────────────────────────────────────
  server.registerTool(
    "pdf_read_page",
    {
      description:
        "Project a single page (1-indexed) with size, rotation, text layer flag, text, annotations, and form fields. Schema: office-agent/pdf-read-page@1.",
      inputSchema: { handle: z.string(), page: z.number().int().positive() },
    },
    async ({ handle, page }) => {
      try {
        return ok(projectPage(lookupPdfSession(handle).agent.getSnapshot(), page));
      } catch (err) {
        return fail(`pdf_read_page failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── pdf_outline ───────────────────────────────────────────────────────
  server.registerTool(
    "pdf_outline",
    {
      description:
        "Return the recursive outline tree (each entry has title, optional pageNumber, and children). Schema: office-agent/pdf-read-outline@1.",
      inputSchema: { handle: z.string() },
    },
    async ({ handle }) => {
      try {
        return ok(projectOutline(lookupPdfSession(handle).agent.getSnapshot()));
      } catch (err) {
        return fail(`pdf_outline failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── pdf_annotations ───────────────────────────────────────────────────
  server.registerTool(
    "pdf_annotations",
    {
      description:
        "Return a flat list of annotations in the PDF, optionally restricted to one 1-indexed page. Schema: office-agent/pdf-read-annotations@1.",
      inputSchema: {
        handle: z.string(),
        page: z.number().int().positive().optional(),
      },
    },
    async ({ handle, page }) => {
      try {
        const snap = lookupPdfSession(handle).agent.getSnapshot();
        return ok(projectAnnotations(snap, page !== undefined ? { page } : undefined));
      } catch (err) {
        return fail(`pdf_annotations failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── pdf_list_form_fields ──────────────────────────────────────────────
  server.registerTool(
    "pdf_list_form_fields",
    {
      description:
        "List AcroForm fields with name, type, value, options, readOnly, required, and the page they live on. Schema: office-agent/pdf-list-form-fields@1.",
      inputSchema: { handle: z.string() },
    },
    async ({ handle }) => {
      try {
        const session = lookupPdfSession(handle);
        return ok(await projectFormFields(session.agent.getSnapshot(), session.bytes));
      } catch (err) {
        return fail(`pdf_list_form_fields failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── pdf_search ────────────────────────────────────────────────────────
  server.registerTool(
    "pdf_search",
    {
      description:
        "Search every page's text layer for `query`, optionally as a regex / case-sensitive. Returns per-page hits ({ page, start, end, preview, match }). Schema: office-agent/pdf-search-text@1.",
      inputSchema: {
        handle: z.string(),
        query: z.string().min(1),
        regex: z.boolean().optional().default(false),
        case_sensitive: z.boolean().optional().default(false),
      },
    },
    async ({ handle, query, regex, case_sensitive }) => {
      try {
        return ok(
          projectSearch(lookupPdfSession(handle).agent, {
            query,
            regex: regex === true,
            caseSensitive: case_sensitive === true,
          })
        );
      } catch (err) {
        return fail(`pdf_search failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── pdf_export_markdown ───────────────────────────────────────────────
  server.registerTool(
    "pdf_export_markdown",
    {
      description:
        "Render the loaded PDF as Markdown via PdfAgent.toMarkdown(). Returns the markdown payload directly. The CLI counterpart `office-agent pdf export-markdown` writes to disk when --out is given.",
      inputSchema: { handle: z.string() },
    },
    async ({ handle }) => {
      try {
        const md = lookupPdfSession(handle).agent.toMarkdown();
        return ok({
          schema: "office-agent/pdf-export-markdown@1",
          format: "markdown",
          content: md,
        });
      } catch (err) {
        return fail(`pdf_export_markdown failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  // ── Mutations: file in → file out (no handle, mirrors CLI) ────────────
  registerPdfMutationTool(
    server,
    "pdf_rotate_pages",
    "office-agent/pdf-rotate-pages@1",
    {
      pages: z.array(z.number().int().positive()).min(1),
      delta: z.number().int(),
    },
    async (bytes, opts) => {
      const pages = opts.pages as number[];
      const delta = opts.delta as number;
      if (![90, 180, 270, -90, -180, -270].includes(delta)) {
        throw new Error(`pdf_rotate_pages: delta must be ±90/±180/±270 (got ${delta})`);
      }
      const out = await rotatePages(bytes, {
        pages,
        delta: delta as 90 | 180 | 270 | -90 | -180 | -270,
      });
      return { bytes: out, summary: `rotated ${pages.length} page(s) by ${delta}°` };
    }
  );

  registerPdfMutationTool(
    server,
    "pdf_reorder_pages",
    "office-agent/pdf-reorder-pages@1",
    { order: z.array(z.number().int().positive()).min(1) },
    async (bytes, opts) => {
      const order = opts.order as number[];
      return {
        bytes: await reorderPages(bytes, { order }),
        summary: `reordered ${order.length} pages`,
      };
    }
  );

  registerPdfMutationTool(
    server,
    "pdf_delete_pages",
    "office-agent/pdf-delete-pages@1",
    { pages: z.array(z.number().int().positive()).min(1) },
    async (bytes, opts) => {
      const pages = opts.pages as number[];
      return {
        bytes: await deletePages(bytes, { pages }),
        summary: `deleted ${pages.length} page(s)`,
      };
    }
  );

  registerPdfMutationTool(
    server,
    "pdf_extract_pages",
    "office-agent/pdf-extract-pages@1",
    { pages: z.array(z.number().int().positive()).min(1) },
    async (bytes, opts) => {
      const pages = opts.pages as number[];
      return {
        bytes: await extractPages(bytes, { pages }),
        summary: `extracted ${pages.length} page(s)`,
      };
    }
  );

  // pdf_merge: variadic input list, no source-bytes parameter.
  server.registerTool(
    "pdf_merge",
    {
      description: "Concatenate two-or-more PDFs into a single document. Schema: office-agent/pdf-merge@1.",
      inputSchema: {
        inputs: z.array(z.string()).min(2),
        out_path: z.string(),
      },
    },
    async ({ inputs, out_path }) => {
      try {
        const buffers = await Promise.all(
          inputs.map(async (p) => {
            const b = await readFile(resolve(p));
            return new Uint8Array(b.buffer, b.byteOffset, b.byteLength);
          })
        );
        const out = await mergePdfs({ inputs: buffers });
        const target = resolve(out_path);
        const buf = Buffer.from(out);
        await writeFile(target, buf);
        return ok({
          schema: "office-agent/pdf-merge@1",
          inputs,
          out: out_path,
          bytes: buf.byteLength,
          summary: `merged ${inputs.length} PDFs`,
        });
      } catch (err) {
        return fail(`pdf_merge failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );

  registerPdfMutationTool(
    server,
    "pdf_set_metadata",
    "office-agent/pdf-set-metadata@1",
    {
      title: z.string().optional(),
      author: z.string().optional(),
      subject: z.string().optional(),
      keywords: z.string().optional(),
      creator: z.string().optional(),
      producer: z.string().optional(),
    },
    async (bytes, opts) => {
      const patch: Record<string, string> = {};
      for (const k of ["title", "author", "subject", "keywords", "creator", "producer"] as const) {
        const v = opts[k];
        if (typeof v === "string") patch[k] = v;
      }
      if (Object.keys(patch).length === 0) {
        throw new Error(
          "pdf_set_metadata: pass at least one of title/author/subject/keywords/creator/producer"
        );
      }
      return {
        bytes: await setMetadata(bytes, patch),
        summary: `set ${Object.keys(patch).join(", ")}`,
      };
    }
  );

  registerPdfMutationTool(
    server,
    "pdf_add_watermark",
    "office-agent/pdf-add-watermark@1",
    {
      text: z.string().min(1),
      opacity: z.number().min(0).max(1).optional(),
      rotation: z.number().optional(),
      font_size: z.number().positive().optional(),
      pages: z.array(z.number().int().positive()).optional(),
    },
    async (bytes, opts) => {
      const text = opts.text as string;
      const opacity = opts.opacity as number | undefined;
      const rotation = opts.rotation as number | undefined;
      const fontSize = opts.font_size as number | undefined;
      const pages = opts.pages as number[] | undefined;
      return {
        bytes: await addWatermark(bytes, {
          text,
          ...(opacity !== undefined ? { opacity } : {}),
          ...(rotation !== undefined ? { rotate: rotation } : {}),
          ...(fontSize !== undefined ? { fontSize } : {}),
          ...(pages ? { pages } : {}),
        }),
        summary: `stamped "${text}" watermark`,
      };
    }
  );

  registerPdfMutationTool(
    server,
    "pdf_add_page_numbers",
    "office-agent/pdf-add-page-numbers@1",
    {
      start: z.number().int().positive().optional(),
      position: z
        .enum(["top-left", "top-center", "top-right", "bottom-left", "bottom-center", "bottom-right"])
        .optional(),
      font_size: z.number().positive().optional(),
      margin: z.number().nonnegative().optional(),
      format: z.string().optional(),
    },
    async (bytes, opts) => {
      const start = opts.start as number | undefined;
      const position = opts.position as
        | "top-left"
        | "top-center"
        | "top-right"
        | "bottom-left"
        | "bottom-center"
        | "bottom-right"
        | undefined;
      const fontSize = opts.font_size as number | undefined;
      const margin = opts.margin as number | undefined;
      const format = opts.format as string | undefined;
      return {
        bytes: await addPageNumbers(bytes, {
          ...(start !== undefined ? { startAt: start } : {}),
          ...(position ? { position } : {}),
          ...(fontSize !== undefined ? { fontSize } : {}),
          ...(margin !== undefined ? { margin } : {}),
          ...(format ? { format } : {}),
        }),
        summary: `added page numbers (position=${position ?? "bottom-center"})`,
      };
    }
  );

  registerPdfMutationTool(
    server,
    "pdf_fill_form",
    "office-agent/pdf-fill-form@1",
    { values: z.record(z.string(), z.unknown()) },
    async (bytes, opts) => {
      const values = opts.values as Record<string, unknown>;
      const coerced: Record<string, string | boolean | ReadonlyArray<string>> = {};
      for (const [k, v] of Object.entries(values)) {
        if (typeof v === "string" || typeof v === "boolean") coerced[k] = v;
        else if (Array.isArray(v) && v.every((x) => typeof x === "string")) {
          coerced[k] = v as ReadonlyArray<string>;
        } else {
          throw new Error(`pdf_fill_form: field "${k}" expects string|boolean|string[] (got ${typeof v})`);
        }
      }
      return {
        bytes: await fillForm(bytes, { values: coerced }),
        summary: `filled ${Object.keys(coerced).length} form field(s)`,
      };
    }
  );

  registerPdfMutationTool(
    server,
    "pdf_flatten_form",
    "office-agent/pdf-flatten-form@1",
    {},
    async (bytes) => ({ bytes: await flattenForm(bytes), summary: "flattened form fields" })
  );

  registerPdfMutationTool(server, "pdf_reset_form", "office-agent/pdf-reset-form@1", {}, async (bytes) => ({
    bytes: await resetForm(bytes),
    summary: "reset form fields to defaults",
  }));
}

/**
 * Shared registration for every file-in / file-out PDF mutation tool.
 * Reads `in_path` from disk, runs `mutate(bytes, opts)`, writes the
 * resulting bytes to `out_path`, and returns the standard
 * `{ schema, in, out, bytes, summary }` envelope.
 */
function registerPdfMutationTool(
  server: McpServer,
  name: string,
  schema: string,
  extraSchema: z.ZodRawShape,
  mutate: (
    bytes: Uint8Array,
    opts: Record<string, unknown>
  ) => Promise<{ bytes: Uint8Array; summary: string }>
): void {
  server.registerTool(
    name,
    {
      description: `${name} — mutation tool emitting JSON envelope ${schema}. File-in / file-out.`,
      inputSchema: {
        in_path: z.string().describe("Path to the source .pdf"),
        out_path: z.string().describe("Path to write the resulting .pdf"),
        ...extraSchema,
      },
    },
    async (args: Record<string, unknown>) => {
      try {
        const inPath = args.in_path as string;
        const outPath = args.out_path as string;
        const buf = await readFile(resolve(inPath));
        const bytes = new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength);
        const { bytes: outBytes, summary } = await mutate(bytes, args);
        const target = resolve(outPath);
        const outBuf = Buffer.from(outBytes);
        await writeFile(target, outBuf);
        return ok({
          schema,
          in: inPath,
          out: outPath,
          bytes: outBuf.byteLength,
          summary,
        });
      } catch (err) {
        return fail(`${name} failed: ${err instanceof Error ? err.message : String(err)}`);
      }
    }
  );
}

/**
 * Bridge each format's `ActionDescriptor[]` into the generic
 * `registerActionsAsMcpTools` adapter. Every catalogue entry that
 * declares `agentCallable`, has a non-null `commandType`, and supplies
 * catalogue-owned `args` + `buildPayload` becomes an MCP tool of the
 * form `<format>_<verb>_<noun>` (e.g. `docx_delete_row`).
 *
 * Hand-rolled tools registered earlier in `createMcpServer` stay
 * authoritative because the adapter doesn't overwrite existing tools
 * with the same name — but no overlap should exist, since the read
 * tools have `commandType: null` and are skipped here.
 */
function registerCatalogueActionTools(server: McpServer): void {
  registerActionsAsMcpTools(server, docxActions, dispatchContextFor("docx"));
  registerActionsAsMcpTools(server, xlsxActions, dispatchContextFor("xlsx"));
  registerActionsAsMcpTools(server, pptxActions, dispatchContextFor("pptx"));
}

function dispatchContextFor(format: "docx" | "xlsx" | "pptx"): McpDispatchContext {
  switch (format) {
    case "docx":
      return {
        format,
        lookup: (handle) => lookupAgent(handle),
        async dispatch({ handle, agent, commandType, payload, outPath, source, approve }) {
          const docxAgent = agent as DocxAgent;
          const mut = await docxAgent.applyCommand({
            type: commandType,
            payload,
            source,
          } as Parameters<DocxAgent["applyCommand"]>[0]);
          if (mut.status === "rejected") {
            throw new Error(
              `${commandType} rejected: ${mut.rejection?.code ?? "unknown"} — ${mut.rejection?.message ?? ""}`
            );
          }
          if (approve && mut.status === "pending") {
            docxAgent.approveMutation(mut.id);
          }
          let wrote: string | null = null;
          if (outPath) {
            const target = resolve(outPath);
            const buf = Buffer.from(await docxAgent.exportFile());
            await writeFile(target, buf);
            wrote = target;
          }
          return {
            handle,
            commandType,
            mutationId: mut.id,
            status: approve ? "approved" : mut.status,
            revision: docxAgent.getSnapshot().revision,
            wrote,
          };
        },
      };
    case "xlsx":
      return {
        format,
        lookup: (handle) => lookupXlsxAgent(handle),
        async dispatch({ handle, agent, commandType, payload, outPath, source, approve }) {
          const xlsxAgent = agent as XlsxAgent;
          const mut = await xlsxAgent.applyCommand({
            type: commandType,
            payload,
            source,
          } as Parameters<XlsxAgent["applyCommand"]>[0]);
          if (mut.status === "rejected") {
            throw new Error(
              `${commandType} rejected: ${mut.rejection?.code ?? "unknown"} — ${mut.rejection?.message ?? ""}`
            );
          }
          if (approve && mut.status === "pending") {
            xlsxAgent.approveMutation(mut.id);
          }
          let wrote: string | null = null;
          if (outPath) {
            const target = resolve(outPath);
            const buf = Buffer.from(await xlsxAgent.exportFile());
            await writeFile(target, buf);
            wrote = target;
          }
          return {
            handle,
            commandType,
            mutationId: mut.id,
            status: approve ? "approved" : mut.status,
            revision: xlsxAgent.getSnapshot().revision,
            wrote,
          };
        },
      };
    case "pptx":
      return {
        format,
        lookup: (handle) => lookupPptxAgent(handle),
        async dispatch({ handle, agent, commandType, payload, outPath, source, approve }) {
          const pptxAgent = agent as PptxAgent;
          const mut = await pptxAgent.applyCommand({
            type: commandType,
            payload,
            source,
          } as Parameters<PptxAgent["applyCommand"]>[0]);
          if (mut.status === "rejected") {
            throw new Error(
              `${commandType} rejected: ${mut.rejection?.code ?? "unknown"} — ${mut.rejection?.message ?? ""}`
            );
          }
          if (approve && mut.status === "pending") {
            pptxAgent.approveMutation(mut.id);
          }
          let wrote: string | null = null;
          if (outPath) {
            const target = resolve(outPath);
            const buf = Buffer.from(await pptxAgent.exportFile());
            await writeFile(target, buf);
            wrote = target;
          }
          return {
            handle,
            commandType,
            mutationId: mut.id,
            status: approve ? "approved" : mut.status,
            revision: pptxAgent.getSnapshot().revision,
            wrote,
          };
        },
      };
    default: {
      const _exhaustive: never = format;
      void _exhaustive;
      throw new Error(`unknown format: ${String(format)}`);
    }
  }
}

/**
 * Run the MCP server over stdio. This is the entry point used by
 * `office-agent mcp` — it never returns under normal operation; the SDK's
 * stdio transport keeps the process alive until the parent closes stdin.
 */
export async function runMcpStdioServer(): Promise<void> {
  const server = createMcpServer();
  const transport = new StdioServerTransport();
  await server.connect(transport);
  // Keep the event loop alive until stdin closes.
  await new Promise<void>((resolveP) => {
    process.stdin.on("close", () => resolveP());
  });
}
