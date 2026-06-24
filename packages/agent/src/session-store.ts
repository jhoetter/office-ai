import { randomUUID } from "node:crypto";
import { copyFile, mkdir, readFile, readdir, rename, rm, stat, writeFile } from "node:fs/promises";
import { homedir } from "node:os";
import { basename, dirname, join, resolve } from "node:path";

export const DATA_DIR_SCHEMA_VERSION = 1;
export const SESSION_RECORD_SCHEMA_VERSION = 1;
export const DOCUMENT_RECORD_SCHEMA_VERSION = 1;

export type StoredOfficeFormat = "docx" | "xlsx" | "pptx" | "pdf";
export type StoredDiagnosticLevel = "info" | "warning" | "error" | "destructive";

export interface StoredDiagnostic {
  readonly level: StoredDiagnosticLevel;
  readonly code: string;
  readonly message: string;
}

export interface StoredExportRecord {
  readonly path: string;
  readonly bytes: number;
  readonly exportedAt: string;
}

export interface StoredPendingChange {
  readonly id: string;
  readonly operation: string;
  readonly status: string;
  readonly source: string;
  readonly actorId?: string;
  readonly timestamp?: number;
  readonly diff?: unknown;
  readonly rejection?: { readonly code: string; readonly message: string };
}

export interface StoredCommandLogEntry {
  readonly schema?: "office-ai/audit-log-entry@1";
  readonly schemaVersion?: 1;
  readonly id: string;
  readonly commandId?: string;
  readonly operation: string;
  readonly status: string;
  readonly stage: string;
  readonly source: string;
  readonly actorId?: string;
  readonly recordedAt: string;
  readonly diff?: unknown;
  readonly diagnostics?: ReadonlyArray<StoredDiagnostic>;
  readonly provenance?: StoredCommandProvenance;
  readonly exportRef?: StoredCommandExportRef;
}

export interface StoredCommandProvenance {
  readonly surface: string;
  readonly actorId?: string;
  readonly clientName?: string;
  readonly sessionId?: string;
  readonly documentId?: string;
  readonly targetRevision?: number;
  readonly anchor?: unknown;
  readonly argumentsSummary?: string;
}

export interface StoredCommandExportRef {
  readonly exportedAt: string;
  readonly bytes: number;
  readonly commandIds: ReadonlyArray<string>;
}

export interface StoredSessionRecord {
  readonly schema: "office-ai/session-record@1";
  readonly schemaVersion: 1;
  readonly version: 1;
  readonly id: string;
  readonly title: string;
  readonly createdAt: string;
  readonly updatedAt: string;
  readonly documentIds: ReadonlyArray<string>;
  readonly lease?: {
    readonly pid: number;
    readonly host: string;
    readonly updatedAt: string;
  };
}

export interface StoredDocumentRecord {
  readonly schema: "office-ai/document-record@1";
  readonly schemaVersion: 1;
  readonly version: 1;
  readonly id: string;
  readonly sessionId: string;
  readonly format: StoredOfficeFormat;
  readonly name: string;
  readonly status: "ready" | "error";
  readonly sourcePath?: string;
  readonly createdAt: string;
  readonly updatedAt: string;
  readonly revision: number;
  readonly diagnostics: ReadonlyArray<StoredDiagnostic>;
  readonly exportHistory: ReadonlyArray<StoredExportRecord>;
  readonly artifacts: {
    readonly originalPath?: string;
    readonly workingPath?: string;
  };
  readonly pendingChanges: ReadonlyArray<StoredPendingChange>;
  readonly commandLog: ReadonlyArray<StoredCommandLogEntry>;
}

export type SessionStorageOperation = "ensure-dir" | "list" | "exists" | "read" | "write" | "copy" | "remove";

export interface SessionStorageCapabilities {
  readonly atomicWrite: boolean;
  readonly localPaths: boolean;
  readonly locks: "advisory" | "none";
  readonly watch: boolean;
}

export interface SessionStorageRemoveOptions {
  readonly recursive?: boolean;
  readonly force?: boolean;
}

export interface SessionStorageAdapter {
  readonly kind: string;
  readonly root: string;
  readonly capabilities: SessionStorageCapabilities;
  join(...segments: ReadonlyArray<string>): string;
  ensureDir(path: string): Promise<void>;
  list(path: string): Promise<ReadonlyArray<string>>;
  exists(path: string): Promise<boolean>;
  readBytes(path: string): Promise<Uint8Array>;
  writeBytesAtomic(path: string, bytes: Uint8Array | Buffer): Promise<void>;
  copyFromLocalFile(sourcePath: string, targetPath: string): Promise<void>;
  remove(path: string, opts?: SessionStorageRemoveOptions): Promise<void>;
}

export interface LocalSessionStoreOptions {
  readonly dataDir?: string;
  readonly storage?: SessionStorageAdapter;
}

export interface SessionStoreInspection {
  readonly schema: "office-ai/session-store-inspection@1";
  readonly dataDir: string;
  readonly schemaVersion: 1;
  readonly inspectedAt: string;
  readonly needsMigration: boolean;
  readonly diagnostics: ReadonlyArray<StoredDiagnostic>;
}

export interface SessionStoreMigrationRecord {
  readonly kind: "session" | "document";
  readonly id: string;
  readonly path: string;
  readonly backupPath: string;
  readonly fromSchema?: string;
  readonly fromVersion?: number;
  readonly toSchema: string;
  readonly toVersion: number;
}

export interface SessionStoreMigrationResult {
  readonly schema: "office-ai/session-store-migration@1";
  readonly dataDir: string;
  readonly migratedAt: string;
  readonly migrations: ReadonlyArray<SessionStoreMigrationRecord>;
  readonly diagnostics: ReadonlyArray<StoredDiagnostic>;
}

export interface SessionStoreCleanupResult {
  readonly schema: "office-ai/session-store-cleanup@1";
  readonly dataDir: string;
  readonly cleanedAt: string;
  readonly removed: ReadonlyArray<string>;
  readonly diagnostics: ReadonlyArray<StoredDiagnostic>;
}

export class SessionStoreCorruptError extends Error {
  readonly path: string;
  constructor(path: string, message: string, cause?: unknown) {
    super(`${message}: ${path}`);
    this.name = "SessionStoreCorruptError";
    this.path = path;
    if (cause !== undefined) (this as unknown as { cause: unknown }).cause = cause;
  }
}

export class SessionStoreStorageError extends Error {
  readonly operation: SessionStorageOperation;
  readonly path: string;
  readonly causeCode?: string;
  readonly diagnostic: StoredDiagnostic;

  constructor(operation: SessionStorageOperation, path: string, cause: unknown) {
    const message = cause instanceof Error ? cause.message : String(cause);
    super(`Storage ${operation} failed for ${path}: ${message}`);
    this.name = "SessionStoreStorageError";
    this.operation = operation;
    this.path = path;
    this.causeCode = (cause as NodeJS.ErrnoException | undefined)?.code;
    this.diagnostic = {
      level: "error",
      code: `storage-${operation}`,
      message: `Storage ${operation} failed for ${path}: ${message}`,
    };
    if (cause !== undefined) (this as unknown as { cause: unknown }).cause = cause;
  }
}

export class LocalFilesystemSessionStorageAdapter implements SessionStorageAdapter {
  readonly kind = "local-filesystem";
  readonly root: string;
  readonly capabilities: SessionStorageCapabilities = {
    atomicWrite: true,
    localPaths: true,
    locks: "advisory",
    watch: false,
  };

  constructor(root: string) {
    this.root = resolve(root);
  }

  join(...segments: ReadonlyArray<string>): string {
    return join(...segments);
  }

  async ensureDir(path: string): Promise<void> {
    await this.wrap("ensure-dir", path, () => mkdir(path, { recursive: true }));
  }

  async list(path: string): Promise<ReadonlyArray<string>> {
    try {
      return await readdir(path);
    } catch (err) {
      if (isEnoent(err)) return [];
      throw new SessionStoreStorageError("list", path, err);
    }
  }

  async exists(path: string): Promise<boolean> {
    try {
      await stat(path);
      return true;
    } catch (err) {
      if (isEnoent(err)) return false;
      throw new SessionStoreStorageError("exists", path, err);
    }
  }

  async readBytes(path: string): Promise<Uint8Array> {
    return this.wrap("read", path, async () => new Uint8Array(await readFile(path)));
  }

  async writeBytesAtomic(path: string, bytes: Uint8Array | Buffer): Promise<void> {
    await this.wrap("write", path, async () => {
      await mkdir(dirname(path), { recursive: true });
      const tmp = join(dirname(path), `.${basename(path)}.${process.pid}.${randomUUID()}.tmp`);
      await writeFile(tmp, bytes);
      await rename(tmp, path);
    });
  }

  async copyFromLocalFile(sourcePath: string, targetPath: string): Promise<void> {
    await this.wrap("copy", targetPath, async () => {
      await mkdir(dirname(targetPath), { recursive: true });
      await copyFile(sourcePath, targetPath);
    });
  }

  async remove(path: string, opts: SessionStorageRemoveOptions = {}): Promise<void> {
    await this.wrap("remove", path, () =>
      rm(path, { recursive: opts.recursive ?? false, force: opts.force ?? false })
    );
  }

  private async wrap<T>(operation: SessionStorageOperation, path: string, fn: () => Promise<T>): Promise<T> {
    try {
      return await fn();
    } catch (err) {
      throw new SessionStoreStorageError(operation, path, err);
    }
  }
}

export class LocalSessionStore {
  readonly dataDir: string;
  readonly storage: SessionStorageAdapter;

  constructor(opts: LocalSessionStoreOptions = {}) {
    this.storage = opts.storage ?? createLocalFilesystemSessionStorageAdapter(opts.dataDir);
    this.dataDir = this.storage.root;
  }

  async init(): Promise<void> {
    await this.storage.ensureDir(this.sessionsDir());
    await atomicWriteJson(this.storage, this.storage.join(this.dataDir, "VERSION.json"), {
      schema: "office-ai/data-dir@1",
      schemaVersion: DATA_DIR_SCHEMA_VERSION,
      version: DATA_DIR_SCHEMA_VERSION,
      createdOrUpdatedAt: new Date().toISOString(),
    });
  }

  sessionDir(sessionId: string): string {
    return this.storage.join(this.sessionsDir(), safeSegment(sessionId));
  }

  documentDir(sessionId: string, documentId: string): string {
    return this.storage.join(this.sessionDir(sessionId), "documents", safeSegment(documentId));
  }

  artifactPath(
    sessionId: string,
    documentId: string,
    kind: "original" | "working",
    format: StoredOfficeFormat
  ): string {
    return this.storage.join(this.documentDir(sessionId, documentId), "artifacts", `${kind}.${format}`);
  }

  async putSession(
    record: Omit<StoredSessionRecord, "schema" | "schemaVersion" | "version" | "lease">
  ): Promise<void> {
    await this.init();
    const full: StoredSessionRecord = {
      schema: "office-ai/session-record@1",
      schemaVersion: SESSION_RECORD_SCHEMA_VERSION,
      version: 1,
      ...record,
      lease: {
        pid: process.pid,
        host: "local",
        updatedAt: new Date().toISOString(),
      },
    };
    await atomicWriteJson(this.storage, this.sessionJsonPath(record.id), full);
  }

  async putDocument(
    record: Omit<StoredDocumentRecord, "schema" | "schemaVersion" | "version" | "artifacts">,
    opts: {
      readonly originalBytes?: Uint8Array | Buffer;
      readonly originalSourcePath?: string;
      readonly workingBytes?: Uint8Array | Buffer;
    } = {}
  ): Promise<StoredDocumentRecord> {
    await this.init();
    await this.storage.ensureDir(
      this.storage.join(this.documentDir(record.sessionId, record.id), "artifacts")
    );

    const existing = await this.getDocument(record.id).catch((err: unknown) => {
      if (err instanceof SessionStoreCorruptError) throw err;
      return null;
    });
    const originalPath =
      existing?.artifacts.originalPath ??
      (opts.originalBytes || opts.originalSourcePath
        ? this.artifactPath(record.sessionId, record.id, "original", record.format)
        : undefined);
    const workingPath =
      existing?.artifacts.workingPath ??
      (opts.workingBytes
        ? this.artifactPath(record.sessionId, record.id, "working", record.format)
        : undefined);

    if (opts.originalBytes && originalPath) {
      await this.storage.writeBytesAtomic(originalPath, opts.originalBytes);
    } else if (opts.originalSourcePath && originalPath && !(await this.storage.exists(originalPath))) {
      await this.storage.copyFromLocalFile(opts.originalSourcePath, originalPath);
    }
    if (opts.workingBytes && workingPath) {
      await this.storage.writeBytesAtomic(workingPath, opts.workingBytes);
    }

    const full: StoredDocumentRecord = {
      schema: "office-ai/document-record@1",
      schemaVersion: DOCUMENT_RECORD_SCHEMA_VERSION,
      version: 1,
      ...record,
      artifacts: {
        ...(originalPath ? { originalPath } : {}),
        ...(workingPath ? { workingPath } : {}),
      },
    };
    await atomicWriteJson(this.storage, this.documentJsonPath(record.sessionId, record.id), full);
    return full;
  }

  async listSessions(): Promise<ReadonlyArray<StoredSessionRecord>> {
    await this.init();
    const entries = await this.storage.list(this.sessionsDir());
    const out: StoredSessionRecord[] = [];
    for (const id of entries) {
      const path = this.sessionJsonPath(id);
      if (!(await this.storage.exists(path))) continue;
      out.push(await readStoredJson<StoredSessionRecord>(this.storage, path, "office-ai/session-record@1"));
    }
    out.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
    return out;
  }

  async listDocuments(sessionId?: string): Promise<ReadonlyArray<StoredDocumentRecord>> {
    const sessions = sessionId ? [await this.getSession(sessionId)] : await this.listSessions();
    const out: StoredDocumentRecord[] = [];
    for (const session of sessions) {
      for (const documentId of session.documentIds) {
        const doc = await this.getDocument(documentId);
        if (doc) out.push(doc);
      }
    }
    out.sort((a, b) => b.updatedAt.localeCompare(a.updatedAt));
    return out;
  }

  async getSession(sessionId: string): Promise<StoredSessionRecord> {
    return readStoredJson<StoredSessionRecord>(
      this.storage,
      this.sessionJsonPath(sessionId),
      "office-ai/session-record@1"
    );
  }

  async getDocument(documentId: string): Promise<StoredDocumentRecord | null> {
    const sessions = await this.listSessions();
    for (const session of sessions) {
      if (!session.documentIds.includes(documentId)) continue;
      try {
        return await readStoredJson<StoredDocumentRecord>(
          this.storage,
          this.documentJsonPath(session.id, documentId),
          "office-ai/document-record@1"
        );
      } catch (err) {
        if (isEnoent(err)) return null;
        throw err;
      }
    }
    return null;
  }

  async readWorkingBytes(record: StoredDocumentRecord): Promise<Uint8Array> {
    const path = record.artifacts.workingPath ?? record.artifacts.originalPath;
    if (!path) throw new Error(`Document ${record.id} has no persisted artifact.`);
    return this.storage.readBytes(path);
  }

  async inspectDataDir(): Promise<SessionStoreInspection> {
    const inspectedAt = new Date().toISOString();
    const diagnostics: StoredDiagnostic[] = [];
    const sessionEntries = await this.storage.list(this.sessionsDir());
    for (const sessionId of sessionEntries) {
      const sessionPath = this.sessionJsonPath(sessionId);
      if (!(await this.storage.exists(sessionPath))) continue;
      const parsed = await readLooseJson(this.storage, sessionPath).catch((err: unknown) => {
        diagnostics.push({
          level: "error",
          code: "session-store-corrupt-metadata",
          message: `Cannot read session metadata ${sessionPath}: ${
            err instanceof Error ? err.message : String(err)
          }`,
        });
        return null;
      });
      if (!parsed) continue;
      if (needsSessionMigration(parsed)) {
        diagnostics.push({
          level: "warning",
          code: "session-migration-required",
          message: `Session metadata ${sessionPath} needs migration to schema version ${SESSION_RECORD_SCHEMA_VERSION}.`,
        });
      }

      const documentEntries = await this.storage.list(
        this.storage.join(this.sessionDir(sessionId), "documents")
      );
      for (const documentId of documentEntries) {
        const documentPath = this.documentJsonPath(sessionId, documentId);
        if (!(await this.storage.exists(documentPath))) continue;
        const document = await readLooseJson(this.storage, documentPath).catch((err: unknown) => {
          diagnostics.push({
            level: "error",
            code: "session-store-corrupt-metadata",
            message: `Cannot read document metadata ${documentPath}: ${
              err instanceof Error ? err.message : String(err)
            }`,
          });
          return null;
        });
        if (document && needsDocumentMigration(document)) {
          diagnostics.push({
            level: "warning",
            code: "document-migration-required",
            message: `Document metadata ${documentPath} needs migration to schema version ${DOCUMENT_RECORD_SCHEMA_VERSION}.`,
          });
        }
      }
    }
    return {
      schema: "office-ai/session-store-inspection@1",
      dataDir: this.dataDir,
      schemaVersion: DATA_DIR_SCHEMA_VERSION,
      inspectedAt,
      needsMigration: diagnostics.some((diagnostic) => diagnostic.code.endsWith("migration-required")),
      diagnostics,
    };
  }

  async migrateDataDir(): Promise<SessionStoreMigrationResult> {
    const migratedAt = new Date().toISOString();
    await this.init();
    const backupRoot = this.storage.join(
      this.dataDir,
      "backups",
      `migration-${migratedAt.replace(/[^0-9A-Za-z_.-]/g, "_")}`
    );
    const migrations: SessionStoreMigrationRecord[] = [];
    const diagnostics: StoredDiagnostic[] = [];
    const sessionEntries = await this.storage.list(this.sessionsDir());

    for (const sessionId of sessionEntries) {
      const sessionPath = this.sessionJsonPath(sessionId);
      if (!(await this.storage.exists(sessionPath))) continue;
      const session = await readLooseJson(this.storage, sessionPath);
      if (needsSessionMigration(session)) {
        const backupPath = this.storage.join(backupRoot, "sessions", safeSegment(sessionId), "session.json");
        await copyStoredFile(this.storage, sessionPath, backupPath);
        await atomicWriteJson(this.storage, sessionPath, migrateSessionRecord(session));
        migrations.push({
          kind: "session",
          id: sessionId,
          path: sessionPath,
          backupPath,
          fromSchema: objectString(session, "schema"),
          fromVersion: objectNumber(session, "version"),
          toSchema: "office-ai/session-record@1",
          toVersion: SESSION_RECORD_SCHEMA_VERSION,
        });
      }

      const documentEntries = await this.storage.list(
        this.storage.join(this.sessionDir(sessionId), "documents")
      );
      for (const documentId of documentEntries) {
        const documentPath = this.documentJsonPath(sessionId, documentId);
        if (!(await this.storage.exists(documentPath))) continue;
        const document = await readLooseJson(this.storage, documentPath);
        if (!needsDocumentMigration(document)) continue;
        const backupPath = this.storage.join(
          backupRoot,
          "sessions",
          safeSegment(sessionId),
          "documents",
          safeSegment(documentId),
          "document.json"
        );
        await copyStoredFile(this.storage, documentPath, backupPath);
        await atomicWriteJson(this.storage, documentPath, migrateDocumentRecord(document));
        migrations.push({
          kind: "document",
          id: documentId,
          path: documentPath,
          backupPath,
          fromSchema: objectString(document, "schema"),
          fromVersion: objectNumber(document, "version"),
          toSchema: "office-ai/document-record@1",
          toVersion: DOCUMENT_RECORD_SCHEMA_VERSION,
        });
      }
    }

    diagnostics.push({
      level: "info",
      code: migrations.length > 0 ? "session-store-migrated" : "session-store-current",
      message:
        migrations.length > 0
          ? `Migrated ${migrations.length} session-store metadata file(s).`
          : "Session store metadata is already current.",
    });
    return {
      schema: "office-ai/session-store-migration@1",
      dataDir: this.dataDir,
      migratedAt,
      migrations,
      diagnostics,
    };
  }

  async cleanupTemporaryArtifacts(): Promise<SessionStoreCleanupResult> {
    const cleanedAt = new Date().toISOString();
    const removed: string[] = [];
    const diagnostics: StoredDiagnostic[] = [];
    await this.cleanupTemporaryFilesInDir(this.dataDir, removed);
    const sessionEntries = await this.storage.list(this.sessionsDir());
    for (const sessionId of sessionEntries) {
      const sessionDir = this.sessionDir(sessionId);
      await this.cleanupTemporaryFilesInDir(sessionDir, removed);
      const documentsDir = this.storage.join(sessionDir, "documents");
      for (const documentId of await this.storage.list(documentsDir)) {
        const documentDir = this.documentDir(sessionId, documentId);
        await this.cleanupTemporaryFilesInDir(documentDir, removed);
        await this.cleanupTemporaryFilesInDir(this.storage.join(documentDir, "artifacts"), removed);
      }
    }
    diagnostics.push({
      level: "info",
      code: "session-store-cleanup",
      message: `Removed ${removed.length} temporary session-store file(s); original and working artifacts are preserved.`,
    });
    return {
      schema: "office-ai/session-store-cleanup@1",
      dataDir: this.dataDir,
      cleanedAt,
      removed,
      diagnostics,
    };
  }

  async clear(): Promise<void> {
    await this.storage.remove(this.dataDir, { recursive: true, force: true });
  }

  private sessionsDir(): string {
    return this.storage.join(this.dataDir, "sessions");
  }

  private sessionJsonPath(sessionId: string): string {
    return this.storage.join(this.sessionDir(sessionId), "session.json");
  }

  private documentJsonPath(sessionId: string, documentId: string): string {
    return this.storage.join(this.documentDir(sessionId, documentId), "document.json");
  }

  private async cleanupTemporaryFilesInDir(dir: string, removed: string[]): Promise<void> {
    for (const entry of await this.storage.list(dir)) {
      if (!isTemporaryStoreFile(entry)) continue;
      const path = this.storage.join(dir, entry);
      await this.storage.remove(path, { force: true });
      removed.push(path);
    }
  }
}

export function createLocalSessionStore(opts: LocalSessionStoreOptions = {}): LocalSessionStore {
  return new LocalSessionStore(opts);
}

export function createLocalFilesystemSessionStorageAdapter(
  dataDir?: string
): LocalFilesystemSessionStorageAdapter {
  return new LocalFilesystemSessionStorageAdapter(resolveOfficeAiDataDir(dataDir));
}

export function resolveOfficeAiDataDir(explicit?: string): string {
  if (explicit && explicit.trim().length > 0) return resolve(explicit);
  const env = process.env.OFFICEAI_DATA_DIR;
  if (env && env.trim().length > 0) return resolve(env);
  const xdg = process.env.XDG_DATA_HOME;
  if (xdg && xdg.trim().length > 0) return resolve(xdg, "office-ai");
  return resolve(homedir(), ".local", "share", "office-ai");
}

async function readStoredJson<T extends { readonly schema?: string }>(
  storage: SessionStorageAdapter,
  path: string,
  schema: string
): Promise<T> {
  let parsed: unknown;
  try {
    parsed = JSON.parse(Buffer.from(await storage.readBytes(path)).toString("utf8"));
  } catch (err) {
    if (isEnoent(err)) throw err;
    if (err instanceof SessionStoreStorageError) throw err;
    throw new SessionStoreCorruptError(path, "Corrupt metadata file", err);
  }
  if (!parsed || typeof parsed !== "object" || (parsed as { schema?: unknown }).schema !== schema) {
    throw new SessionStoreCorruptError(path, `Metadata schema mismatch, expected ${schema}`);
  }
  return parsed as T;
}

async function readLooseJson(storage: SessionStorageAdapter, path: string): Promise<unknown> {
  try {
    return JSON.parse(Buffer.from(await storage.readBytes(path)).toString("utf8"));
  } catch (err) {
    if (isEnoent(err)) throw err;
    if (err instanceof SessionStoreStorageError) throw err;
    throw new SessionStoreCorruptError(path, "Corrupt metadata file", err);
  }
}

async function copyStoredFile(
  storage: SessionStorageAdapter,
  sourcePath: string,
  targetPath: string
): Promise<void> {
  await storage.writeBytesAtomic(targetPath, await storage.readBytes(sourcePath));
}

function needsSessionMigration(value: unknown): boolean {
  if (!isRecord(value)) return false;
  if (
    value.schema === "office-ai/session-record@1" &&
    value.schemaVersion === SESSION_RECORD_SCHEMA_VERSION &&
    value.version === 1
  ) {
    return false;
  }
  return isSessionLike(value);
}

function needsDocumentMigration(value: unknown): boolean {
  if (!isRecord(value)) return false;
  if (
    value.schema === "office-ai/document-record@1" &&
    value.schemaVersion === DOCUMENT_RECORD_SCHEMA_VERSION &&
    value.version === 1
  ) {
    return false;
  }
  return isDocumentLike(value);
}

function migrateSessionRecord(value: unknown): StoredSessionRecord {
  if (!isRecord(value) || !isSessionLike(value)) {
    throw new SessionStoreCorruptError("<memory>", "Cannot migrate invalid session metadata");
  }
  const now = new Date().toISOString();
  return {
    schema: "office-ai/session-record@1",
    schemaVersion: SESSION_RECORD_SCHEMA_VERSION,
    version: 1,
    id: String(value.id),
    title: String(value.title),
    createdAt: String(value.createdAt),
    updatedAt: String(value.updatedAt),
    documentIds: Array.isArray(value.documentIds) ? value.documentIds.map(String) : [],
    lease: isRecord(value.lease)
      ? {
          pid: typeof value.lease.pid === "number" ? value.lease.pid : process.pid,
          host: typeof value.lease.host === "string" ? value.lease.host : "local",
          updatedAt: typeof value.lease.updatedAt === "string" ? value.lease.updatedAt : now,
        }
      : {
          pid: process.pid,
          host: "local",
          updatedAt: now,
        },
  };
}

function migrateDocumentRecord(value: unknown): StoredDocumentRecord {
  if (!isRecord(value) || !isDocumentLike(value)) {
    throw new SessionStoreCorruptError("<memory>", "Cannot migrate invalid document metadata");
  }
  const artifacts = isRecord(value.artifacts) ? value.artifacts : {};
  return {
    schema: "office-ai/document-record@1",
    schemaVersion: DOCUMENT_RECORD_SCHEMA_VERSION,
    version: 1,
    id: String(value.id),
    sessionId: String(value.sessionId),
    format: value.format as StoredOfficeFormat,
    name: String(value.name),
    status: value.status === "error" ? "error" : "ready",
    ...(typeof value.sourcePath === "string" ? { sourcePath: value.sourcePath } : {}),
    createdAt: String(value.createdAt),
    updatedAt: String(value.updatedAt),
    revision: typeof value.revision === "number" ? value.revision : 0,
    diagnostics: Array.isArray(value.diagnostics) ? (value.diagnostics as StoredDiagnostic[]) : [],
    exportHistory: Array.isArray(value.exportHistory) ? (value.exportHistory as StoredExportRecord[]) : [],
    artifacts: {
      ...(typeof artifacts.originalPath === "string" ? { originalPath: artifacts.originalPath } : {}),
      ...(typeof artifacts.workingPath === "string" ? { workingPath: artifacts.workingPath } : {}),
    },
    pendingChanges: Array.isArray(value.pendingChanges)
      ? (value.pendingChanges as StoredPendingChange[])
      : [],
    commandLog: Array.isArray(value.commandLog) ? (value.commandLog as StoredCommandLogEntry[]) : [],
  };
}

function isSessionLike(value: Record<string, unknown>): boolean {
  return (
    typeof value.id === "string" &&
    typeof value.title === "string" &&
    typeof value.createdAt === "string" &&
    typeof value.updatedAt === "string" &&
    Array.isArray(value.documentIds)
  );
}

function isDocumentLike(value: Record<string, unknown>): boolean {
  return (
    typeof value.id === "string" &&
    typeof value.sessionId === "string" &&
    isStoredOfficeFormat(value.format) &&
    typeof value.name === "string" &&
    typeof value.createdAt === "string" &&
    typeof value.updatedAt === "string"
  );
}

function isStoredOfficeFormat(value: unknown): value is StoredOfficeFormat {
  return value === "docx" || value === "xlsx" || value === "pptx" || value === "pdf";
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value) && typeof value === "object";
}

function objectString(value: unknown, key: string): string | undefined {
  return isRecord(value) && typeof value[key] === "string" ? value[key] : undefined;
}

function objectNumber(value: unknown, key: string): number | undefined {
  return isRecord(value) && typeof value[key] === "number" ? value[key] : undefined;
}

function isTemporaryStoreFile(name: string): boolean {
  return name.startsWith(".") && name.endsWith(".tmp");
}

function isEnoent(err: unknown): boolean {
  return (
    (err as NodeJS.ErrnoException)?.code === "ENOENT" ||
    (err instanceof SessionStoreStorageError && err.causeCode === "ENOENT")
  );
}

async function atomicWriteJson(storage: SessionStorageAdapter, path: string, value: unknown): Promise<void> {
  await storage.writeBytesAtomic(path, Buffer.from(`${JSON.stringify(value, null, 2)}\n`, "utf8"));
}

function safeSegment(value: string): string {
  const cleaned = value.replace(/[^a-zA-Z0-9_.-]/g, "_");
  if (cleaned.length === 0 || cleaned === "." || cleaned === "..") {
    throw new Error(`Invalid data-dir segment "${value}".`);
  }
  return cleaned;
}
