export { runCli } from "./cli.js";
export {
  parseSelector,
  type Selector,
  type ParagraphSelector,
  type ParagraphRangeSelector,
} from "./selector.js";
export {
  createLocalSessionStore,
  createLocalFilesystemSessionStorageAdapter,
  LocalFilesystemSessionStorageAdapter,
  resolveOfficeAiDataDir,
  SessionStoreCorruptError,
  SessionStoreStorageError,
  type LocalSessionStoreOptions,
  type SessionStorageAdapter,
  type SessionStorageCapabilities,
  type SessionStorageOperation,
  type SessionStorageRemoveOptions,
  type StoredCommandLogEntry,
  type StoredDiagnostic,
  type StoredDocumentRecord,
  type StoredExportRecord,
  type StoredOfficeFormat,
  type StoredPendingChange,
  type StoredSessionRecord,
} from "./session-store.js";
export {
  projectOfficeDocument,
  projectionDocumentEnvelope,
  type ProjectionDocumentMeta,
  type ProjectionFormat,
  type ProjectionKind,
  type ProjectionOptions,
  type ProjectionSource,
} from "./projections.js";
