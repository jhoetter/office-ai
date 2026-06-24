export { runCli } from "./cli.js";
export {
  parseSelector,
  type Selector,
  type ParagraphSelector,
  type ParagraphRangeSelector,
} from "./selector.js";
export {
  createLocalSessionStore,
  resolveOfficeAiDataDir,
  SessionStoreCorruptError,
  type LocalSessionStoreOptions,
  type StoredCommandLogEntry,
  type StoredDiagnostic,
  type StoredDocumentRecord,
  type StoredExportRecord,
  type StoredOfficeFormat,
  type StoredPendingChange,
  type StoredSessionRecord,
} from "./session-store.js";
