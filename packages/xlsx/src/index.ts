/**
 * @officeai/xlsx — XLSX editor entry point.
 *
 * Phase 4 surface (this commit):
 *   - parseXlsx / serializeXlsx (byte-preserving round-trip)
 *   - thin model: XlsxSnapshot, XlsxWorkbook, Sheet, OpaquePart
 *
 * Surfaces still to ship:
 *   - Phase 5 — typed cell layer + 13 P0 command handlers
 *   - Phase 6 — XlsxAgent (DocumentAgent contract)
 *   - Phase 7 — formula engine
 *   - Phase 9 — virtualized renderer
 */

export const XLSX_PACKAGE_VERSION = "0.1.0";

export {
  parseXlsx,
  resolveTargetPath,
  XlsxParseError,
  type ParseOptions,
  type XlsxParseErrorCode,
} from "./parser/index.js";

export { serializeXlsx, XlsxSerializeError, type XlsxSerializeErrorCode } from "./serializer/index.js";

export {
  emptyDirty,
  type OpaquePart,
  type Sheet,
  type XlsxDirtyFlags,
  type XlsxSnapshot,
  type XlsxWorkbook,
} from "./model/index.js";
