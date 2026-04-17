/**
 * @officeai/xlsx — XLSX editor entry point.
 *
 * Surfaces are filled in over the XLSX phases (see `spec/xlsx/` and
 * `docs/build-log/xlsx.md`):
 *
 * - Phase 4 — `parseXlsx`, `serializeXlsx` (OOXML I/O with opaque-part preservation)
 * - Phase 5 — model types + command handlers
 * - Phase 6 — `XlsxAgent`
 * - Phase 7 — formula engine
 * - Phase 9 — virtualized renderer
 *
 * Until those land, the package exports types only so it can be referenced
 * from the architecture graph and the workspace dep map without breaking
 * downstream typechecks.
 */
export const XLSX_PACKAGE_VERSION = "0.1.0";
