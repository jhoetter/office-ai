/**
 * Top-level entry point for `@officeai/react-editors`.
 *
 * Re-exports the always-useful subset (mime constants, props
 * contract, blank-file builders) so simple hosts can `import { … }
 * from "@officeai/react-editors"` without learning the subpath
 * surface. Hosts that care about bundle size should prefer the
 * subpaths (`./blanks`, `./mime`, `./contract`) for tighter
 * tree-shaking — the subpath consumers don't pull in the other
 * format's agent code at all.
 */
export {
  DOCX_MIME,
  XLSX_MIME,
  PPTX_MIME,
  PDF_MIME,
  DEFAULT_BLANK_FILENAME,
  MIME_BY_FORMAT,
  detectFormatFromFilename,
  type OfficeFormat,
  type OfficeMime,
} from "./mime.js";

export { makeBlankDocx, makeBlankXlsx, makeBlankPptx, makeBlankPdf, makeBlank } from "./blanks/index.js";

export type { EmbeddedEditorProps, EmbeddedEditorOnSave, Locale, Theme } from "./contract.js";
