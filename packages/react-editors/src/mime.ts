/**
 * Canonical MIME constants for the four office-ai-supported formats.
 *
 * Hosts use these when handing bytes back to their own upload
 * pipelines (e.g. embedding host's presigned object-storage PUT). Centralising them here
 * keeps the contract that both `@officeai/react-editors`'s editor
 * components and any host code agree on the same content-type strings.
 */

export const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document" as const;
export const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" as const;
export const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation" as const;
export const PDF_MIME = "application/pdf" as const;

export type OfficeMime = typeof DOCX_MIME | typeof XLSX_MIME | typeof PPTX_MIME | typeof PDF_MIME;

export type OfficeFormat = "docx" | "xlsx" | "pptx" | "pdf";

export const DEFAULT_BLANK_FILENAME: Readonly<Record<OfficeFormat, string>> = {
  docx: "Untitled.docx",
  xlsx: "Untitled.xlsx",
  pptx: "Untitled.pptx",
  pdf: "Untitled.pdf",
} as const;

export const MIME_BY_FORMAT: Readonly<Record<OfficeFormat, OfficeMime>> = {
  docx: DOCX_MIME,
  xlsx: XLSX_MIME,
  pptx: PPTX_MIME,
  pdf: PDF_MIME,
} as const;

const EXTENSION_BY_FORMAT: Readonly<Record<OfficeFormat, string>> = {
  docx: ".docx",
  xlsx: ".xlsx",
  pptx: ".pptx",
  pdf: ".pdf",
};

/**
 * Map a filename to its OfficeFormat by extension. Returns `null` for
 * anything not recognised so hosts can decide whether to fall back to
 * a plain download flow.
 */
export function detectFormatFromFilename(name: string): OfficeFormat | null {
  const lower = name.toLowerCase();
  for (const fmt of ["docx", "xlsx", "pptx", "pdf"] as const) {
    if (lower.endsWith(EXTENSION_BY_FORMAT[fmt])) return fmt;
  }
  return null;
}
