import type { WebDiagnosticEntry, WebOfficeFormat } from "./web-sessions";

export type WebParityStatus = "full" | "partial" | "review-only" | "planned";

export interface WebFormatParityRow {
  readonly label: string;
  readonly status: WebParityStatus;
  readonly detail: string;
}

export interface WebFormatParity {
  readonly format: WebOfficeFormat;
  readonly title: string;
  readonly rows: ReadonlyArray<WebFormatParityRow>;
  readonly knownLimits: ReadonlyArray<string>;
  readonly diagnostic: WebDiagnosticEntry;
}

export const WEB_FORMAT_PARITY: Record<WebOfficeFormat, WebFormatParity> = {
  docx: {
    format: "docx",
    title: "DOCX web editing",
    rows: [
      { label: "Import", status: "full", detail: "Session import/create/export is live." },
      { label: "Read view", status: "full", detail: "Markdown, text, page and JSON projections are live." },
      {
        label: "Editable structures",
        status: "partial",
        detail: "Text, styles, comments and many structural commands are reviewable.",
      },
      {
        label: "Review and diff",
        status: "partial",
        detail: "Pending changes, semantic diff summary and audit activity are visible.",
      },
      { label: "Export", status: "full", detail: "Working artifact exports back to DOCX." },
    ],
    knownLimits: ["Manual web controls do not expose every table/style/header/footer command yet."],
    diagnostic: {
      level: "info",
      code: "web-parity-docx-partial-edit",
      message:
        "DOCX has full import/read/export and partial manual editing; unsupported structures stay review/roadmap items.",
    },
  },
  xlsx: {
    format: "xlsx",
    title: "XLSX web editing",
    rows: [
      { label: "Import", status: "full", detail: "Session import/create/export is live." },
      { label: "Read view", status: "full", detail: "Sheet/range markdown and JSON projections are live." },
      {
        label: "Editable structures",
        status: "partial",
        detail: "Cells, formatting, sheets and selected data operations are supported.",
      },
      {
        label: "Review and diff",
        status: "partial",
        detail: "Agent changes can be reviewed through pending changes and activity.",
      },
      { label: "Export", status: "full", detail: "Working artifact exports back to XLSX." },
    ],
    knownLimits: [
      "Charts, advanced tables and all formula-management controls are not fully manual-editable in the web UI.",
    ],
    diagnostic: {
      level: "info",
      code: "web-parity-xlsx-partial-edit",
      message:
        "XLSX supports import/read/export and partial sheet editing; charts and advanced structures are roadmap-visible.",
    },
  },
  pptx: {
    format: "pptx",
    title: "PPTX web editing",
    rows: [
      { label: "Import", status: "full", detail: "Session import/create/export is live." },
      {
        label: "Read view",
        status: "full",
        detail: "Slide markdown, JSON and page-style projections are live.",
      },
      {
        label: "Editable structures",
        status: "partial",
        detail: "Slides, text, shapes and selected layout commands are supported.",
      },
      {
        label: "Review and diff",
        status: "partial",
        detail: "Agent changes can be reviewed through pending changes and activity.",
      },
      { label: "Export", status: "full", detail: "Working artifact exports back to PPTX." },
    ],
    knownLimits: [
      "Assets, complex layouts, charts and animations are not fully manual-editable in the web UI.",
    ],
    diagnostic: {
      level: "info",
      code: "web-parity-pptx-partial-edit",
      message:
        "PPTX supports import/read/export and partial slide editing; complex assets/layouts remain explicit roadmap items.",
    },
  },
  pdf: {
    format: "pdf",
    title: "PDF web review",
    rows: [
      { label: "Import", status: "full", detail: "Session import/create/export is live." },
      { label: "Read view", status: "full", detail: "Page, text, metadata and JSON projections are live." },
      {
        label: "Editable structures",
        status: "review-only",
        detail: "PDF is currently review/annotation-oriented in the web product.",
      },
      {
        label: "Review and diff",
        status: "partial",
        detail: "Pending changes, diagnostics and activity are visible.",
      },
      { label: "Export", status: "full", detail: "Working artifact exports back to PDF." },
    ],
    knownLimits: [
      "Manual PDF annotations, highlights and text-layer edits are planned rather than implied as available.",
    ],
    diagnostic: {
      level: "warning",
      code: "web-parity-pdf-review-only",
      message:
        "PDF is web-readable and exportable; annotation/highlight editing is planned and not silently presented as available.",
    },
  },
  email: {
    format: "email",
    title: "Email web viewer",
    rows: [
      { label: "Import", status: "full", detail: ".eml and .msg files import into sessions." },
      {
        label: "Read view",
        status: "full",
        detail: "Headers, text/html body and attachments are normalized.",
      },
      {
        label: "Editable structures",
        status: "review-only",
        detail: "Messages are viewer-only; attachments can route into their own editors.",
      },
      {
        label: "Review and diff",
        status: "planned",
        detail: "Message-level semantic changes are intentionally not presented as editable.",
      },
      { label: "Export", status: "full", detail: "The original working message artifact exports back." },
    ],
    knownLimits: [
      "MSG parsing extracts normalized headers/body text when possible and preserves the source artifact for export.",
    ],
    diagnostic: {
      level: "info",
      code: "web-parity-email-viewer",
      message:
        "Email files are session-backed, readable and attachment-routable; the message body is viewer-only.",
    },
  },
  image: {
    format: "image",
    title: "Image web viewer",
    rows: [
      {
        label: "Import",
        status: "full",
        detail: "PNG, JPEG, WebP, GIF, SVG, BMP, TIFF and HEIC imports are live.",
      },
      {
        label: "Read view",
        status: "full",
        detail: "Browser-safe formats render directly; HEIC/TIFF get a normalized preview state.",
      },
      {
        label: "Editable structures",
        status: "review-only",
        detail: "Image files are viewer/export artifacts rather than canvas-editable documents.",
      },
      {
        label: "Review and diff",
        status: "planned",
        detail: "Pixel-level diff and annotation review are not exposed as document commands yet.",
      },
      { label: "Export", status: "full", detail: "The original working image artifact exports back." },
    ],
    knownLimits: [
      "HEIC/HEIF/TIFF decoding is lazy and represented with metadata when the browser cannot render the codec.",
    ],
    diagnostic: {
      level: "info",
      code: "web-parity-image-viewer",
      message:
        "Images are session-backed, readable and exportable; unsupported browser codecs get a normalized preview.",
    },
  },
};

export function formatParityFor(format: WebOfficeFormat): WebFormatParity {
  return WEB_FORMAT_PARITY[format];
}

export function formatParityDiagnostics(format: WebOfficeFormat): ReadonlyArray<WebDiagnosticEntry> {
  return [WEB_FORMAT_PARITY[format].diagnostic];
}
