import type { DocumentFormat } from "../types/document.js";

export type OoxmlPreservationFormat = Extract<DocumentFormat, "docx" | "xlsx" | "pptx">;
export type OoxmlPreservationDiagnosticLevel = "info" | "warning" | "error";
export type OoxmlPreservationDiagnosticCode =
  | "ooxml-opaque-part-preserved"
  | "ooxml-opaque-preservation-risk"
  | "ooxml-opaque-mutation-blocked";

export interface OoxmlPreservationContract {
  readonly format: OoxmlPreservationFormat;
  readonly parsedPartPatterns: ReadonlyArray<string>;
  readonly opaquePartPatterns: ReadonlyArray<string>;
  readonly relationshipPolicy: string;
  readonly mutationPolicy: string;
  readonly diagnosticCodes: ReadonlyArray<OoxmlPreservationDiagnosticCode>;
}

export interface OoxmlPreservationDiagnostic {
  readonly level: OoxmlPreservationDiagnosticLevel;
  readonly code: OoxmlPreservationDiagnosticCode;
  readonly format: OoxmlPreservationFormat;
  readonly message: string;
  readonly partPath?: string;
}

const DIAGNOSTIC_LEVEL: Record<OoxmlPreservationDiagnosticCode, OoxmlPreservationDiagnosticLevel> = {
  "ooxml-opaque-part-preserved": "info",
  "ooxml-opaque-preservation-risk": "warning",
  "ooxml-opaque-mutation-blocked": "error",
};

export const OOXML_PRESERVATION_CONTRACT: Record<OoxmlPreservationFormat, OoxmlPreservationContract> = {
  docx: {
    format: "docx",
    parsedPartPatterns: [
      "word/document.xml",
      "word/styles.xml",
      "word/numbering.xml",
      "word/comments*.xml",
      "word/header*.xml",
      "word/footer*.xml",
      "word/footnotes.xml",
      "word/charts/chart*.xml",
    ],
    opaquePartPatterns: [
      "customXml/**",
      "word/embeddings/**",
      "word/media/** unless explicitly dirtied",
      "word/theme/**",
      "word/settings.xml unless explicitly dirtied",
      "unknown word/** package extensions",
    ],
    relationshipPolicy:
      "Untouched .rels parts are emitted byte-for-byte. Commands that add/remove a known relationship dirty only that owning .rels part and must keep unrelated relationship ids and targets.",
    mutationPolicy:
      "Commands may rewrite known typed parts only. Opaque parts are preserved; commands targeting opaque structures must emit a blocking diagnostic instead of dropping bytes.",
    diagnosticCodes: [
      "ooxml-opaque-part-preserved",
      "ooxml-opaque-preservation-risk",
      "ooxml-opaque-mutation-blocked",
    ],
  },
  xlsx: {
    format: "xlsx",
    parsedPartPatterns: [
      "xl/workbook.xml",
      "xl/worksheets/sheet*.xml",
      "xl/sharedStrings.xml",
      "xl/styles.xml",
      "xl/comments*.xml",
      "xl/tables/table*.xml",
      "xl/drawings/drawing*.xml",
      "xl/charts/chart*.xml",
    ],
    opaquePartPatterns: [
      "customXml/**",
      "xl/embeddings/**",
      "xl/pivotCache/**",
      "xl/pivotTables/** unless explicitly promoted to typed model",
      "xl/slicerCaches/**",
      "xl/timelines/**",
      "xl/vbaProject.bin",
      "unknown xl/** package extensions",
    ],
    relationshipPolicy:
      "Workbook, worksheet, drawing and chart .rels graphs preserve unrelated relationship ids and targets when a command dirties a known owner part.",
    mutationPolicy:
      "Cell, sheet, table, chart and drawing commands dirty only their owning known parts. Unsupported opaque workbook features are preserved or blocked with diagnostics.",
    diagnosticCodes: [
      "ooxml-opaque-part-preserved",
      "ooxml-opaque-preservation-risk",
      "ooxml-opaque-mutation-blocked",
    ],
  },
  pptx: {
    format: "pptx",
    parsedPartPatterns: [
      "ppt/presentation.xml",
      "ppt/slides/slide*.xml",
      "ppt/slideLayouts/slideLayout*.xml",
      "ppt/slideMasters/slideMaster*.xml",
      "ppt/theme/theme*.xml",
      "ppt/notesSlides/notesSlide*.xml",
      "ppt/charts/chart*.xml",
    ],
    opaquePartPatterns: [
      "customXml/**",
      "ppt/embeddings/**",
      "ppt/media/** unless explicitly dirtied",
      "ppt/tags/**",
      "ppt/viewProps.xml",
      "ppt/presProps.xml unless explicitly dirtied",
      "unknown ppt/** package extensions",
    ],
    relationshipPolicy:
      "Presentation, slide, layout, master, notes and chart .rels graphs preserve unrelated relationship ids and targets when a known part is rewritten.",
    mutationPolicy:
      "Slide and media commands may rewrite known slide/media parts. Unknown embedded objects, tags and package extensions stay opaque or produce blocking diagnostics.",
    diagnosticCodes: [
      "ooxml-opaque-part-preserved",
      "ooxml-opaque-preservation-risk",
      "ooxml-opaque-mutation-blocked",
    ],
  },
};

export function ooxmlPreservationDiagnostic(args: {
  readonly format: OoxmlPreservationFormat;
  readonly code: OoxmlPreservationDiagnosticCode;
  readonly partPath?: string;
  readonly detail?: string;
}): OoxmlPreservationDiagnostic {
  const level = DIAGNOSTIC_LEVEL[args.code];
  const subject = args.partPath ? `opaque OOXML part ${args.partPath}` : "opaque OOXML content";
  const prefix = `${args.format.toUpperCase()} ${subject}`;
  const detail = args.detail ? ` ${args.detail}` : "";
  switch (args.code) {
    case "ooxml-opaque-part-preserved":
      return {
        level,
        code: args.code,
        format: args.format,
        ...(args.partPath ? { partPath: args.partPath } : {}),
        message: `${prefix} was preserved byte-for-byte.${detail}`,
      };
    case "ooxml-opaque-preservation-risk":
      return {
        level,
        code: args.code,
        format: args.format,
        ...(args.partPath ? { partPath: args.partPath } : {}),
        message: `${prefix} may not be safe to rewrite.${detail}`,
      };
    case "ooxml-opaque-mutation-blocked":
      return {
        level,
        code: args.code,
        format: args.format,
        ...(args.partPath ? { partPath: args.partPath } : {}),
        message: `${prefix} cannot be safely mutated.${detail}`,
      };
  }
}
