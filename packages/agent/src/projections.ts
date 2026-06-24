import { DocxAgent } from "@officeai/docx";
import { XlsxAgent } from "@officeai/xlsx";
import { PptxAgent } from "@officeai/pptx";
import { PdfAgent } from "@officeai/pdf";

export type ProjectionFormat = "docx" | "xlsx" | "pptx" | "pdf";
export type ProjectionKind = "summary" | "markdown" | "json" | "text" | "page";

export interface ProjectionDocumentMeta {
  readonly documentId: string;
  readonly sessionId: string;
  readonly format: ProjectionFormat;
  readonly name: string;
  readonly status: "ready" | "error";
  readonly revision: number;
  readonly sourcePath?: string;
  readonly diagnostics: ReadonlyArray<unknown>;
  readonly exportHistory: ReadonlyArray<unknown>;
  readonly createdAt: string;
  readonly updatedAt: string;
}

export type ProjectionSource =
  | { readonly format: "docx"; readonly agent: DocxAgent }
  | { readonly format: "xlsx"; readonly agent: XlsxAgent }
  | { readonly format: "pptx"; readonly agent: PptxAgent }
  | { readonly format: "pdf"; readonly agent: PdfAgent };

export interface ProjectionOptions {
  readonly projection: ProjectionKind;
  readonly page?: number;
  readonly sheet?: string;
  readonly range?: string;
  readonly slide?: number;
  readonly maxRows?: number;
  readonly maxCols?: number;
}

export function projectOfficeDocument(
  meta: ProjectionDocumentMeta,
  source: ProjectionSource,
  opts: ProjectionOptions
): Record<string, unknown> {
  const document = projectionDocumentEnvelope(meta);
  const base = {
    ...document,
    schema: "office-ai/document-projection@1",
    document,
  };
  if (opts.projection === "summary") {
    return { ...base, projection: "summary", summary: summaryFor(source) };
  }

  switch (source.format) {
    case "docx":
      return projectDocx(base, source.agent, opts);
    case "xlsx":
      return projectXlsx(base, source.agent, opts);
    case "pptx":
      return projectPptx(base, source.agent, opts);
    case "pdf":
      return projectPdf(base, source.agent, opts);
  }
}

export function projectionDocumentEnvelope(meta: ProjectionDocumentMeta): Record<string, unknown> {
  return {
    schema: "office-ai/document@1",
    documentId: meta.documentId,
    sessionId: meta.sessionId,
    format: meta.format,
    name: meta.name,
    status: meta.status,
    revision: meta.revision,
    ...(meta.sourcePath ? { sourcePath: meta.sourcePath } : {}),
    diagnostics: meta.diagnostics,
    exportHistory: meta.exportHistory,
    createdAt: meta.createdAt,
    updatedAt: meta.updatedAt,
  };
}

function summaryFor(source: ProjectionSource): unknown {
  switch (source.format) {
    case "docx":
      return {
        format: "docx",
        revision: source.agent.getSnapshot().revision,
        blocks: source.agent.getSnapshot().root.body.length,
        pages: source.agent.getPages().length,
      };
    case "xlsx":
      return {
        format: "xlsx",
        revision: source.agent.getSnapshot().revision,
        sheets: source.agent.listSheets(),
      };
    case "pptx": {
      const snap = source.agent.getSnapshot();
      return {
        format: "pptx",
        revision: snap.revision,
        slides: snap.root.slides.length,
      };
    }
    case "pdf":
      return {
        format: "pdf",
        revision: source.agent.getSnapshot().revision,
        pages: source.agent.getSnapshot().root.pages.length,
      };
  }
}

function projectDocx(
  base: Record<string, unknown>,
  agent: DocxAgent,
  opts: ProjectionOptions
): Record<string, unknown> {
  if (opts.projection === "json") {
    return { ...base, projection: "json", content: agent.getSnapshot() };
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

function projectXlsx(
  base: Record<string, unknown>,
  agent: XlsxAgent,
  opts: ProjectionOptions
): Record<string, unknown> {
  if (opts.projection === "json") {
    return {
      ...base,
      projection: "json",
      content: opts.sheet
        ? agent.getRange(
            opts.range
              ? { kind: "xlsx-range", sheet: opts.sheet, range: opts.range }
              : { kind: "xlsx-sheet", sheet: opts.sheet }
          )
        : agent.getSnapshot(),
    };
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

function projectPptx(
  base: Record<string, unknown>,
  agent: PptxAgent,
  opts: ProjectionOptions
): Record<string, unknown> {
  if (opts.projection === "json" || opts.projection === "page") {
    return {
      ...base,
      projection: opts.projection,
      content:
        opts.slide !== undefined
          ? agent.getRange({ kind: "pptx-slides", start: opts.slide, end: opts.slide + 1 })
          : agent.getSnapshot(),
    };
  }
  return {
    ...base,
    projection: "markdown",
    content:
      opts.slide !== undefined
        ? agent.getRange({ kind: "pptx-slides", start: opts.slide, end: opts.slide + 1 })
        : agent.toMarkdown(),
  };
}

function projectPdf(
  base: Record<string, unknown>,
  agent: PdfAgent,
  opts: ProjectionOptions
): Record<string, unknown> {
  if (opts.projection === "json") {
    return { ...base, projection: "json", content: agent.getSnapshot() };
  }
  if (opts.projection === "page") {
    const page = opts.page ?? 1;
    return {
      ...base,
      projection: "page",
      page,
      content: agent.getRange({ kind: "pdf-pages", start: page, end: page }),
    };
  }
  return { ...base, projection: "markdown", content: agent.toMarkdown() };
}

function docxPlainText(agent: DocxAgent): string {
  const lines: string[] = [];
  for (const b of agent.getSnapshot().root.body) {
    if (b.kind !== "paragraph") continue;
    lines.push(
      b.children
        .map((c) =>
          c.kind === "run" ? c.children.map((g) => (g.kind === "text" ? g.text : "")).join("") : ""
        )
        .join("")
    );
  }
  return lines.join("\n");
}
