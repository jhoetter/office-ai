import type { WebOfficeFormat } from "./web-sessions";

export interface SessionDocumentRouteTarget {
  readonly documentId: string;
  readonly format: WebOfficeFormat;
}

const EDITOR_PATH_BY_FORMAT: Record<WebOfficeFormat, string> = {
  docx: "/editor",
  xlsx: "/xlsx-editor",
  pptx: "/pptx-editor",
  pdf: "/pdf-viewer",
  image: "/image-viewer",
};

export function editorPathForFormat(format: WebOfficeFormat): string {
  return EDITOR_PATH_BY_FORMAT[format];
}

export function editorHrefForSessionDocument(target: SessionDocumentRouteTarget): string {
  const params = new URLSearchParams({ session: target.documentId });
  return `${editorPathForFormat(target.format)}?${params.toString()}`;
}

export function inspectorHrefForDocumentId(documentId: string): string {
  return `/sessions/${encodeURIComponent(documentId)}`;
}

export function sampleHrefForFormat(args: {
  readonly format: WebOfficeFormat;
  readonly url: string;
  readonly name: string;
}): string {
  const params = new URLSearchParams({ src: args.url, name: args.name });
  return `${editorPathForFormat(args.format)}?${params.toString()}`;
}
