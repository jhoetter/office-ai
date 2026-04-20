import type { AnnotationInput } from "./types.js";

/**
 * XFDF JSON projection. We ship a pragmatic, lossless JSON form for
 * collaborative round-trip. The binary FDF format and the full XFDF
 * XML serializer are follow-ups; the import/export here is the LLM
 * agent surface.
 */
export interface XfdfDocument {
  readonly version: 1;
  readonly annotations: ReadonlyArray<AnnotationInput>;
}

export const exportXfdf = (annotations: ReadonlyArray<AnnotationInput>): XfdfDocument => ({
  version: 1,
  annotations,
});

export const importXfdf = (doc: XfdfDocument): ReadonlyArray<AnnotationInput> => {
  if (doc.version !== 1) {
    throw new Error(`pdf-annotations/xfdf: unsupported version ${(doc as { version: unknown }).version}`);
  }
  return doc.annotations;
};
