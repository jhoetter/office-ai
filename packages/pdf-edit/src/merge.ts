import { PDFDocument } from "pdf-lib";

export interface MergePdfsOptions {
  readonly inputs: ReadonlyArray<Uint8Array>;
}

export const mergePdfs = async (opts: MergePdfsOptions): Promise<Uint8Array> => {
  if (opts.inputs.length === 0) {
    throw new Error("pdf-edit/merge: at least one input is required");
  }
  const out = await PDFDocument.create();
  for (const buf of opts.inputs) {
    const src = await PDFDocument.load(buf, { updateMetadata: false });
    const indices = Array.from({ length: src.getPageCount() }, (_, i) => i);
    const copied = await out.copyPages(src, indices);
    for (const page of copied) out.addPage(page);
  }
  return out.save();
};
