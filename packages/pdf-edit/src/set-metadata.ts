import { loadPdf } from "./internal.js";

export interface SetMetadataOptions {
  readonly title?: string;
  readonly author?: string;
  readonly subject?: string;
  readonly keywords?: string;
  readonly creator?: string;
  readonly producer?: string;
}

export const setMetadata = async (buffer: Uint8Array, opts: SetMetadataOptions): Promise<Uint8Array> => {
  const pdf = await loadPdf(buffer);
  if (opts.title !== undefined) pdf.setTitle(opts.title);
  if (opts.author !== undefined) pdf.setAuthor(opts.author);
  if (opts.subject !== undefined) pdf.setSubject(opts.subject);
  if (opts.keywords !== undefined) pdf.setKeywords([opts.keywords]);
  if (opts.creator !== undefined) pdf.setCreator(opts.creator);
  if (opts.producer !== undefined) pdf.setProducer(opts.producer);
  return pdf.save();
};
