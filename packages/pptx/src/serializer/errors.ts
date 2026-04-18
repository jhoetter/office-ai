export class PptxSerializeError extends Error {
  readonly code: string;
  readonly partPath?: string;
  constructor(code: string, message: string, opts?: { partPath?: string; cause?: unknown }) {
    super(message);
    this.name = "PptxSerializeError";
    this.code = code;
    this.partPath = opts?.partPath;
    if (opts?.cause !== undefined) {
      (this as unknown as { cause: unknown }).cause = opts.cause;
    }
  }
}
