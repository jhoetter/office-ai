export class DocxSerializeError extends Error {
  readonly code: string;
  constructor(code: string, message: string, opts?: { cause?: unknown }) {
    super(message);
    this.name = "DocxSerializeError";
    this.code = code;
    if (opts?.cause !== undefined) (this as unknown as { cause: unknown }).cause = opts.cause;
  }
}
