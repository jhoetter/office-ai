export type XlsxSerializeErrorCode =
  | "container-failed"
  | "workbook-failed"
  | "sheet-failed"
  | "shared-strings-failed"
  | "styles-failed"
  | "comments-failed"
  | "rels-failed";

export class XlsxSerializeError extends Error {
  readonly code: XlsxSerializeErrorCode;
  readonly partPath?: string;
  constructor(code: XlsxSerializeErrorCode, message: string, opts?: { partPath?: string; cause?: unknown }) {
    super(message);
    this.name = "XlsxSerializeError";
    this.code = code;
    this.partPath = opts?.partPath;
    if (opts?.cause !== undefined) {
      (this as unknown as { cause: unknown }).cause = opts.cause;
    }
  }
}
