export type XlsxParseErrorCode =
  | "zip-corruption"
  | "missing-workbook-part"
  | "missing-content-types"
  | "invalid-xml"
  | "invalid-workbook"
  | "missing-sheet-target"
  | "sheetjs-failure";

export class XlsxParseError extends Error {
  readonly code: XlsxParseErrorCode;
  readonly partPath?: string;
  constructor(code: XlsxParseErrorCode, message: string, opts?: { partPath?: string; cause?: unknown }) {
    super(message);
    this.name = "XlsxParseError";
    this.code = code;
    this.partPath = opts?.partPath;
    if (opts?.cause !== undefined) {
      (this as unknown as { cause: unknown }).cause = opts.cause;
    }
  }
}
