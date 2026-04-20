export class PdfSerializeError extends Error {
  constructor(
    message: string,
    readonly cause?: unknown
  ) {
    super(message);
    this.name = "PdfSerializeError";
  }
}
