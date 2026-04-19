/**
 * Client-side helper around `POST /api/convert`.
 *
 * The three editors call this to turn their native binary into PDF
 * or HTML via LibreOffice on the server. Keep the surface narrow on
 * purpose — each product owns its own filename / options dialog and
 * just hands us the bytes.
 */

export type ConvertSourceExt = "docx" | "xlsx" | "pptx";
export type ConvertTargetExt = "pdf" | "html";

export interface ConvertViaServerArgs {
  /** The source document bytes. */
  readonly bytes: Uint8Array | ArrayBuffer;
  /** The format of `bytes`. */
  readonly sourceExt: ConvertSourceExt;
  /** The format the server should produce. */
  readonly targetExt: ConvertTargetExt;
  /** Base filename (no extension). The server uses it for the
   * Content-Disposition header — we use it for our own `<a download>`
   * fallback when the response doesn't surface it back. */
  readonly filename: string;
  /** MIME type of the source bytes (used to construct the multipart
   * Blob). Optional — we default to the right OOXML MIME for each
   * source extension. */
  readonly sourceMime?: string;
  /** Optional abort signal for cancellation. */
  readonly signal?: AbortSignal;
}

const SOURCE_MIME: Record<ConvertSourceExt, string> = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
};

export class ConvertError extends Error {
  readonly status: number;
  constructor(message: string, status: number) {
    super(message);
    this.name = "ConvertError";
    this.status = status;
  }
}

/**
 * POST the bytes to `/api/convert` and return the converted Blob.
 * Throws `ConvertError` on non-2xx responses with the server message
 * if available.
 */
export async function convertViaServer(args: ConvertViaServerArgs): Promise<Blob> {
  const sourceMime = args.sourceMime ?? SOURCE_MIME[args.sourceExt];
  const buffer =
    args.bytes instanceof Uint8Array
      ? new Uint8Array(args.bytes)
      : new Uint8Array(args.bytes);
  const sourceBlob = new Blob([buffer as BlobPart], { type: sourceMime });

  const form = new FormData();
  form.append("file", sourceBlob, `input.${args.sourceExt}`);
  form.append("sourceExt", args.sourceExt);
  form.append("targetExt", args.targetExt);
  form.append("filename", args.filename);

  const response = await fetch("/api/convert", {
    method: "POST",
    body: form,
    signal: args.signal,
  });

  if (!response.ok) {
    let message = `Conversion failed (${response.status}).`;
    try {
      const data = (await response.json()) as { message?: string };
      if (data?.message) message = data.message;
    } catch {
      // Body isn't JSON — fall back to the generic message.
    }
    throw new ConvertError(message, response.status);
  }

  return await response.blob();
}
