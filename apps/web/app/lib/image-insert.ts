import type { DocxAgent } from "@officeai/docx";
import type { EditorState } from "prosemirror-state";
import { currentParagraphIndex } from "./format-helpers";

/**
 * Common pipeline for "user picked / dropped / pasted an image":
 *
 *  1. Validate MIME and reject anything we can't write into OOXML.
 *  2. Read the file's bytes once.
 *  3. Decode the bytes into an `<img>` element so we can read the
 *     intrinsic pixel dimensions (Word stores image size in EMUs and
 *     uses the file's intrinsic 1× size by default; matching that
 *     produces the same visual result as Word's "Insert Picture").
 *  4. Compute the target paragraph index from the current PM caret.
 *  5. Dispatch `docx:insert-image`.
 *
 * Throws plain `Error` on any validation / decode failure so the caller
 * can route the message into a toast.
 */
export const SUPPORTED_IMAGE_MIME: ReadonlySet<string> = new Set([
  "image/png",
  "image/jpeg",
  "image/jpg",
  "image/gif",
  "image/bmp",
  "image/webp",
  "image/svg+xml",
]);

export interface InsertImageOptions {
  /**
   * Hard cap on the displayed width in CSS pixels. If the image's
   * intrinsic width exceeds this, both width and height are scaled down
   * proportionally — it's the same "fit to page" behaviour Word
   * applies on Insert Picture so the freshly-inserted image doesn't
   * blow past the page margins.
   */
  maxWidthPx?: number;
  /**
   * Override the auto-detected paragraph target. When undefined we use
   * the paragraph the caret is currently in, falling back to the last
   * paragraph if no editor state is provided.
   */
  paragraphIndex?: number;
}

export async function insertImageIntoDocx(
  agent: DocxAgent,
  file: File,
  state: EditorState | null,
  opts: InsertImageOptions = {}
): Promise<void> {
  const mime = (file.type || guessMimeFromName(file.name)).toLowerCase();
  if (!SUPPORTED_IMAGE_MIME.has(mime)) {
    throw new Error(`Unsupported image type "${mime || "unknown"}". Use PNG, JPEG, GIF, BMP, WEBP, or SVG.`);
  }
  const buf = await file.arrayBuffer();
  const bytes = new Uint8Array(buf);
  if (bytes.byteLength === 0) {
    throw new Error("Selected image is empty.");
  }

  const intrinsic = await readIntrinsicSize(bytes, mime);
  const cap = opts.maxWidthPx ?? 600;
  let { width, height } = intrinsic;
  if (width > cap) {
    height = Math.round((height / width) * cap);
    width = cap;
  }
  if (width <= 0 || height <= 0) {
    width = Math.max(1, Math.round(intrinsic.width || 1));
    height = Math.max(1, Math.round(intrinsic.height || 1));
  }

  const snap = agent.getSnapshot();
  const lastParagraphIndex = lastParagraphBodyIndex(snap.root.body);
  const paragraphIndex =
    opts.paragraphIndex !== undefined
      ? opts.paragraphIndex
      : state
        ? clampToParagraph(snap.root.body, currentParagraphIndex(state))
        : lastParagraphIndex;
  if (paragraphIndex < 0) {
    throw new Error("Document has no paragraph to insert into.");
  }

  await agent.applyCommand({
    type: "docx:insert-image",
    payload: {
      at: { paragraph: paragraphIndex, run: 0, offset: 0 },
      data: bytes,
      mimeType: mime,
      width,
      height,
      name: file.name || `Picture`,
      altText: file.name || undefined,
    },
    source: "human",
  });
}

function lastParagraphBodyIndex(body: ReadonlyArray<{ kind: string }>): number {
  for (let i = body.length - 1; i >= 0; i--) {
    if (body[i].kind === "paragraph") return i;
  }
  return -1;
}

function clampToParagraph(body: ReadonlyArray<{ kind: string }>, candidateIndex: number): number {
  if (candidateIndex >= 0 && candidateIndex < body.length && body[candidateIndex]?.kind === "paragraph") {
    return candidateIndex;
  }
  // Walk backwards from the candidate to find the nearest paragraph.
  for (let i = Math.min(candidateIndex, body.length - 1); i >= 0; i--) {
    if (body[i].kind === "paragraph") return i;
  }
  return lastParagraphBodyIndex(body);
}

function guessMimeFromName(name: string): string {
  const ext = name.toLowerCase().split(".").pop() ?? "";
  switch (ext) {
    case "png":
      return "image/png";
    case "jpg":
    case "jpeg":
      return "image/jpeg";
    case "gif":
      return "image/gif";
    case "bmp":
      return "image/bmp";
    case "webp":
      return "image/webp";
    case "svg":
      return "image/svg+xml";
    default:
      return "";
  }
}

async function readIntrinsicSize(
  bytes: Uint8Array,
  mime: string
): Promise<{ width: number; height: number }> {
  // Browser path: decode via createImageBitmap when available (faster
  // and works for SVG-via-blob with explicit MIME). Fall back to an
  // `<img>` element so we still get a result on older Safari.
  if (typeof window === "undefined") {
    return { width: 0, height: 0 };
  }
  const blob = new Blob([bytes as BlobPart], { type: mime });
  if (typeof createImageBitmap === "function" && mime !== "image/svg+xml") {
    try {
      const bitmap = await createImageBitmap(blob);
      const w = bitmap.width;
      const h = bitmap.height;
      bitmap.close?.();
      return { width: w, height: h };
    } catch {
      // fall through to the <img> path
    }
  }
  return await readWithImgElement(blob);
}

function readWithImgElement(blob: Blob): Promise<{ width: number; height: number }> {
  return new Promise((resolve, reject) => {
    const url = URL.createObjectURL(blob);
    const img = new Image();
    img.onload = () => {
      const out = { width: img.naturalWidth, height: img.naturalHeight };
      URL.revokeObjectURL(url);
      resolve(out);
    };
    img.onerror = () => {
      URL.revokeObjectURL(url);
      reject(new Error("Could not decode image to read intrinsic size."));
    };
    img.src = url;
  });
}
