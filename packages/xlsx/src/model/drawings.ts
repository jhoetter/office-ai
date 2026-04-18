import type { NodeId } from "@officeai/core";

/**
 * Free-floating image attached to a worksheet.
 *
 * Modeled after Excel's `<xdr:twoCellAnchor>` / `<xdr:oneCellAnchor>`
 * payloads (see ECMA-376 §20.5):
 *   - `from-cell + offset` is the top-left anchor expressed as
 *     "row R, column C, plus DX/DY pixels into that cell".
 *   - Width/height are stored in CSS pixels (the renderer's native
 *     coordinate system) and converted to EMUs at serialize time.
 *   - `editAs` mirrors Excel's drag-handle UI:
 *       - `"oneCell"`  — Move with cells, do not size with cells (default).
 *       - `"twoCell"`  — Move and size with cells (we read it but treat
 *                        like oneCell at runtime in v1).
 *       - `"absolute"` — Don't move or size with cells.
 *
 * v1 only authors `editAs = "oneCell"`; existing OOXML files load with
 * their original `editAs` preserved so a re-save doesn't downgrade
 * them.
 */
export interface ImageAnchor {
  /** 0-based row of the top-left anchor. */
  readonly fromRow: number;
  /** 0-based column of the top-left anchor. */
  readonly fromCol: number;
  /** Horizontal pixel offset INTO `fromCol`. */
  readonly fromOffsetXPx: number;
  /** Vertical pixel offset INTO `fromRow`. */
  readonly fromOffsetYPx: number;
  /** Rendered width in CSS pixels. */
  readonly widthPx: number;
  /** Rendered height in CSS pixels. */
  readonly heightPx: number;
  readonly editAs: "oneCell" | "twoCell" | "absolute";
}

/** A media blob keyed by its OOXML container path (`xl/media/imageN.png`). */
export interface ImageBlob {
  readonly partPath: string;
  readonly bytes: Uint8Array;
  readonly contentType: ImageContentType;
  /** SHA-256 hex of `bytes`, used to dedupe identical media. */
  readonly hash: string;
}

export type ImageContentType = "image/png" | "image/jpeg" | "image/gif";

/** Image instance on a sheet, pointing at a media blob in `XlsxWorkbook.images`. */
export interface SheetImage {
  readonly id: NodeId;
  readonly anchor: ImageAnchor;
  /** Key into `XlsxWorkbook.images` (the media part path). */
  readonly mediaRef: string;
  /** Optional friendly name shown in Excel's selection pane. */
  readonly name?: string;
  readonly altText?: string;
}

export const EXTENSION_BY_CONTENT_TYPE: Readonly<Record<ImageContentType, string>> = {
  "image/png": "png",
  "image/jpeg": "jpeg",
  "image/gif": "gif",
};

export function contentTypeForExtension(ext: string): ImageContentType | undefined {
  const lower = ext.toLowerCase();
  if (lower === "png") return "image/png";
  if (lower === "jpg" || lower === "jpeg") return "image/jpeg";
  if (lower === "gif") return "image/gif";
  return undefined;
}
