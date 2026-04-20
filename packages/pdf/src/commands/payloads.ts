import type { NodeId } from "@officeai/core";
import type { PdfRect, PdfRotation } from "../model/types.js";

export const PDF_COMMAND_TYPES = [
  "pdf:rotate-pages",
  "pdf:set-page-rotation",
  "pdf:reorder-pages",
  "pdf:delete-pages",
  "pdf:set-metadata",
  "pdf:add-bookmark",
  "pdf:add-comment",
  "pdf:reply-comment",
  "pdf:edit-comment",
  "pdf:resolve-comment",
  "pdf:delete-comment",
] as const;

export type PdfCommandType = (typeof PDF_COMMAND_TYPES)[number];

export interface RotatePagesPayload {
  /** 1-indexed page numbers. */
  readonly pages: ReadonlyArray<number>;
  /** Delta in degrees, must be a multiple of 90. */
  readonly delta: 90 | 180 | 270 | -90 | -180 | -270;
}

export interface SetPageRotationPayload {
  readonly pageNumber: number;
  readonly rotation: PdfRotation;
}

export interface ReorderPagesPayload {
  /**
   * New page order, expressed as a permutation of 1..N. Length MUST equal
   * the document page count and be a valid permutation.
   */
  readonly order: ReadonlyArray<number>;
}

export interface DeletePagesPayload {
  /** 1-indexed page numbers to remove. */
  readonly pages: ReadonlyArray<number>;
}

export interface SetMetadataPayload {
  readonly title?: string;
  readonly author?: string;
  readonly subject?: string;
  readonly keywords?: string;
  readonly creator?: string;
  readonly producer?: string;
}

export interface AddBookmarkPayload {
  readonly title: string;
  readonly pageNumber: number;
  readonly parentId?: NodeId;
}

export interface AddCommentPayload {
  readonly id?: NodeId;
  readonly author: string;
  readonly text: string;
  readonly pageNumber: number;
  readonly normalizedRect: PdfRect;
}

export interface ReplyCommentPayload {
  readonly id?: NodeId;
  readonly parentId: NodeId;
  readonly author: string;
  readonly text: string;
}

export interface EditCommentPayload {
  readonly commentId: NodeId;
  readonly text: string;
}

export interface ResolveCommentPayload {
  readonly commentId: NodeId;
  readonly resolved: boolean;
}

export interface DeleteCommentPayload {
  readonly commentId: NodeId;
}
