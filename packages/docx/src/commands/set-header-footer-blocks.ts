import { CommandError, type CommandHandler } from "@officeai/core";
import type { BlockNode, DocxSnapshot, HeaderFooterPart } from "../model/types.js";
import { buildDiff, evolveSnapshot } from "./helpers.js";
import { mergeHeaderFooterDirty } from "./set-header-text.js";
import type { SetHeaderFooterBlocksPayload } from "./payloads.js";

/**
 * Replace the entire `body` of a header or footer part while
 * preserving the part's `partPath`, `kind`, `target`, and
 * `rootAttrs` (so namespace declarations and the part's identity in
 * the OOXML package survive the rewrite).
 *
 * The thin-wrapper alternatives `docx:set-header-text` and
 * `docx:set-footer-text` flatten the body to one paragraph + one
 * run; this command lets the editor commit a rich body
 * (multi-paragraph, page-number fields, inline images) without
 * destroying tokens that the part already carries. Used by the
 * Word-style in-place header/footer authoring path in the page
 * decorations plugin.
 *
 * Errors:
 * - `unknown-target` when no part matches `partPath`.
 * - `invalid-payload` when `body` is not a non-empty array.
 */
export const setHeaderFooterBlocksHandler: CommandHandler<SetHeaderFooterBlocksPayload, DocxSnapshot> = {
  type: "docx:set-header-footer-blocks",
  apply(snapshot, payload) {
    const { partPath, body } = payload;
    if (typeof partPath !== "string" || partPath.length === 0) {
      throw new CommandError("invalid-payload", "partPath is required");
    }
    if (!Array.isArray(body) || body.length === 0) {
      throw new CommandError(
        "invalid-payload",
        "body must be a non-empty BlockNode[] (use a single empty paragraph to clear the part)"
      );
    }
    const partIdx = snapshot.root.headersAndFooters.findIndex((p) => p.partPath === partPath);
    if (partIdx < 0) {
      throw new CommandError(
        "unknown-target",
        `no header/footer part with path "${partPath}" (looked across ${snapshot.root.headersAndFooters.length} parts)`
      );
    }
    const part = snapshot.root.headersAndFooters[partIdx];
    const updatedPart: HeaderFooterPart = { ...part, body: body as ReadonlyArray<BlockNode> };
    const newParts = snapshot.root.headersAndFooters.slice();
    newParts[partIdx] = updatedPart;

    const dirty = mergeHeaderFooterDirty(snapshot.dirty, part.partPath);
    const next = evolveSnapshot(
      snapshot,
      { ...snapshot.root, headersAndFooters: newParts },
      { headersAndFooters: dirty }
    );

    return {
      next,
      diff: buildDiff(snapshot.revision, next.revision, {
        kind: "node-updated",
        nodeId: part.id,
        path: ["headersAndFooters", partIdx, "body"],
        field: "body",
        summary: `${part.kind} ${part.partPath} body := ${body.length} block(s)`,
      }),
    };
  },
};
