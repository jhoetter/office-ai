import { type CommandHandler } from "@officeai/core";
import type { DocxSnapshot } from "../model/types.js";
import { applySetTextToHeaderFooter } from "./set-header-text.js";
import type { SetFooterTextPayload } from "./payloads.js";

/**
 * Mirror of `set-header-text` for footers. Both share `applySetTextToHeaderFooter`
 * since the OOXML shape is identical (only the wrapping tag differs:
 * `w:hdr` vs `w:ftr`).
 */
export const setFooterTextHandler: CommandHandler<SetFooterTextPayload, DocxSnapshot> = {
  type: "docx:set-footer-text",
  apply(snapshot, payload, ctx) {
    return applySetTextToHeaderFooter(snapshot, "footer", payload, ctx.mintNodeId);
  },
};
