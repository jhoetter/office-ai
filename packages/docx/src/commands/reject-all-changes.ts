import type { CommandHandler } from "@officeai/core";
import type { DocxSnapshot } from "../model/types.js";
import { resolveAll } from "./accept-all-changes.js";
import type { RejectAllChangesPayload } from "./payloads.js";

/**
 * B8 — Reject every tracked change in the document in a single
 * transaction; mirror of {@link acceptAllChangesHandler}. Insertions
 * disappear, deletions are restored.
 */
export const rejectAllChangesHandler: CommandHandler<RejectAllChangesPayload, DocxSnapshot> = {
  type: "docx:reject-all-changes",
  apply(snapshot) {
    return resolveAll(snapshot, "reject");
  },
};
