import { type CommandHandler } from "@officeai/core";
import type { DocxSnapshot } from "../model/types.js";
import { resolveTrackedChange } from "./accept-change.js";
import type { RejectChangePayload } from "./payloads.js";

/**
 * Inverse of `docx:accept-change`.
 *
 *   - `<w:ins>` reject: drop the wrapper AND its children (the proposed
 *     insertion never lands).
 *   - `<w:del>` reject: drop the wrapper, keep its children (the deletion
 *     is undone, the original text stays).
 *
 * Implementation is a one-line delegation to the shared resolver in
 * `accept-change.ts`. Errors and the "no `RevisionWrapper` for this id
 * survives a round-trip" invariant are identical.
 */
export const rejectChangeHandler: CommandHandler<RejectChangePayload, DocxSnapshot> = {
  type: "docx:reject-change",
  apply(snapshot, payload) {
    return resolveTrackedChange(snapshot, payload.revisionId, "reject");
  },
};
