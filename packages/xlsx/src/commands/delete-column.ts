import type { CommandHandler } from "@officeai/core";
import type { XlsxSnapshot } from "../model/types.js";
import type { DeleteColumnPayload } from "./payloads.js";
import { applyStructuralShift } from "./structural-shift.js";

/**
 * `xlsx:delete-column` — drop columns `at..at+count-1` (1-based) on
 * `sheet`, shifting everything to the right left. Spec:
 * `spec/xlsx/agent-commands.md` §8.
 */
export const deleteColumnHandler: CommandHandler<DeleteColumnPayload, XlsxSnapshot> = {
  type: "xlsx:delete-column",
  apply(snapshot, payload) {
    return applyStructuralShift(snapshot, payload, "column", "delete");
  },
};
