import type { CommandHandler } from "@officeai/core";
import type { XlsxSnapshot } from "../model/types.js";
import type { DeleteRowPayload } from "./payloads.js";
import { applyStructuralShift } from "./structural-shift.js";

/**
 * `xlsx:delete-row` — drop rows `at..at+count-1` (1-based) on `sheet`,
 * shifting everything below up. Spec: `spec/xlsx/agent-commands.md` §7.
 */
export const deleteRowHandler: CommandHandler<DeleteRowPayload, XlsxSnapshot> = {
  type: "xlsx:delete-row",
  apply(snapshot, payload) {
    return applyStructuralShift(snapshot, payload, "row", "delete");
  },
};
