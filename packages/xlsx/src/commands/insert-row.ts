import type { CommandHandler } from "@officeai/core";
import type { XlsxSnapshot } from "../model/types.js";
import type { InsertRowPayload } from "./payloads.js";
import { applyStructuralShift } from "./structural-shift.js";

/**
 * `xlsx:insert-row` — insert `count` blank rows above row `at`
 * (1-based) on `sheet`. Spec: `spec/xlsx/agent-commands.md` §5.
 */
export const insertRowHandler: CommandHandler<InsertRowPayload, XlsxSnapshot> = {
  type: "xlsx:insert-row",
  apply(snapshot, payload) {
    return applyStructuralShift(snapshot, payload, "row", "insert");
  },
};
