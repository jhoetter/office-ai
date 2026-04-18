import type { CommandHandler } from "@officeai/core";
import type { XlsxSnapshot } from "../model/types.js";
import type { InsertColumnPayload } from "./payloads.js";
import { applyStructuralShift } from "./structural-shift.js";

/**
 * `xlsx:insert-column` — insert `count` blank columns to the left of
 * column `at` (1-based, A=1) on `sheet`. Spec: `spec/xlsx/agent-commands.md` §6.
 */
export const insertColumnHandler: CommandHandler<InsertColumnPayload, XlsxSnapshot> = {
  type: "xlsx:insert-column",
  apply(snapshot, payload) {
    return applyStructuralShift(snapshot, payload, "column", "insert");
  },
};
