import { indexActionsById, indexActionsByCommandType, type ActionDescriptor } from "@officeai/core";
import { docxActions } from "./catalogue.js";

export { docxActions } from "./catalogue.js";
export type { DocxActionSection } from "./catalogue.js";

/** O(1) lookup of a docx action descriptor by its `id`. */
export const docxActionsById: ReadonlyMap<string, ActionDescriptor> = indexActionsById(docxActions);

/**
 * Index from bus command type → catalogue entries. Multiple
 * descriptors may share a `commandType` (e.g. one CLI subcommand
 * plus a hidden palette wrapper). Used by the parity check.
 */
export const docxActionsByCommandType: ReadonlyMap<string, ReadonlyArray<ActionDescriptor>> =
  indexActionsByCommandType(docxActions);
