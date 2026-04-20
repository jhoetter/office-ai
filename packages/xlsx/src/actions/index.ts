import { indexActionsById, indexActionsByCommandType, type ActionDescriptor } from "@officeai/core";
import { xlsxActions } from "./catalogue.js";

export { xlsxActions } from "./catalogue.js";

export const xlsxActionsById: ReadonlyMap<string, ActionDescriptor> = indexActionsById(xlsxActions);

export const xlsxActionsByCommandType: ReadonlyMap<string, ReadonlyArray<ActionDescriptor>> =
  indexActionsByCommandType(xlsxActions);
