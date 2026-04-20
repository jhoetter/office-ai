import { indexActionsById, indexActionsByCommandType, type ActionDescriptor } from "@officeai/core";
import { pdfActions } from "./catalogue.js";

export { pdfActions } from "./catalogue.js";

export const pdfActionsById: ReadonlyMap<string, ActionDescriptor> = indexActionsById(pdfActions);

export const pdfActionsByCommandType: ReadonlyMap<string, ReadonlyArray<ActionDescriptor>> =
  indexActionsByCommandType(pdfActions);
