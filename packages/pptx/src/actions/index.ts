import { indexActionsById, indexActionsByCommandType, type ActionDescriptor } from "@officeai/core";
import { pptxActions } from "./catalogue.js";

export { pptxActions } from "./catalogue.js";

export const pptxActionsById: ReadonlyMap<string, ActionDescriptor> = indexActionsById(pptxActions);

export const pptxActionsByCommandType: ReadonlyMap<
  string,
  ReadonlyArray<ActionDescriptor>
> = indexActionsByCommandType(pptxActions);
