import type { CommandHandler } from "@officeai/core";
import type { PptxSnapshot } from "../model/types.js";

/**
 * Registry of all PPTX command handlers. Empty in P4; populated in P5–P6.
 */
export const allPptxHandlers: ReadonlyArray<CommandHandler<unknown, PptxSnapshot>> = [];
