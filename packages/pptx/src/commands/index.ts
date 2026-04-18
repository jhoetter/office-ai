import type { CommandHandler } from "@officeai/core";
import type { PptxSnapshot } from "../model/types.js";
import { addSlideHandler } from "./add-slide.js";
import { deleteSlideHandler } from "./delete-slide.js";
import { duplicateSlideHandler } from "./duplicate-slide.js";
import { moveSlideHandler } from "./move-slide.js";
import { setPositionHandler } from "./set-position.js";
import { setSizeHandler } from "./set-size.js";
import { setTextHandler } from "./set-text.js";

export * from "./payloads.js";
export { addSlideHandler } from "./add-slide.js";
export { deleteSlideHandler } from "./delete-slide.js";
export { duplicateSlideHandler } from "./duplicate-slide.js";
export { moveSlideHandler } from "./move-slide.js";
export { setPositionHandler } from "./set-position.js";
export { setSizeHandler } from "./set-size.js";
export { setTextHandler } from "./set-text.js";

export const allPptxHandlers: ReadonlyArray<CommandHandler<unknown, PptxSnapshot>> = [
  addSlideHandler,
  deleteSlideHandler,
  duplicateSlideHandler,
  moveSlideHandler,
  setTextHandler,
  setPositionHandler,
  setSizeHandler,
] as ReadonlyArray<CommandHandler<unknown, PptxSnapshot>>;

export const pptxHandlersById: ReadonlyMap<string, CommandHandler<unknown, PptxSnapshot>> =
  new Map(allPptxHandlers.map((h) => [h.type, h]));
