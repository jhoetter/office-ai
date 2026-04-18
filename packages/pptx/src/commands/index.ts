import type { CommandHandler } from "@officeai/core";
import type { PptxSnapshot } from "../model/types.js";
import { addSlideHandler } from "./add-slide.js";
import { addTextBoxHandler } from "./add-text-box.js";
import { deleteSlideHandler } from "./delete-slide.js";
import { duplicateSlideHandler } from "./duplicate-slide.js";
import { formatTextHandler } from "./format-text.js";
import { insertImageHandler } from "./insert-image.js";
import { moveSlideHandler } from "./move-slide.js";
import { setPositionHandler } from "./set-position.js";
import { setSizeHandler } from "./set-size.js";
import { setTextHandler } from "./set-text.js";
import {
  tableAddColumnHandler,
  tableAddRowHandler,
  tableDeleteColumnHandler,
  tableDeleteRowHandler,
  tableSetCellTextHandler,
} from "./table-commands.js";
import {
  setChartDataHandler,
  setChartTitleHandler,
  setChartTypeHandler,
} from "./chart-commands.js";
import {
  addShapeAnimationHandler,
  removeShapeAnimationHandler,
  reorderShapeAnimationsHandler,
  setSlideTransitionHandler,
} from "./animation-commands.js";

export * from "./payloads.js";
export { addSlideHandler } from "./add-slide.js";
export { addTextBoxHandler } from "./add-text-box.js";
export { deleteSlideHandler } from "./delete-slide.js";
export { duplicateSlideHandler } from "./duplicate-slide.js";
export { formatTextHandler } from "./format-text.js";
export { insertImageHandler } from "./insert-image.js";
export { moveSlideHandler } from "./move-slide.js";
export { setPositionHandler } from "./set-position.js";
export { setSizeHandler } from "./set-size.js";
export { setTextHandler } from "./set-text.js";
export {
  tableAddColumnHandler,
  tableAddRowHandler,
  tableDeleteColumnHandler,
  tableDeleteRowHandler,
  tableSetCellTextHandler,
} from "./table-commands.js";
export {
  setChartDataHandler,
  setChartTitleHandler,
  setChartTypeHandler,
} from "./chart-commands.js";
export {
  addShapeAnimationHandler,
  removeShapeAnimationHandler,
  reorderShapeAnimationsHandler,
  setSlideTransitionHandler,
} from "./animation-commands.js";

export const allPptxHandlers: ReadonlyArray<CommandHandler<unknown, PptxSnapshot>> = [
  addSlideHandler,
  deleteSlideHandler,
  duplicateSlideHandler,
  moveSlideHandler,
  setTextHandler,
  setPositionHandler,
  setSizeHandler,
  formatTextHandler,
  insertImageHandler,
  addTextBoxHandler,
  tableSetCellTextHandler,
  tableAddRowHandler,
  tableDeleteRowHandler,
  tableAddColumnHandler,
  tableDeleteColumnHandler,
  setChartTitleHandler,
  setChartDataHandler,
  setChartTypeHandler,
  setSlideTransitionHandler,
  addShapeAnimationHandler,
  removeShapeAnimationHandler,
  reorderShapeAnimationsHandler,
] as ReadonlyArray<CommandHandler<unknown, PptxSnapshot>>;

export const pptxHandlersById: ReadonlyMap<string, CommandHandler<unknown, PptxSnapshot>> =
  new Map(allPptxHandlers.map((h) => [h.type, h]));
