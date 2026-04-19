import type { CommandHandler } from "@officeai/core";
import type { PptxSnapshot } from "../model/types.js";
import { addShapeHandler } from "./add-shape.js";
import { addSlideHandler } from "./add-slide.js";
import { addTextBoxHandler } from "./add-text-box.js";
import { alignShapesHandler, distributeShapesHandler } from "./align-shapes.js";
import { deleteShapeHandler } from "./delete-shape.js";
import { deleteSlideHandler } from "./delete-slide.js";
import { setShapeFillHandler } from "./set-shape-fill.js";
import { reorderShapeHandler } from "./reorder-shape.js";
import { duplicateShapeHandler } from "./duplicate-shape.js";
import { groupShapesHandler, ungroupShapeHandler } from "./group-shapes.js";
import { duplicateSlideHandler } from "./duplicate-slide.js";
import { formatTextHandler } from "./format-text.js";
import { insertImageHandler } from "./insert-image.js";
import { replacePictureMediaHandler } from "./replace-picture-media.js";
import { moveSlideHandler } from "./move-slide.js";
import { setParagraphAlignmentHandler } from "./set-paragraph-alignment.js";
import { setPositionHandler } from "./set-position.js";
import { setSizeHandler } from "./set-size.js";
import { setTextAnchorHandler } from "./set-text-anchor.js";
import { setTextHandler } from "./set-text.js";
import {
  tableAddColumnHandler,
  tableAddRowHandler,
  tableDeleteColumnHandler,
  tableDeleteRowHandler,
  tableSetCellTextHandler,
} from "./table-commands.js";
import { setChartDataHandler, setChartTitleHandler, setChartTypeHandler } from "./chart-commands.js";
import {
  addShapeAnimationHandler,
  removeShapeAnimationHandler,
  reorderShapeAnimationsHandler,
  setSlideTransitionHandler,
} from "./animation-commands.js";
import {
  addConnectorHandler,
  rerouteConnectorHandler,
  setConnectorEndpointHandler,
  setConnectorStyleHandler,
  setConnectorWaypointHandler,
  swapConnectorDirectionHandler,
} from "./connector-commands.js";
import { setSlideLayoutHandler } from "./set-slide-layout.js";
import { setSlideNotesHandler } from "./set-slide-notes.js";
import {
  addCommentHandler,
  deleteCommentHandler,
  editCommentHandler,
  replyCommentHandler,
  resolveCommentHandler,
} from "./comment-commands.js";

export * from "./payloads.js";
export { addShapeHandler } from "./add-shape.js";
export { addSlideHandler } from "./add-slide.js";
export { addTextBoxHandler } from "./add-text-box.js";
export { alignShapesHandler, distributeShapesHandler } from "./align-shapes.js";
export { deleteShapeHandler } from "./delete-shape.js";
export { deleteSlideHandler } from "./delete-slide.js";
export { setShapeFillHandler } from "./set-shape-fill.js";
export { reorderShapeHandler } from "./reorder-shape.js";
export { duplicateShapeHandler } from "./duplicate-shape.js";
export { groupShapesHandler, ungroupShapeHandler } from "./group-shapes.js";
export { duplicateSlideHandler } from "./duplicate-slide.js";
export { formatTextHandler } from "./format-text.js";
export { insertImageHandler } from "./insert-image.js";
export { replacePictureMediaHandler } from "./replace-picture-media.js";
export { moveSlideHandler } from "./move-slide.js";
export { setParagraphAlignmentHandler } from "./set-paragraph-alignment.js";
export { setPositionHandler } from "./set-position.js";
export { setSizeHandler } from "./set-size.js";
export { setTextAnchorHandler } from "./set-text-anchor.js";
export { setTextHandler } from "./set-text.js";
export {
  tableAddColumnHandler,
  tableAddRowHandler,
  tableDeleteColumnHandler,
  tableDeleteRowHandler,
  tableSetCellTextHandler,
} from "./table-commands.js";
export { setChartDataHandler, setChartTitleHandler, setChartTypeHandler } from "./chart-commands.js";
export {
  addShapeAnimationHandler,
  removeShapeAnimationHandler,
  reorderShapeAnimationsHandler,
  setSlideTransitionHandler,
} from "./animation-commands.js";
export {
  addConnectorHandler,
  rerouteConnectorHandler,
  setConnectorEndpointHandler,
  setConnectorStyleHandler,
  setConnectorWaypointHandler,
  swapConnectorDirectionHandler,
} from "./connector-commands.js";
export { setSlideLayoutHandler } from "./set-slide-layout.js";
export { setSlideNotesHandler } from "./set-slide-notes.js";
export {
  addCommentHandler,
  deleteCommentHandler,
  editCommentHandler,
  replyCommentHandler,
  resolveCommentHandler,
} from "./comment-commands.js";

export const allPptxHandlers: ReadonlyArray<CommandHandler<unknown, PptxSnapshot>> = [
  addSlideHandler,
  deleteSlideHandler,
  duplicateSlideHandler,
  moveSlideHandler,
  setTextHandler,
  setPositionHandler,
  setSizeHandler,
  formatTextHandler,
  setParagraphAlignmentHandler,
  setTextAnchorHandler,
  insertImageHandler,
  replacePictureMediaHandler,
  addTextBoxHandler,
  addShapeHandler,
  deleteShapeHandler,
  setShapeFillHandler,
  reorderShapeHandler,
  duplicateShapeHandler,
  groupShapesHandler,
  ungroupShapeHandler,
  alignShapesHandler,
  distributeShapesHandler,
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
  addConnectorHandler,
  setConnectorEndpointHandler,
  setConnectorStyleHandler,
  setConnectorWaypointHandler,
  rerouteConnectorHandler,
  swapConnectorDirectionHandler,
  setSlideLayoutHandler,
  setSlideNotesHandler,
  addCommentHandler,
  replyCommentHandler,
  resolveCommentHandler,
  deleteCommentHandler,
  editCommentHandler,
] as ReadonlyArray<CommandHandler<unknown, PptxSnapshot>>;

export const pptxHandlersById: ReadonlyMap<string, CommandHandler<unknown, PptxSnapshot>> = new Map(
  allPptxHandlers.map((h) => [h.type, h])
);
