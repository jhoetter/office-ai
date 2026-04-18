export * from "./model/types.js";
export { parsePptx, type ParseOptions } from "./parser/parse.js";
export { PptxParseError } from "./parser/errors.js";
export { serializePptx } from "./serializer/serialize.js";
export { PptxSerializeError } from "./serializer/errors.js";
export { allPptxHandlers } from "./commands/index.js";
export type {
  AddShapePayload,
  AddSlidePayload,
  AddTextBoxPayload,
  DeleteShapePayload,
  DeleteSlidePayload,
  DuplicateSlidePayload,
  FormatTextPayload,
  InsertImagePayload,
  MoveSlidePayload,
  PptxTextRange,
  SetPositionPayload,
  SetShapeFillPayload,
  SetSizePayload,
  SetTextPayload,
  ShapePreset,
  TextFormatPayload,
} from "./commands/payloads.js";
export {
  PptxAgent,
  type PptxAgentOptions,
  type PptxRangeRequest,
  type PptxRangeSnapshot,
  type PptxSearchSpec,
  type PptxSearchResult,
  snapshotToMarkdown,
} from "./agent/index.js";
