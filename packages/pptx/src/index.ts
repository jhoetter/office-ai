export * from "./model/types.js";
export { parsePptx, type ParseOptions } from "./parser/parse.js";
export { PptxParseError } from "./parser/errors.js";
export { serializePptx } from "./serializer/serialize.js";
export { PptxSerializeError } from "./serializer/errors.js";
export { allPptxHandlers } from "./commands/index.js";
export type {
  AddConnectorPayload,
  AddShapePayload,
  AddSlidePayload,
  AddTextBoxPayload,
  AddCommentPayload,
  AlignMode,
  AlignShapesPayload,
  ConnectorEndShapePayload,
  DeleteCommentPayload,
  EditCommentPayload,
  ReplyCommentPayload,
  ResolveCommentPayload,
  ConnectorEndpointPayload,
  ConnectorTypePayload,
  DeleteShapePayload,
  DeleteSlidePayload,
  DistributeShapesPayload,
  DuplicateSlidePayload,
  FormatTextPayload,
  InsertImagePayload,
  LayoutKindPayload,
  MoveSlidePayload,
  PptxTextRange,
  SetConnectorEndpointPayload,
  SetConnectorStylePayload,
  SetPositionPayload,
  SetShapeFillPayload,
  ReorderShapePayload,
  ReorderShapeMode,
  SetSizePayload,
  SetSlideLayoutPayload,
  SetSlideNotesPayload,
  SetTextPayload,
  ShapePreset,
  TextFormatPayload,
} from "./commands/payloads.js";
export { BUILTIN_LAYOUTS, type BuiltinLayout } from "./layouts/builtin.js";
export { PPTX_COMMAND_TYPES, type PptxCommandType } from "./commands/payloads.js";
export {
  PptxAgent,
  type PptxAgentOptions,
  type PptxRangeRequest,
  type PptxRangeSnapshot,
  type PptxSearchSpec,
  type PptxSearchResult,
  snapshotToMarkdown,
} from "./agent/index.js";
