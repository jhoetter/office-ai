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
  CropPicturePayload,
  DeleteShapePayload,
  DeleteSlidePayload,
  DistributeShapesPayload,
  DuplicateSlidePayload,
  FormatTextPayload,
  InsertImagePayload,
  InsertTablePayload,
  LayoutKindPayload,
  MoveSlidePayload,
  PptxTextRange,
  SetConnectorEndpointPayload,
  SetConnectorStylePayload,
  SetPositionPayload,
  SetShapeFillPayload,
  ReorderShapePayload,
  ReorderShapeMode,
  SetParagraphAlignmentPayload,
  SetSizePayload,
  SetSlideLayoutPayload,
  SetSlideNotesPayload,
  SetTextAnchorPayload,
  SetTextPayload,
  ShapePreset,
  TextAnchor,
  TextFormatPayload,
} from "./commands/payloads.js";
export { BUILTIN_LAYOUTS, type BuiltinLayout } from "./layouts/builtin.js";
export { resolveEndpoint as resolveConnectorEndpoint } from "./model/connector-geometry.js";
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
