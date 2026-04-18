export * from "./model/index.js";
export { parseDocx, DocxParseError } from "./parser/index.js";
export { serializeDocx, DocxSerializeError } from "./serializer/index.js";
export {
  DocxAgent,
  resolveEffectiveRpr,
  resolveEffectivePpr,
  resolveHeaderFooterParts,
  type ResolvedHeaderFooter,
  type ResolvedSectionHeaderFooters,
} from "./agent/index.js";
export { docxPlugin } from "./plugin.js";
export * from "./commands/index.js";
export {
  docxSchema,
  docToPM,
  transactionToCommands,
  mountDocxEditor,
  chunkIntoPages,
  geometryFromProperties,
  type MountOptions,
  type MountResult,
  type TranslationOptions,
  type TranslationResult,
  type UnsupportedTx,
  type PageChunk,
  type PageGeometry,
  type Measure,
} from "./renderer/index.js";
