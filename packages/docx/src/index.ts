export * from "./model/index.js";
export { parseDocx, DocxParseError } from "./parser/index.js";
export { serializeDocx, DocxSerializeError } from "./serializer/index.js";
export {
  DocxAgent,
  resolveEffectiveRpr,
  resolveEffectivePpr,
  resolveThemeFont,
  WORD_DEFAULT_THEME_FONTS,
  resolveHeaderFooterParts,
  snapshotToMarkdown,
  getPageInfos,
  getPageMarkdown,
  getPagePlainText,
  pageForParagraph,
  type ResolvedHeaderFooter,
  type ResolvedSectionHeaderFooters,
  type SnapshotToMarkdownOptions,
  type PageInfo,
  type PageTrigger,
} from "./agent/index.js";
export { docxPlugin } from "./plugin.js";
export * from "./commands/index.js";
export {
  docxSchema,
  docToPM,
  transactionToCommands,
  mountDocxEditor,
  chunkIntoPages,
  documentPageGeometry,
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
