export {
  DocxAgent,
  type DocxAgentOptions,
  type DocxRangeRequest,
  type DocxRangeSnapshot,
  type DocxSearchSpec,
  type DocxSearchResult,
} from "./agent.js";
export { snapshotToMarkdown, type SnapshotToMarkdownOptions } from "./markdown.js";
export {
  getPageInfos,
  getPageMarkdown,
  getPagePlainText,
  pageForParagraph,
  type PageInfo,
  type PageTrigger,
} from "./pages.js";
export { diffDocxSnapshots } from "./diff.js";
export {
  resolveEffectiveRpr,
  resolveEffectivePpr,
  resolveThemeFont,
  WORD_DEFAULT_THEME_FONTS,
} from "./style-resolver.js";
export {
  resolveHeaderFooterParts,
  type ResolvedHeaderFooter,
  type ResolvedSectionHeaderFooters,
} from "./header-footer-graph.js";
