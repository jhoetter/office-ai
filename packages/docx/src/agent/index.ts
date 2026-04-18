export {
  DocxAgent,
  type DocxAgentOptions,
  type DocxRangeRequest,
  type DocxRangeSnapshot,
  type DocxSearchSpec,
  type DocxSearchResult,
} from "./agent.js";
export { snapshotToMarkdown } from "./markdown.js";
export { diffDocxSnapshots } from "./diff.js";
export { resolveEffectiveRpr, resolveEffectivePpr } from "./style-resolver.js";
