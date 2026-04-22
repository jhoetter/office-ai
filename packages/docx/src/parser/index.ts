export { parseDocx } from "./parse.js";
export { DocxParseError } from "./errors.js";
export { parseFootnotesPart } from "./footnotes.js";
export {
  captureOpaque,
  opaqueToEntry,
  elementEntries,
  findElementEntry,
  rootEntry,
  attrOf,
  readText,
  ATTR_KEY,
  ATTR_PREFIX,
} from "./xml-helpers.js";
