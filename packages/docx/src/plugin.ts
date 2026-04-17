import type { FormatPlugin } from "@officeai/core";
import { allDocxHandlers } from "./commands/index.js";
import type { DocxSnapshot } from "./model/types.js";
import { parseDocx } from "./parser/index.js";
import { serializeDocx } from "./serializer/index.js";
import { snapshotToMarkdown } from "./agent/markdown.js";

export const docxPlugin: FormatPlugin<DocxSnapshot> = {
  format: "docx",
  handlers: allDocxHandlers,
  async parse(buffer) {
    return parseDocx(buffer);
  },
  async serialize(snapshot) {
    return serializeDocx(snapshot);
  },
  toMarkdown(snapshot) {
    return snapshotToMarkdown(snapshot);
  },
};
