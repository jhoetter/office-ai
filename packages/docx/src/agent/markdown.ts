import type { DocxSnapshot, Paragraph } from "../model/types.js";
import { paragraphPlainText } from "../commands/helpers.js";

const HEADING_STYLES: Record<string, number> = {
  Title: 1,
  Heading1: 1,
  Heading2: 2,
  Heading3: 3,
  Heading4: 4,
  Heading5: 5,
  Heading6: 6,
};

export function snapshotToMarkdown(snapshot: DocxSnapshot): string {
  const lines: string[] = [];
  for (const block of snapshot.root.body) {
    if (block.kind === "paragraph") {
      lines.push(paragraphToMarkdown(block));
      lines.push("");
    } else if (block.kind === "table") {
      lines.push("> [table preserved — content omitted in this digest]");
      lines.push("");
    } else if (block.kind === "section-break") {
      lines.push("---");
      lines.push("");
    } else if (block.kind === "opaque-block") {
      lines.push(`> [opaque block: ${block.raw.tag}]`);
      lines.push("");
    }
  }
  return (
    lines
      .join("\n")
      .replace(/\n{3,}/g, "\n\n")
      .trimEnd() + "\n"
  );
}

function paragraphToMarkdown(p: Paragraph): string {
  const text = paragraphPlainText(p);
  const styleId = p.properties.styleId;
  if (styleId && HEADING_STYLES[styleId]) {
    const level = HEADING_STYLES[styleId];
    return `${"#".repeat(level)} ${text}`;
  }
  if (styleId === "ListParagraph" || p.properties.numbering) {
    return `- ${text}`;
  }
  return text;
}
